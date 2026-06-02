from __future__ import annotations

import argparse
import csv
import json
import os
import re
import subprocess
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Protocol

import pandas as pd
import yaml

from utils import compact_spaces, ensure_dir, is_blank, normalize_header, read_products, write_products


TYPE_COLUMN_CANDIDATES = ["attr_typ", "Typ produktu", "Typ", "Rodzaj produktu"]
DEFAULT_GEO_TARGETS = ["2616"]  # Poland
DEFAULT_LANGUAGE = "1045"  # Polish
DEFAULT_NETWORK = "GOOGLE_SEARCH"
DEFAULT_CACHE_PATH = "cache/keyword_planner_metrics.json"
OUTPUT_COLUMNS = [
    "primary_keyword",
    "primary_keyword_avg_monthly_searches",
    "secondary_keyword",
    "secondary_keyword_avg_monthly_searches",
    "keyword_candidates_json",
    "keyword_status",
    "keyword_source",
]
BRAND_COLUMN_CANDIDATES = ["attr_producent", "Producent", "manufacturer_name", "producer", "Marka"]
CODE_COLUMN_CANDIDATES = ["SKU", "sku", "Kod Producenta", "Kod", "EAN", "ean", "GTIN"]


@dataclass(frozen=True)
class KeywordMetric:
    keyword: str
    avg_monthly_searches: int
    source: str = ""


@dataclass
class KeywordAnalysis:
    product_type: str
    primary_keyword: str = ""
    primary_keyword_avg_monthly_searches: int | str = ""
    secondary_keyword: str = ""
    secondary_keyword_avg_monthly_searches: int | str = ""
    keyword_candidates_json: str = "[]"
    keyword_status: str = "NO_DATA"
    keyword_source: str = ""


class KeywordProvider(Protocol):
    source_name: str

    def generate_ideas(self, product_type: str) -> list[str]:
        ...

    def historical_metrics(self, keywords: list[str]) -> dict[str, KeywordMetric]:
        ...


class OfflineKeywordProvider:
    source_name = "offline_csv"

    def __init__(self, path: str | Path) -> None:
        self.rows: list[dict[str, str]] = []
        with Path(path).open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            for row in reader:
                keyword = compact_spaces(first_existing(row, ["keyword", "Keyword", "słowo", "slowo"]))
                if keyword:
                    self.rows.append(row)

    def generate_ideas(self, product_type: str) -> list[str]:
        product_type_key = normalize_header(product_type)
        exact: list[str] = []
        fuzzy: list[str] = []
        type_tokens = set(product_type_key.split())
        for row in self.rows:
            keyword = compact_spaces(first_existing(row, ["keyword", "Keyword", "słowo", "slowo"]))
            row_type = normalize_header(first_existing(row, ["product_type", "type", "typ", "Typ produktu", "seed"]))
            keyword_key = normalize_header(keyword)
            if row_type and row_type == product_type_key:
                exact.append(keyword)
                continue
            keyword_tokens = set(keyword_key.split())
            if type_tokens and (type_tokens <= keyword_tokens or type_tokens & keyword_tokens):
                fuzzy.append(keyword)
        return unique_preserve_order([*exact, *fuzzy])

    def historical_metrics(self, keywords: list[str]) -> dict[str, KeywordMetric]:
        wanted = {normalize_keyword(keyword): keyword for keyword in keywords}
        result: dict[str, KeywordMetric] = {}
        for row in self.rows:
            keyword = compact_spaces(first_existing(row, ["keyword", "Keyword", "słowo", "slowo"]))
            normalized = normalize_keyword(keyword)
            if normalized not in wanted:
                continue
            searches = parse_int(first_existing(row, ["avg_monthly_searches", "Avg. monthly searches", "searches"]))
            original = wanted[normalized]
            result[normalize_keyword(original)] = KeywordMetric(original, searches, self.source_name)
        return result


class GoogleAdsKeywordProvider:
    source_name = "google_ads_keyword_planner"

    def __init__(
        self,
        customer_id: str,
        google_ads_config: str = "",
        geo_targets: list[str] | None = None,
        language: str = DEFAULT_LANGUAGE,
        network: str = DEFAULT_NETWORK,
    ) -> None:
        self.customer_id = re.sub(r"\D", "", customer_id)
        if not self.customer_id:
            raise ValueError("Brak poprawnego Google Ads customer ID.")
        self.geo_targets = geo_targets or DEFAULT_GEO_TARGETS
        self.language = language
        self.network = network
        self.client = self._load_client(google_ads_config)
        self.service = self.client.get_service("KeywordPlanIdeaService")

    @staticmethod
    def _load_client(config_path: str = ""):
        ensure_grpc_ssl_roots()
        try:
            from google.ads.googleads.client import GoogleAdsClient
        except ImportError as exc:
            raise RuntimeError("Brak biblioteki google-ads. Uruchom: pip install -r requirements.txt") from exc

        if config_path:
            return GoogleAdsClient.load_from_storage(config_path, version="v24")
        if os.environ.get("GOOGLE_ADS_CONFIGURATION_FILE_PATH"):
            return GoogleAdsClient.load_from_storage(version="v24")
        return GoogleAdsClient.load_from_storage(version="v24")

    def generate_ideas(self, product_type: str) -> list[str]:
        request = self.client.get_type("GenerateKeywordIdeasRequest")
        request.customer_id = self.customer_id
        request.language = self.resource_name("language_constant", self.language)
        request.geo_target_constants.extend(self.resource_names("geo_target_constant", self.geo_targets))
        request.keyword_plan_network = self.network_enum()
        request.keyword_seed.keywords.append(product_type)

        ideas: list[str] = []
        for idea in self.service.generate_keyword_ideas(request=request):
            text = compact_spaces(getattr(idea, "text", ""))
            if text:
                ideas.append(text)
        return unique_preserve_order(ideas)

    def historical_metrics(self, keywords: list[str]) -> dict[str, KeywordMetric]:
        if not keywords:
            return {}
        request = self.client.get_type("GenerateKeywordHistoricalMetricsRequest")
        request.customer_id = self.customer_id
        request.keywords.extend(keywords)
        request.language = self.resource_name("language_constant", self.language)
        request.geo_target_constants.extend(self.resource_names("geo_target_constant", self.geo_targets))
        request.keyword_plan_network = self.network_enum()

        response = self.service.generate_keyword_historical_metrics(request=request)
        result: dict[str, KeywordMetric] = {}
        for item in response.results:
            text = compact_spaces(getattr(item, "text", ""))
            metrics = getattr(item, "keyword_metrics", None)
            searches = int(getattr(metrics, "avg_monthly_searches", 0) or 0) if metrics else 0
            if text:
                result[normalize_keyword(text)] = KeywordMetric(text, searches, self.source_name)
        return result

    def resource_name(self, resource: str, value: str) -> str:
        value = str(value).strip()
        if "/" in value:
            return value
        if resource == "language_constant":
            return f"languageConstants/{value}"
        if resource == "geo_target_constant":
            return f"geoTargetConstants/{value}"
        raise ValueError(f"Nieznany resource: {resource}")

    def resource_names(self, resource: str, values: list[str]) -> list[str]:
        return [self.resource_name(resource, value) for value in values if str(value).strip()]

    def network_enum(self) -> Any:
        enum = self.client.enums.KeywordPlanNetworkEnum
        return getattr(enum, self.network, enum.GOOGLE_SEARCH)


def ensure_grpc_ssl_roots() -> None:
    if os.environ.get("GRPC_DEFAULT_SSL_ROOTS_FILE_PATH"):
        return
    try:
        import certifi
    except ImportError:
        return
    os.environ["GRPC_DEFAULT_SSL_ROOTS_FILE_PATH"] = str(build_windows_grpc_roots_bundle(certifi.where()))


def build_windows_grpc_roots_bundle(certifi_path: str) -> Path:
    if os.name != "nt":
        return Path(certifi_path)

    extra_roots = export_windows_tls_inspection_roots()
    if not extra_roots:
        return Path(certifi_path)

    bundle_path = Path("cache") / "grpc_roots_with_windows.pem"
    ensure_dir(bundle_path.parent)
    pem = Path(certifi_path).read_text(encoding="ascii")
    for root in extra_roots:
        pem += f"\n-----BEGIN CERTIFICATE-----\n{root}\n-----END CERTIFICATE-----\n"
    bundle_path.write_text(pem, encoding="ascii")
    return bundle_path


def export_windows_tls_inspection_roots() -> list[str]:
    ps_script = r"""
$certs = Get-ChildItem Cert:\CurrentUser\Root,Cert:\LocalMachine\Root -ErrorAction SilentlyContinue |
  Where-Object {
    $_.Subject -like '*Norton Web/Mail Shield Root*' -or
    $_.Subject -like '*Norton*SSL*' -or
    $_.Subject -like '*Norton*TLS*'
  } |
  Sort-Object Thumbprint -Unique
foreach ($cert in $certs) {
  [Convert]::ToBase64String($cert.RawData, 'InsertLineBreaks')
  '---CERT---'
}
"""
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
            capture_output=True,
            text=True,
            check=True,
            timeout=15,
        )
    except (OSError, subprocess.SubprocessError):
        return []

    roots: list[str] = []
    for block in result.stdout.split("---CERT---"):
        cleaned = "\n".join(line.strip() for line in block.splitlines() if line.strip())
        if cleaned:
            roots.append(cleaned)
    return roots


class CachedKeywordProvider:
    def __init__(
        self,
        provider: KeywordProvider,
        cache_path: str | Path,
        geo_targets: list[str],
        language: str,
        network: str,
    ) -> None:
        self.provider = provider
        self.source_name = provider.source_name
        self.cache_path = Path(cache_path)
        self.geo_targets = geo_targets
        self.language = language
        self.network = network
        self.cache = load_json_cache(self.cache_path)

    def generate_ideas(self, product_type: str) -> list[str]:
        return self.provider.generate_ideas(product_type)

    def historical_metrics(self, keywords: list[str]) -> dict[str, KeywordMetric]:
        result: dict[str, KeywordMetric] = {}
        missing: list[str] = []
        for keyword in keywords:
            key = self.cache_key(keyword)
            cached = self.cache.get(key)
            if isinstance(cached, dict):
                result[normalize_keyword(keyword)] = KeywordMetric(
                    keyword=compact_spaces(str(cached.get("keyword") or keyword)),
                    avg_monthly_searches=parse_int(cached.get("avg_monthly_searches", 0)),
                    source=compact_spaces(str(cached.get("source") or self.source_name)),
                )
            else:
                missing.append(keyword)
        if missing:
            fresh = self.provider.historical_metrics(missing)
            for normalized, metric in fresh.items():
                result[normalized] = metric
                self.cache[self.cache_key(metric.keyword)] = {
                    "keyword": metric.keyword,
                    "avg_monthly_searches": metric.avg_monthly_searches,
                    "source": metric.source or self.source_name,
                }
            save_json_cache(self.cache_path, self.cache)
        return result

    def cache_key(self, keyword: str) -> str:
        geo = ",".join(sorted(str(value).strip() for value in self.geo_targets if str(value).strip()))
        return "|".join([normalize_keyword(keyword), geo, str(self.language), str(self.network)])


def detect_type_column(df: pd.DataFrame) -> str | None:
    normalized_columns = {normalize_header(str(column)): str(column) for column in df.columns}
    for candidate in TYPE_COLUMN_CANDIDATES:
        match = normalized_columns.get(normalize_header(candidate))
        if match:
            return match
    return None


def normalize_product_type(value: Any) -> str:
    text = compact_spaces(str(value or ""))
    if not text or normalize_header(text) in {"nan", "none", "brak"}:
        return ""
    return text


def analyze_keywords_for_dataframe(
    df: pd.DataFrame,
    provider: KeywordProvider,
    max_candidates: int = 5,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    result = df.copy()
    type_column = detect_type_column(result)
    for column in OUTPUT_COLUMNS:
        result[column] = ""
    if not type_column:
        missing = pd.DataFrame([{"reason": "missing_type_column", "type_column_candidates": ";".join(TYPE_COLUMN_CANDIDATES)}])
        return result, empty_summary(), missing, empty_errors()

    result["__keyword_product_type"] = result[type_column].map(normalize_product_type)
    brands = collect_values(result, BRAND_COLUMN_CANDIDATES)
    codes = collect_values(result, CODE_COLUMN_CANDIDATES)
    product_types = unique_preserve_order(
        [value for value in result["__keyword_product_type"].astype(str).tolist() if compact_spaces(value)]
    )

    analyses: dict[str, KeywordAnalysis] = {}
    error_rows: list[dict[str, str]] = []
    for product_type in product_types:
        try:
            analyses[product_type] = analyze_product_type_keywords(
                product_type,
                provider,
                brands,
                codes,
                max_candidates,
            )
        except Exception as exc:  # API errors should not stop enrichment for the whole file.
            error_rows.append({"product_type": product_type, "error": str(exc)})
            analyses[product_type] = KeywordAnalysis(product_type=product_type, keyword_status="ERROR", keyword_source=provider.source_name)

    for index, row in result.iterrows():
        product_type = str(row.get("__keyword_product_type", ""))
        analysis = analyses.get(product_type)
        if not analysis:
            result.at[index, "keyword_status"] = "NO_TYPE"
            result.at[index, "keyword_candidates_json"] = "[]"
            result.at[index, "keyword_source"] = provider.source_name
            continue
        for column, value in analysis_to_row(analysis).items():
            result.at[index, column] = value

    missing_rows = [
        {"row_index": index + 2, "reason": "empty_product_type"}
        for index, value in result["__keyword_product_type"].items()
        if is_blank(value)
    ]
    result = result.drop(columns=["__keyword_product_type"])
    return (
        result,
        pd.DataFrame([analysis_to_summary_row(value) for value in analyses.values()]),
        pd.DataFrame(missing_rows),
        pd.DataFrame(error_rows),
    )


def analyze_product_type_keywords(
    product_type: str,
    provider: KeywordProvider,
    brands: set[str],
    codes: set[str],
    max_candidates: int = 5,
) -> KeywordAnalysis:
    raw_ideas = provider.generate_ideas(product_type)
    candidates = select_candidate_keywords(product_type, raw_ideas, brands, codes, max_candidates)
    if not candidates:
        return KeywordAnalysis(product_type=product_type, keyword_status="NO_DATA", keyword_source=provider.source_name)

    metrics = provider.historical_metrics(candidates)
    ranked = rank_keyword_metrics(product_type, candidates, metrics)
    candidate_rows = [
        {
            "keyword": keyword,
            "avg_monthly_searches": int(metrics.get(normalize_keyword(keyword), KeywordMetric(keyword, 0)).avg_monthly_searches),
        }
        for keyword in candidates
    ]
    if len(ranked) < 2:
        return KeywordAnalysis(
            product_type=product_type,
            primary_keyword=ranked[0].keyword if ranked else "",
            primary_keyword_avg_monthly_searches=ranked[0].avg_monthly_searches if ranked else "",
            keyword_candidates_json=json.dumps(candidate_rows, ensure_ascii=False, separators=(",", ":")),
            keyword_status="NO_DATA",
            keyword_source=provider.source_name,
        )
    return KeywordAnalysis(
        product_type=product_type,
        primary_keyword=ranked[0].keyword,
        primary_keyword_avg_monthly_searches=ranked[0].avg_monthly_searches,
        secondary_keyword=ranked[1].keyword,
        secondary_keyword_avg_monthly_searches=ranked[1].avg_monthly_searches,
        keyword_candidates_json=json.dumps(candidate_rows, ensure_ascii=False, separators=(",", ":")),
        keyword_status="OK",
        keyword_source=provider.source_name,
    )


def select_candidate_keywords(
    product_type: str,
    ideas: list[str],
    brands: set[str],
    codes: set[str],
    limit: int = 5,
) -> list[str]:
    filtered: list[str] = []
    seen: set[str] = set()
    for idea in ideas:
        keyword = compact_spaces(idea).lower()
        normalized = normalize_keyword(keyword)
        if not keyword or normalized in seen:
            continue
        if not keyword_allowed(keyword, brands, codes):
            continue
        seen.add(normalized)
        filtered.append(keyword)
    ranked = sorted(filtered, key=lambda keyword: relevance_sort_key(product_type, keyword))
    return ranked[:limit]


def keyword_allowed(keyword: str, brands: set[str], codes: set[str]) -> bool:
    if len(keyword) > 80 or len(keyword.split()) > 10:
        return False
    normalized = normalize_header(keyword)
    tokens = set(normalized.split())
    if not tokens:
        return False
    if any(normalize_header(brand) in normalized for brand in brands if len(normalize_header(brand)) >= 3):
        return False
    for code in codes:
        clean_code = compact_spaces(str(code))
        if len(clean_code) >= 4 and re.search(rf"(?<!\w){re.escape(clean_code.lower())}(?!\w)", keyword.lower()):
            return False
    if re.search(r"\b\d{8,14}\b", keyword):
        return False
    return True


def relevance_sort_key(product_type: str, keyword: str) -> tuple[int, int, int, str]:
    type_tokens = set(normalize_header(product_type).split())
    keyword_tokens = set(normalize_header(keyword).split())
    overlap = len(type_tokens & keyword_tokens)
    contains_phrase = normalize_header(product_type) in normalize_header(keyword)
    return (-int(contains_phrase), -overlap, len(keyword), keyword)


def rank_keyword_metrics(
    product_type: str,
    candidates: list[str],
    metrics: dict[str, KeywordMetric],
) -> list[KeywordMetric]:
    ranked: list[KeywordMetric] = []
    for keyword in candidates:
        metric = metrics.get(normalize_keyword(keyword), KeywordMetric(keyword=keyword, avg_monthly_searches=0))
        ranked.append(KeywordMetric(keyword=keyword, avg_monthly_searches=int(metric.avg_monthly_searches), source=metric.source))
    return sorted(ranked, key=lambda metric: metric_sort_key(product_type, metric))


def metric_sort_key(product_type: str, metric: KeywordMetric) -> tuple[int, tuple[int, int, int, str]]:
    return (-int(metric.avg_monthly_searches), relevance_sort_key(product_type, metric.keyword))


def collect_values(df: pd.DataFrame, candidates: list[str]) -> set[str]:
    columns = {normalize_header(str(column)): str(column) for column in df.columns}
    result: set[str] = set()
    for candidate in candidates:
        column = columns.get(normalize_header(candidate))
        if not column:
            continue
        for value in df[column].dropna().astype(str):
            cleaned = compact_spaces(value)
            if cleaned:
                result.add(cleaned)
    return result


def analysis_to_row(analysis: KeywordAnalysis) -> dict[str, Any]:
    return {
        "primary_keyword": analysis.primary_keyword,
        "primary_keyword_avg_monthly_searches": analysis.primary_keyword_avg_monthly_searches,
        "secondary_keyword": analysis.secondary_keyword,
        "secondary_keyword_avg_monthly_searches": analysis.secondary_keyword_avg_monthly_searches,
        "keyword_candidates_json": analysis.keyword_candidates_json,
        "keyword_status": analysis.keyword_status,
        "keyword_source": analysis.keyword_source,
    }


def analysis_to_summary_row(analysis: KeywordAnalysis) -> dict[str, Any]:
    return {"product_type": analysis.product_type, **analysis_to_row(analysis)}


def normalize_keyword(value: Any) -> str:
    return normalize_header(compact_spaces(str(value or "")))


def unique_preserve_order(values: list[str]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        cleaned = compact_spaces(str(value or ""))
        key = normalize_keyword(cleaned)
        if not key or key in seen:
            continue
        seen.add(key)
        result.append(cleaned)
    return result


def first_existing(row: dict[str, Any], candidates: list[str]) -> str:
    normalized = {normalize_header(key): value for key, value in row.items()}
    for candidate in candidates:
        value = normalized.get(normalize_header(candidate))
        if value is not None and not is_blank(value):
            return compact_spaces(str(value))
    return ""


def parse_int(value: Any) -> int:
    text = compact_spaces(str(value or ""))
    if not text:
        return 0
    text = text.replace("\u00a0", "").replace(" ", "").replace(",", "")
    match = re.search(r"-?\d+", text)
    return int(match.group(0)) if match else 0


def load_json_cache(path: str | Path) -> dict[str, Any]:
    cache_path = Path(path)
    if not cache_path.exists():
        return {}
    try:
        return json.loads(cache_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}


def save_json_cache(path: str | Path, data: dict[str, Any]) -> None:
    cache_path = Path(path)
    ensure_dir(cache_path.parent)
    cache_path.write_text(json.dumps(data, ensure_ascii=False, indent=2, sort_keys=True), encoding="utf-8")


def empty_summary() -> pd.DataFrame:
    return pd.DataFrame(columns=["product_type", *OUTPUT_COLUMNS])


def empty_errors() -> pd.DataFrame:
    return pd.DataFrame(columns=["product_type", "error"])


def build_provider(args: argparse.Namespace) -> KeywordProvider:
    geo_targets = split_csv_arg(args.geo_targets)
    if args.offline_keywords:
        provider: KeywordProvider = OfflineKeywordProvider(args.offline_keywords)
    else:
        provider = GoogleAdsKeywordProvider(
            customer_id=effective_customer_id(args),
            google_ads_config=args.google_ads_config,
            geo_targets=geo_targets,
            language=args.language,
            network=args.network,
        )
    return CachedKeywordProvider(provider, args.cache, geo_targets, args.language, args.network)


def add_keyword_pipeline_arguments(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--analyze-keywords", action="store_true", help="Dopisuje primary/secondary keyword per typ produktu.")
    parser.add_argument("--keyword-reports-dir", default="", help="Katalog raportow keywordow. Domyslnie: <reports-dir>/keywords.")
    parser.add_argument("--customer-id", default="", help="Google Ads customer ID bez myslnikow. Opcjonalne, jesli YAML ma customer_id.")
    parser.add_argument("--google-ads-config", default="", help="Opcjonalna sciezka do google-ads.yaml.")
    parser.add_argument("--geo-targets", default=",".join(DEFAULT_GEO_TARGETS), help="Geo target constants, np. 2616.")
    parser.add_argument("--language", default=DEFAULT_LANGUAGE, help="Language constant, domyslnie 1045 dla polskiego.")
    parser.add_argument("--network", default=DEFAULT_NETWORK, help="KeywordPlanNetwork, domyslnie GOOGLE_SEARCH.")
    parser.add_argument("--cache", default=DEFAULT_CACHE_PATH, help="Cache metryk Keyword Planner.")
    parser.add_argument("--max-candidates", type=int, default=5, help="Liczba kandydatow sprawdzanych w KP.")
    parser.add_argument("--offline-keywords", default="", help="CSV z kolumnami keyword, avg_monthly_searches i opcjonalnie product_type.")


def maybe_enrich_pipeline_with_keywords(
    df: pd.DataFrame,
    args: argparse.Namespace,
    reports_dir: str | Path,
) -> pd.DataFrame:
    if not getattr(args, "analyze_keywords", False):
        return df
    if not args.offline_keywords and not effective_customer_id(args):
        raise SystemExit("Podaj --customer-id, dodaj customer_id do google-ads.yaml albo uzyj --offline-keywords.")

    keyword_reports_dir = Path(args.keyword_reports_dir) if args.keyword_reports_dir else Path(reports_dir) / "keywords"
    provider = build_provider(args)
    result, summary, missing, errors = analyze_keywords_for_dataframe(df, provider, max_candidates=args.max_candidates)
    write_reports(summary, missing, errors, keyword_reports_dir)
    print(f"OK: przeanalizowano slowa kluczowe dla typow produktu: {len(summary)}")
    print(f"OK: zapisano raporty slow kluczowych w {keyword_reports_dir}")
    if len(errors):
        print(f"UWAGA: bledy API/providerow slow kluczowych: {len(errors)}")
    return result


def split_csv_arg(value: str) -> list[str]:
    return [part.strip() for part in str(value or "").split(",") if part.strip()]


def effective_customer_id(args: argparse.Namespace) -> str:
    cli_customer_id = clean_customer_id(args.customer_id)
    if cli_customer_id:
        return cli_customer_id
    return clean_customer_id(customer_id_from_google_ads_config(args.google_ads_config))


def clean_customer_id(value: Any) -> str:
    return re.sub(r"\D", "", str(value or ""))


def customer_id_from_google_ads_config(config_path: str = "") -> str:
    path = resolve_google_ads_config_path(config_path)
    if not path or not path.exists():
        return ""
    config = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    return first_present_config_value(config, ["customer_id", "login_customer_id"])


def resolve_google_ads_config_path(config_path: str = "") -> Path | None:
    if config_path:
        return Path(config_path)
    env_path = os.environ.get("GOOGLE_ADS_CONFIGURATION_FILE_PATH")
    if env_path:
        return Path(env_path)
    home_path = Path.home() / "google-ads.yaml"
    if home_path.exists():
        return home_path
    local_path = Path("google-ads.yaml")
    if local_path.exists():
        return local_path
    return home_path


def first_present_config_value(config: dict[str, Any], keys: list[str]) -> str:
    for key in keys:
        value = config.get(key)
        if value is not None and not is_blank(value):
            return compact_spaces(str(value))
    return ""


def write_reports(summary: pd.DataFrame, missing: pd.DataFrame, errors: pd.DataFrame, reports_dir: str | Path) -> None:
    reports = ensure_dir(reports_dir)
    summary.to_csv(reports / "keyword_type_summary.csv", index=False, encoding="utf-8-sig")
    missing.to_csv(reports / "keyword_missing_types.csv", index=False, encoding="utf-8-sig")
    errors.to_csv(reports / "keyword_api_errors.csv", index=False, encoding="utf-8-sig")


def main() -> None:
    parser = argparse.ArgumentParser(description="Analizuje slowa kluczowe per typ produktu przez Keyword Planner.")
    parser.add_argument("--input", required=True, help="Plik CSV/XLSX z produktami.")
    parser.add_argument("--sheet", help="Arkusz XLSX, jesli input ma wiele arkuszy.")
    parser.add_argument("--output", required=True, help="Docelowy plik CSV/XLSX z kolumnami keywordow.")
    parser.add_argument("--reports-dir", default="reports/keywords", help="Katalog raportow keywordow.")
    parser.add_argument("--customer-id", default="", help="Google Ads customer ID bez myslnikow. Opcjonalne, jesli YAML ma customer_id.")
    parser.add_argument("--google-ads-config", default="", help="Opcjonalna sciezka do google-ads.yaml.")
    parser.add_argument("--geo-targets", default=",".join(DEFAULT_GEO_TARGETS), help="Geo target constants, np. 2616.")
    parser.add_argument("--language", default=DEFAULT_LANGUAGE, help="Language constant, domyslnie 1045 dla polskiego.")
    parser.add_argument("--network", default=DEFAULT_NETWORK, help="KeywordPlanNetwork, domyslnie GOOGLE_SEARCH.")
    parser.add_argument("--cache", default=DEFAULT_CACHE_PATH, help="Cache metryk Keyword Planner.")
    parser.add_argument("--max-candidates", type=int, default=5, help="Liczba kandydatow sprawdzanych w KP.")
    parser.add_argument("--offline-keywords", default="", help="CSV z kolumnami keyword, avg_monthly_searches i opcjonalnie product_type.")
    args = parser.parse_args()

    if not args.offline_keywords and not effective_customer_id(args):
        raise SystemExit("Podaj --customer-id, dodaj customer_id do google-ads.yaml albo uzyj --offline-keywords.")

    df = read_products(args.input, sheet_name=args.sheet)
    provider = build_provider(args)
    result, summary, missing, errors = analyze_keywords_for_dataframe(df, provider, max_candidates=args.max_candidates)
    write_products(result, args.output)
    write_reports(summary, missing, errors, args.reports_dir)

    print(f"OK: wczytano {len(df)} produktow")
    print(f"OK: zapisano {args.output}")
    print(f"OK: zapisano raporty w {args.reports_dir}")
    print(f"OK: przeanalizowano typy produktow: {len(summary)}")
    if len(errors):
        print(f"UWAGA: bledy API/providerow: {len(errors)}")


if __name__ == "__main__":
    main()
