from __future__ import annotations

import base64
from io import BytesIO
from pathlib import Path
from urllib.parse import urlsplit

from PIL import Image, ImageOps


MAX_IMAGE_COUNT = 16
MAX_BASE64_CHARS = 2_000_000
MAX_SOURCE_BYTES = 50 * 1024 * 1024
MAX_IMAGE_PIXELS = 50_000_000
BASE_CDN_HOSTS = {"upload.cdn.baselinker.com"}


def external_image_value(url: str, *, allow_http: bool = False) -> tuple[str | None, str | None]:
    candidate = str(url or "").strip()
    if len(candidate) > 1000:
        return None, "URL zdjęcia przekracza limit 1000 znaków."
    try:
        parsed = urlsplit(candidate)
    except ValueError:
        return None, "Niepoprawny URL zdjęcia."
    allowed_schemes = {"https", "http"} if allow_http else {"https"}
    if parsed.scheme.lower() not in allowed_schemes or not parsed.hostname:
        return None, "Zdjęcie musi mieć pełny adres HTTPS."
    if parsed.username or parsed.password:
        return None, "URL zdjęcia nie może zawierać danych logowania."
    if parsed.hostname.casefold() in BASE_CDN_HOSTS:
        return None, "Pominięto istniejący adres Base CDN; nie wolno wysyłać go ponownie jako zdjęcia zewnętrznego."
    return "url:" + candidate, None


def _ensure_rgb(image: Image.Image) -> Image.Image:
    if image.mode == "RGB":
        return image
    if image.mode in {"RGBA", "LA"}:
        background = Image.new("RGB", image.size, (255, 255, 255))
        background.paste(image, mask=image.getchannel("A"))
        return background
    return image.convert("RGB")


def local_image_value(
    path: Path,
    *,
    target_size: int = 1000,
    start_quality: int = 85,
) -> tuple[str | None, list[str]]:
    warnings: list[str] = []
    if not path.is_file():
        return None, [f"Brak lokalnego zdjęcia: {path}"]
    if path.stat().st_size > MAX_SOURCE_BYTES:
        return None, [f"Lokalne zdjęcie przekracza {MAX_SOURCE_BYTES} B: {path}"]

    previous_limit = Image.MAX_IMAGE_PIXELS
    Image.MAX_IMAGE_PIXELS = MAX_IMAGE_PIXELS
    try:
        with Image.open(path) as opened:
            opened.load()
            image = ImageOps.exif_transpose(opened)
            image = _ensure_rgb(image)
            image = ImageOps.pad(
                image,
                (target_size, target_size),
                method=Image.Resampling.LANCZOS,
                color=(255, 255, 255),
            )
            quality = min(95, max(30, start_quality))
            for shrink_round in range(10):
                while quality >= 30:
                    buffer = BytesIO()
                    image.save(buffer, format="JPEG", quality=quality, optimize=True, progressive=True)
                    encoded = base64.b64encode(buffer.getvalue()).decode("ascii")
                    if len(encoded) <= MAX_BASE64_CHARS:
                        return "data:" + encoded, warnings
                    quality -= 10
                new_side = int(image.width * 0.9)
                if new_side < 100 or shrink_round == 9:
                    break
                image = image.resize((new_side, new_side), Image.Resampling.LANCZOS)
                warnings.append(f"Zmniejszono zdjęcie do {new_side}x{new_side}, aby zmieścić limit API.")
                quality = min(95, max(30, start_quality))
    except (OSError, ValueError, Image.DecompressionBombError) as exc:
        return None, [f"Nie można przygotować zdjęcia {path}: {exc}"]
    finally:
        Image.MAX_IMAGE_PIXELS = previous_limit
    return None, warnings + [f"Nie udało się zmieścić zdjęcia w limicie API: {path}"]


def build_images_payload(
    urls: list[str],
    paths: list[str],
    *,
    input_dir: Path,
    allow_http: bool = False,
) -> tuple[dict[str, str], list[str]]:
    values: list[str] = []
    warnings: list[str] = []
    seen: set[str] = set()

    for url in urls:
        value, warning = external_image_value(url, allow_http=allow_http)
        if warning:
            warnings.append(f"{url[:120]}: {warning}")
        if value and value not in seen:
            seen.add(value)
            values.append(value)

    for raw_path in paths:
        path = Path(raw_path)
        if not path.is_absolute():
            path = input_dir / path
        identity = str(path.resolve()).casefold()
        if identity in seen:
            continue
        seen.add(identity)
        value, local_warnings = local_image_value(path)
        warnings.extend(local_warnings)
        if value:
            values.append(value)

    if len(values) > MAX_IMAGE_COUNT:
        warnings.append(f"Pominięto zdjęcia ponad limit {MAX_IMAGE_COUNT} pozycji.")
    return {str(index): value for index, value in enumerate(values[:MAX_IMAGE_COUNT])}, warnings
