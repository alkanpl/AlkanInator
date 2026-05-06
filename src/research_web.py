from __future__ import annotations

import argparse


def main() -> None:
    parser = argparse.ArgumentParser(description="Placeholder researchu internetowego. Ten etap jest celowo poza MVP.")
    parser.add_argument("--input", required=True)
    parser.add_argument("--limit", type=int, default=100)
    args = parser.parse_args()
    print(f"Pominieto research internetowy dla {args.input}; limit={args.limit}. Modul nie jest czescia MVP.")


if __name__ == "__main__":
    main()
