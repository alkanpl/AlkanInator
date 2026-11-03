"""Aplikacja "Wgraj zmiany do sklepu": podglad zmian z pliku XLSX i zapis po zatwierdzeniu, bez terminala.

    py src/woo_preview_app.py          (albo dwuklik "Wgraj zmiany do sklepu.bat" w katalogu repo)

Szczegoly w src/woo_preview/server.py.
"""

from woo_preview.server import main


if __name__ == "__main__":
    raise SystemExit(main())
