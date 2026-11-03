#!/usr/bin/env python3
"""Render a local DOCX/PDF to PNG pages for subsequent human/agent inspection."""
import argparse
from contextlib import contextmanager
import importlib.util
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET


def executable(explicit, names):
    if explicit:
        found = shutil.which(explicit)
        if not found:
            raise ValueError(f"Не найден исполняемый файл: {explicit}")
        return found
    return next((found for name in names if (found := shutil.which(name))), None)


@contextmanager
def font_environment(font_dirs):
    """Give child tools a private Fontconfig file; leave process/global settings alone."""
    directories = []
    for directory in font_dirs:
        directory = directory.resolve(strict=True)
        if not directory.is_dir():
            raise ValueError(f"--font-dir должен указывать на существующий каталог: {directory}")
        if directory not in directories:
            directories.append(directory)
    if not directories:
        yield None
        return
    with tempfile.TemporaryDirectory(prefix="doc-format-fonts-") as temporary:
        temp = Path(temporary)
        cache = temp / "cache"
        cache.mkdir()
        config = ET.Element("fontconfig")
        for directory in directories:
            ET.SubElement(config, "dir").text = str(directory)
        ET.SubElement(config, "cachedir").text = str(cache)
        config_path = temp / "doc-format-fonts.conf"
        ET.ElementTree(config).write(config_path, encoding="utf-8", xml_declaration=True)
        env = os.environ.copy()
        env["FONTCONFIG_FILE"] = str(config_path)
        yield env


def run(command, timeout, env=None):
    try:
        result = subprocess.run(command, capture_output=True, text=True,
                                timeout=timeout, check=False, env=env)
    except subprocess.TimeoutExpired as exc:
        raise ValueError(f"Превышено время ожидания ({timeout} с): {command[0]}") from exc
    if result.returncode:
        raise ValueError(f"Ошибка {command[0]} ({result.returncode}): {(result.stderr or result.stdout)[-3000:]}")
    return result


def render(args, env=None):
    source = args.input.resolve(strict=True)
    suffix = source.suffix.lower()
    if suffix not in (".docx", ".pdf"):
        raise ValueError("Поддерживаются только DOCX и PDF.")
    out = args.out_dir.resolve()
    if out.exists() and (not out.is_dir() or any(out.iterdir())):
        raise ValueError("Каталог --out-dir должен быть новым или пустым; существующие файлы не перезаписываются.")
    soffice = executable(args.soffice, ("soffice", "libreoffice")) if suffix == ".docx" else None
    if suffix == ".docx" and not soffice:
        raise ValueError("Для DOCX нужен локальный LibreOffice (soffice в PATH или --soffice ПУТЬ). Отдельный сервис не нужен.")
    use_fitz = importlib.util.find_spec("fitz") is not None
    pdftoppm = executable(args.pdftoppm, ("pdftoppm",)) if (args.pdftoppm or not use_fitz) else None
    if not use_fitz and not pdftoppm:
        raise ValueError("Для PNG нужен PyMuPDF в текущем Python или локальный Poppler (pdftoppm в PATH/--pdftoppm).")
    out.mkdir(parents=True, exist_ok=True)
    pdf = source
    if suffix == ".docx":
        with tempfile.TemporaryDirectory(prefix="doc-format-lo-") as profile:
            run([soffice, f"-env:UserInstallation={Path(profile).as_uri()}",
                 "--headless", "--convert-to", "pdf", "--outdir", str(out), str(source)], args.timeout, env)
        pdf = out / (source.stem + ".pdf")
        if not pdf.exists() or pdf.stat().st_size == 0:
            raise ValueError("LibreOffice не создал PDF. Проверьте доступность шрифтов и журнал конвертера.")
    pages = []
    if use_fitz and not args.pdftoppm:
        import fitz
        with fitz.open(pdf) as document:
            if document.needs_pass:
                raise ValueError("PDF защищён паролем; сначала предоставьте доступную копию.")
            for number, page in enumerate(document, 1):
                target = out / f"page-{number:04d}.png"
                page.get_pixmap(matrix=fitz.Matrix(args.dpi / 72, args.dpi / 72), alpha=False).save(target)
                pages.append(target.name)
        renderer = "PyMuPDF"
    else:
        run([pdftoppm, "-png", "-r", str(args.dpi), str(pdf), str(out / "page")], args.timeout, env)
        generated = sorted(out.glob("page-*.png"), key=lambda p: int(p.stem.rsplit("-", 1)[1]))
        for number, old in enumerate(generated, 1):
            new = out / f"page-{number:04d}.png"
            if old != new:
                old.rename(new)
            pages.append(new.name)
        renderer = "pdftoppm"
    if not pages:
        raise ValueError("Не получены изображения страниц.")
    manifest = {"source": str(source), "pdf": str(pdf), "page_count": len(pages),
                "dpi": args.dpi, "renderer": renderer, "pages": pages,
                "visual_review": "pending", "note": "Рендеринг завершён. Откройте и проверьте все PNG; наличие файлов не доказывает качество документа."}
    if args.font_dir:
        manifest["font_dirs"] = [str(p.resolve()) for p in args.font_dir]
    (out / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return manifest


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path)
    parser.add_argument("--out-dir", type=Path, required=True)
    parser.add_argument("--dpi", type=int, default=144)
    parser.add_argument("--timeout", type=int, default=120, help="Seconds per external command")
    parser.add_argument("--soffice", help="Local executable; defaults to PATH lookup")
    parser.add_argument("--pdftoppm", help="Local executable; explicitly selects Poppler")
    parser.add_argument("--font-dir", type=Path, action="append", default=[],
                        help="Installed font directory for Fontconfig-based child tools; repeatable. No fonts are installed or copied.")
    args = parser.parse_args()
    if not 72 <= args.dpi <= 600 or args.timeout < 1:
        parser.error("--dpi: 72..600; --timeout: положительное число секунд")
    try:
        with font_environment(args.font_dir) as env:
            result = render(args, env)
    except (OSError, ValueError, RuntimeError, ImportError) as exc:
        print(f"Ошибка: {exc}", file=sys.stderr)
        return 2
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
