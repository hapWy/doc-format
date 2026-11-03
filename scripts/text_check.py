#!/usr/bin/env python3
"""Read-only editorial hints, not an AI detector or fact checker. Python 3.9+."""
import argparse
import collections
import json
from pathlib import Path
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from text_zones import protected_zones


def read_text(path):
    if path.suffix.lower() != ".docx":
        return path.read_text(encoding="utf-8-sig")
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    paragraphs = []
    with zipfile.ZipFile(path) as archive:
        names = ["word/document.xml"] + sorted(
            n for n in archive.namelist()
            if re.fullmatch(r"word/(?:header\d+|footer\d+|footnotes|endnotes)\.xml", n)
        )
        for name in names:
            root = ET.fromstring(archive.read(name))
            for paragraph in root.findall(".//w:p", ns):
                # Deleted revisions are not w:t; additions are included.
                parts = [n.text or "" for n in paragraph.findall(".//w:t", ns)]
                if parts:
                    paragraphs.append("".join(parts))
    return "\n\n".join(paragraphs)


def mask(markdown):
    """Keep line numbers while masking shared Markdown exclusion zones."""
    protected, fence = protected_zones(markdown)
    masked = ''.join(c if c in '\r\n' or not protected[i] else ' '
                     for i, c in enumerate(markdown))
    return masked.splitlines(), fence


def check(text):
    lines, fence = mask(text)
    findings = []

    def add(rule, severity, line, excerpt, message):
        findings.append(dict(rule=rule, severity=severity, line=line,
                             excerpt=excerpt[:180], message=message))

    if fence:
        add("unclosed-fence", "error", fence[2], "", "Не закрыт блок кода Markdown.")
    placeholders = re.compile(r"\b(?:TODO|FIXME|TBD)\b|\[(?:уточнить|вставить|заполнить|нужн[аоы]?\b|данные отсутствуют)[^\]]*\]", re.I)
    tokens = re.compile(r"(?:turn\d+(?:search|view|fetch)\d+|[^\n]*?|::(?:chatgpt-content-reference|codex-file-citation)|<\|(?:im_start|im_end|endoftext)\|>)")
    # Ignore NBSP, soft hyphen, ZWJ and variation selectors, which can be meaningful.
    controls = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f\u200b\u200e\u200f\u202a-\u202e\u2060\u2066-\u2069\ufeff]")
    openers = collections.defaultdict(list)
    contrasts = []
    authority = re.compile(r'\b(?:исследования показывают|эксперты (?:считают|согласны)|общеизвестно|как известно|все знают)\b', re.I)
    contrast = re.compile(r'\bне (?:просто|только)\b.+\b(?:а|но)\b', re.I)
    opener_pattern = re.compile(r"^\s*(?:#{1,6}\s+)?(в современном мире|в настоящее время|важно отметить|следует отметить|необходимо отметить|таким образом|подводя итог|в заключение)\b", re.I)
    for number, line in enumerate(lines, 1):
        for match in placeholders.finditer(line):
            add("placeholder", "warning", number, match[0], "Проверьте, заполнен ли фрагмент; не заменяйте его выдуманными данными.")
        for match in tokens.finditer(line):
            add("tool-token", "error", number, match[0], "В тексте осталась служебная ссылка или метка инструмента.")
        for match in controls.finditer(line):
            add("invisible-control", "warning", number, f"U+{ord(match[0]):04X}", "Проверьте невидимый управляющий символ; автоматического удаления нет.")
        opener = opener_pattern.search(line)
        if opener:
            openers[opener[1].lower()].append(number)
        for match in authority.finditer(line):
            add('unattributed-authority', 'warning', number, match[0],
                'Проверьте, назван ли источник и подтверждает ли он утверждение. Не дописывайте источник по догадке.')
        if contrast.search(line):
            contrasts.append(number)
    for phrase, numbers in openers.items():
        if len(numbers) >= 3:
            add("repeated-opener", "warning", numbers[0], phrase,
                "Одинаковое начало встречается %d раз (строки %s). Оцените его функцию; само выражение допустимо." % (len(numbers), ", ".join(map(str, numbers))))
    if len(contrasts) >= 3:
        add('repeated-contrast', 'warning', contrasts[0], 'не просто / не только',
            'Повторяется одна конструкция противопоставления (строки %s). Сохраните полезные сравнения, проверьте шаблонный ритм.' % ', '.join(map(str, contrasts)))

    # Full paragraphs, not coincidentally similar sentences. Short labels are ignored.
    paragraphs = collections.defaultdict(list)
    current, start = [], 1
    for number, line in enumerate(lines + [""], 1):
        if not line.strip():
            paragraph = " ".join(current)
            normalized = re.sub(r"\s+", " ", paragraph).strip().casefold()
            if len(normalized) >= 60:
                paragraphs[normalized].append((start, paragraph))
            current = []
        else:
            if not current:
                start = number
            current.append(line.strip())
    for instances in paragraphs.values():
        if len(instances) > 1:
            add("duplicate-paragraph", "warning", instances[1][0], instances[1][1],
                "Повтор абзаца; первое вхождение в строке %d. Проверьте, нужен ли повтор." % instances[0][0])
    return sorted(findings, key=lambda f: (f["line"], f["rule"]))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()
    try:
        if args.input.suffix.lower() not in (".md", ".markdown", ".txt", ".docx"):
            raise ValueError("Поддерживаются .md, .markdown, .txt и .docx; для PDF сначала извлеките текст.")
        findings = check(read_text(args.input))
    except (OSError, ValueError, UnicodeError, zipfile.BadZipFile, KeyError, ET.ParseError) as exc:
        if args.json:
            print(json.dumps({"error": str(exc)}, ensure_ascii=False))
        else:
            print(f"Ошибка: {exc}", file=sys.stderr)
        return 2
    result = {"input": str(args.input), "findings": findings,
              "limits": "Эвристики не определяют авторство, истинность фактов или качество вёрстки. Для DOCX номера строк относятся к извлечённому тексту, не к страницам."}
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        for item in findings:
            print(f"{item['severity']}:{item['line']} [{item['rule']}] {item['message']} {item['excerpt']}")
        print(f"Находок: {len(findings)}. " + result["limits"])
    # Warnings need an editorial decision; 1 means review, not automatic rejection.
    return 1 if findings else 0


if __name__ == "__main__":
    sys.exit(main())
