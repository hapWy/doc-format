#!/usr/bin/env python3
"""Inspect UTF-8 prose; optionally save a new copy without stray invisible chars.
Preserves Russian punctuation, mathematical symbols and Markdown exclusion zones.
"""
import argparse
import json
from pathlib import Path
import sys
from text_zones import protected_zones


def protected_positions(text):
    return protected_zones(text)[0]


def scan(text):
    protected = protected_positions(text)
    cleaned, findings = [], []
    line, column = 1, 1
    for offset, char in enumerate(text):
        cp = ord(char)
        remove = (cp in (0x200B, 0xFEFF) or cp < 32 and char not in '\t\r\n'
                  or cp == 127 or 0xFDD0 <= cp <= 0xFDEF or cp & 0xFFFF in (0xFFFE, 0xFFFF))
        review = cp in (0x200E, 0x200F, 0x2060) or 0x202A <= cp <= 0x202E or 0x2066 <= cp <= 0x2069
        # A leading BOM is an encoding marker, not a stray character in the prose.
        if offset == 0 and cp == 0xFEFF:
            remove = False
        if not protected[offset] and (remove or review):
            findings.append({'line': line, 'column': column, 'codepoint': 'U+%04X' % cp,
                             'action': 'remove' if remove else 'review',
                             'message': 'Посторонний невидимый символ.' if remove else
                             'Управление направлением или соединением текста: проверьте назначение, автоматического удаления нет.'})
        if protected[offset] or not remove:
            cleaned.append(char)
        if char == '\n':
            line, column = line + 1, 1
        else:
            column += 1
    return ''.join(cleaned), findings


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input', type=Path)
    parser.add_argument('--out', type=Path, help='New UTF-8 file; input and existing output are never overwritten')
    parser.add_argument('--json', action='store_true')
    args = parser.parse_args()
    try:
        if args.input.suffix.lower() not in ('.md', '.markdown', '.txt'):
            raise ValueError('Поддерживается только UTF-8 текст; DOCX/PDF сначала извлеките в рабочую копию.')
        with args.input.open(encoding='utf-8', newline='') as stream:
            cleaned, findings = scan(stream.read())
        if args.out:
            if args.out.resolve() == args.input.resolve():
                raise ValueError('Вход и выход должны различаться.')
            with args.out.open('x', encoding='utf-8', newline='') as stream:
                stream.write(cleaned)
        pending = sum(f['action'] == 'review' for f in findings)
        result = {'input': str(args.input), 'output': str(args.out) if args.out else None,
                  'findings': findings, 'removed_count': sum(f['action'] == 'remove' for f in findings) if args.out else 0,
                  'pending_review_count': pending,
                  'limits': 'Не меняет факты, пунктуацию, математические знаки, неразрывные пробелы и мягкие переносы. Код, цитаты, URL и frontmatter защищены; их проверяют отдельно.'}
        if args.json:
            print(json.dumps(result, ensure_ascii=False, indent=2))
        else:
            for item in findings:
                print('{line}:{column} {codepoint} {action}: {message}'.format(**item))
            print(result['limits'])
        return 1 if pending or findings and not args.out else 0
    except (OSError, UnicodeError, ValueError) as error:
        print(json.dumps({'error': str(error)}, ensure_ascii=False) if args.json else str(error), file=sys.stderr)
        return 2


if __name__ == '__main__':
    sys.exit(main())
