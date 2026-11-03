#!/usr/bin/env python3
"""Report local document tools without installing packages or invoking renderers."""
import argparse
import importlib.util
import json
import shutil
import sys

CAPABILITIES = ('structural-checks', 'docx-create', 'docx-render',
                'pdf-create', 'pdf-render', 'pdf-text')


def module(name):
    try:
        return importlib.util.find_spec(name) is not None
    except (ImportError, ValueError):
        return False


def program(explicit, names):
    return bool(shutil.which(explicit)) if explicit else any(shutil.which(n) for n in names)


def report(soffice=None, pdftoppm=None, required=()):
    python = sys.version_info >= (3, 9)
    modules = {name: module(name) for name in ('docx', 'reportlab', 'fitz', 'pypdf', 'pdfplumber')}
    tools = {'soffice': program(soffice, ('soffice', 'libreoffice')),
             'pdftoppm': program(pdftoppm, ('pdftoppm',)),
             'pdftotext': program(None, ('pdftotext',))}
    images = tools['pdftoppm'] if pdftoppm else modules['fitz'] or tools['pdftoppm']
    capabilities = {'structural-checks': python,
                    'docx-create': python and modules['docx'],
                    'docx-render': python and tools['soffice'] and images,
                    'pdf-create': python and modules['reportlab'],
                    'pdf-render': python and images,
                    'pdf-text': python and (modules['pypdf'] or modules['pdfplumber'] or tools['pdftotext'])}
    missing = [name for name in required if not capabilities[name]]
    return {'python': '.'.join(map(str, sys.version_info[:3])),
            'modules': modules, 'programs': tools, 'capabilities': capabilities,
            'required': list(required), 'missing': missing,
            'limits': 'Presence checks only. Renderers were not launched. Fonts, actual page layout and the agent image viewer were not checked.'}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--require', action='append', choices=CAPABILITIES, default=[])
    parser.add_argument('--soffice', help='Same executable path passed to render_pages.py')
    parser.add_argument('--pdftoppm', help='Same executable path passed to render_pages.py')
    parser.add_argument('--json', action='store_true')
    args = parser.parse_args()
    result = report(args.soffice, args.pdftoppm, args.require)
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        for name, available in result['capabilities'].items():
            print(f"{name}: {'available' if available else 'missing'}")
        print(result['limits'])
    return 1 if result['missing'] else 0


if __name__ == '__main__':
    sys.exit(main())
