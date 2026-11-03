#!/usr/bin/env python3
"""Copy this skill into a supported agent's project or personal skills folder."""
import argparse
import json
from pathlib import Path
import shutil
import sys
import tempfile

ROOT = Path(__file__).resolve().parent.parent
AGENTS = {'codex': '.agents', 'claude': '.claude', 'cursor': '.cursor'}
ITEMS = ('SKILL.md', 'README.md', 'INSTALL.md', 'rules', 'profiles',
         'scripts', 'agents', 'examples')
IGNORED = shutil.ignore_patterns('.git', '.venv', '__pycache__', '*.pyc',
                                '*.pyo', '.DS_Store')


def install(agent, project=None, user=False, dry_run=False):
    base = Path.home() if user else Path(project).expanduser().resolve(strict=True)
    if not base.is_dir():
        raise ValueError('Project must be an existing directory.')
    parent = base / AGENTS[agent] / 'skills'
    target = parent / 'doc-format'
    if target.resolve() == ROOT:
        raise ValueError('This is already the installed skill; no files changed.')
    for name in ('rules', 'profiles', 'scripts', 'agents', 'examples'):
        if target.resolve().is_relative_to(ROOT / name):
            raise ValueError('Choose a project outside the skill resources.')
    for name in ('SKILL.md', 'rules', 'profiles', 'scripts'):
        if not (ROOT / name).exists():
            raise ValueError('Incomplete skill package: ' + name)
    if target.exists() or target.is_symlink():
        raise FileExistsError('Skill destination already exists; nothing overwritten: ' + str(target))
    result = {'agent': agent, 'scope': 'user' if user else 'project',
              'destination': str(target), 'dry_run': dry_run}
    if dry_run:
        return result
    parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='.doc-format-install-', dir=parent) as temporary:
        stage = Path(temporary) / 'doc-format'
        stage.mkdir()
        for name in ITEMS:
            source = ROOT / name
            if source.is_dir():
                shutil.copytree(source, stage / name, ignore=IGNORED)
            elif source.is_file():
                shutil.copy2(source, stage / name)
        if target.exists() or target.is_symlink():
            raise FileExistsError('Skill destination appeared during installation; nothing overwritten.')
        stage.rename(target)
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    if sys.version_info < (3, 9):
        parser.error('Python 3.9 or newer is required.')
    parser.add_argument('--agent', required=True, choices=tuple(AGENTS))
    scope = parser.add_mutually_exclusive_group(required=True)
    scope.add_argument('--project', type=Path, help='Existing project directory')
    scope.add_argument('--user', action='store_true', help='Install for all local projects')
    parser.add_argument('--dry-run', action='store_true')
    args = parser.parse_args()
    try:
        result = install(args.agent, args.project, args.user, args.dry_run)
    except (OSError, ValueError) as error:
        print(str(error), file=sys.stderr)
        return 2
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == '__main__':
    sys.exit(main())
