"""Build a portable source release, excluding local credentials and mailbox data."""
from pathlib import Path
import zipfile

ROOT = Path(__file__).resolve().parent
RELEASE_FILES = (
    'README.md', '.env.example', '.gitignore', 'requirements.txt',
    'requirements-lock.txt', 'hr_digest.py', 'setup.ps1', 'install_task.ps1',
    'build_release.py',
)


def build_release(root=ROOT):
    root = Path(root).resolve()
    sources = [root / name for name in RELEASE_FILES]
    sources.extend(sorted((root / 'tests').glob('test_*.py')))
    for source in sources:
        if source.is_symlink() or not source.is_file() or root not in source.resolve().parents:
            raise ValueError('Missing or unsafe release source: ' + source.name)
    target = root / 'dist' / 'hotmail-hr-digest.zip'
    target.parent.mkdir(exist_ok=True)
    with zipfile.ZipFile(str(target), 'w', compression=zipfile.ZIP_DEFLATED) as archive:
        for source in sources:
            archive.write(str(source), 'hotmail-hr-digest/' + source.relative_to(root).as_posix())
    return target


if __name__ == '__main__':
    print(build_release())
