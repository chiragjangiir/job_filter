"""Package desktop build, documentation, synthetic samples and checksum."""
import hashlib
from pathlib import Path
import platform
import shutil
import tempfile
import os
import stat
import zipfile

root = Path(__file__).resolve().parents[1]
system = platform.system().lower()
arch = platform.machine().lower().replace('amd64', 'x86_64')
name = f'JobFilter-{system}-{arch}'
release = root / 'release'
release.mkdir(exist_ok=True)
with tempfile.TemporaryDirectory() as tmp:
    stage = Path(tmp) / name
    stage.mkdir()
    binary = root / 'dist' / ('JobFilter.app' if system == 'darwin' else 'JobFilter')
    if not binary.is_dir():
        raise SystemExit(f'Expected desktop bundle not found: {binary.name}')
    shutil.copytree(binary, stage / binary.name, symlinks=True)
    for file in ['README.md', 'README.en.md', 'LICENSE', 'CHANGELOG.md']:
        shutil.copy2(root / file, stage / file)
    shutil.copytree(root / 'examples', stage / 'examples')
    shutil.copytree(root / 'docs', stage / 'docs')
    archive = release / f'{name}.zip'
    with zipfile.ZipFile(archive, 'w', compression=zipfile.ZIP_DEFLATED) as output:
        for path in sorted(stage.rglob('*')):
            relative = path.relative_to(tmp).as_posix()
            if path.is_symlink():
                item = zipfile.ZipInfo(relative)
                item.create_system = 3
                item.external_attr = (stat.S_IFLNK | 0o777) << 16
                output.writestr(item, os.readlink(path))
            else:
                output.write(path, relative)
digest = hashlib.sha256(archive.read_bytes()).hexdigest()
archive.with_suffix('.zip.sha256').write_text(f'{digest}  {archive.name}\n')
print(f'Packaged {archive.name} with SHA-256 checksum.')
