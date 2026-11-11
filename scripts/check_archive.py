"""Verify the distributable archive, including macOS framework symlinks."""
from pathlib import Path
import platform
import subprocess
import tempfile
import zipfile

root = Path(__file__).resolve().parents[1]
archive, = (root / 'release').glob('*.zip')
with tempfile.TemporaryDirectory() as directory:
    if platform.system() == 'Darwin':
        subprocess.run(['ditto', '-xk', str(archive), directory], check=True)
        relative = 'JobFilter.app/Contents/MacOS/JobFilter'
    else:
        with zipfile.ZipFile(archive) as bundle:
            bundle.extractall(directory)
        relative = 'JobFilter/JobFilter.exe'
    executable = Path(directory) / archive.stem / relative
    subprocess.run([str(executable), '--self-test'], check=True, timeout=45)
print('Extracted desktop package startup verified.')
