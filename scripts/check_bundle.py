"""Verify the packaged executable can initialize its desktop UI."""
from pathlib import Path
import platform
import subprocess
root = Path(__file__).resolve().parents[1]
path = root / 'dist' / ('JobFilter.app/Contents/MacOS/JobFilter' if platform.system() == 'Darwin' else 'JobFilter/JobFilter.exe')
subprocess.run([str(path), '--self-test'], check=True, timeout=45)
print('Packaged application startup verified.')
