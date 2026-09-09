"""
Build script to compile PDF2text into a standalone Windows executable.
Usage:
    python build_exe.py
"""

import os
import subprocess
import sys

def build():
    print("=" * 60)
    print("PDF2text - Building Windows Executable")
    print("=" * 60)

    spec_file = "PDF2text.spec"
    if not os.path.isfile(spec_file):
        print(f"Error: Spec file '{spec_file}' not found.")
        sys.exit(1)

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        spec_file
    ]

    print(f"Running command: {' '.join(cmd)}")
    result = subprocess.run(cmd)

    if result.returncode == 0:
        exe_path = os.path.abspath(os.path.join("dist", "PDF2text", "PDF2text.exe"))
        print("\n" + "=" * 60)
        print("[OK] Build completed successfully!")
        print(f"Standalone executable located at:\n  {exe_path}")
        print("=" * 60)
    else:
        print("\n" + "=" * 60)
        print(f"[ERROR] Build failed with exit code: {result.returncode}")
        print("=" * 60)
        sys.exit(result.returncode)

if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    build()
