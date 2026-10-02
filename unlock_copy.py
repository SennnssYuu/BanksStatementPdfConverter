"""
Save an unlocked copy of a password-protected PDF (for checking/debugging).

Usage:
    python unlock_copy.py "C:\path\to\statement.pdf"

The password is typed hidden and never saved. The copy is written next to the
original as "<name> - unlocked.pdf". Delete it when you're done with it.
"""

import getpass
import sys
from pathlib import Path

from pypdf import PdfReader, PdfWriter


def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    src = Path(sys.argv[1])
    reader = PdfReader(src)
    if reader.is_encrypted and not reader.decrypt(getpass.getpass("PDF password (input hidden): ")):
        sys.exit("Wrong password.")
    out = src.with_name(f"{src.stem} - unlocked.pdf")
    PdfWriter(clone_from=reader).write(out)
    print(f"Saved {out}")


if __name__ == "__main__":
    main()
