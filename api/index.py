"""Entrada alternativa; Vercel FastAPI usa main.py como entrypoint principal."""
import sys
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
from main import app  # noqa: E402
