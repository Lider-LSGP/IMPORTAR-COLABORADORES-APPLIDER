"""
Entry-point exigido pelo Streamlit Cloud (raiz do repositório).
Apenas delega para `app/streamlit_app.py`.
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
APP_DIR = ROOT / "app"
sys.path.insert(0, str(APP_DIR))

# Executa o app principal
exec(compile((APP_DIR / "streamlit_app.py").read_text(encoding="utf-8"),
             str(APP_DIR / "streamlit_app.py"), "exec"))
