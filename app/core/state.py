"""
Persistência simples do "último ID cadastrado no EasyApp".
Usa um arquivo JSON local (state.json).
"""
from __future__ import annotations

import json
from pathlib import Path

STATE_FILE = Path(__file__).resolve().parent.parent.parent / "state.json"


def carregar_estado() -> dict:
    if STATE_FILE.exists():
        try:
            return json.loads(STATE_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def salvar_estado(estado: dict) -> None:
    STATE_FILE.write_text(json.dumps(estado, indent=2, ensure_ascii=False), encoding="utf-8")


def get_ultimo_id() -> int | None:
    estado = carregar_estado()
    v = estado.get("ultimo_id_easyapp")
    try:
        return int(v) if v is not None else None
    except Exception:
        return None


def set_ultimo_id(valor: int) -> None:
    estado = carregar_estado()
    estado["ultimo_id_easyapp"] = int(valor)
    salvar_estado(estado)
