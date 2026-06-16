"""
アプリ設定の読み書き（AIコメント機能のオン/オフ・APIキー）

設定は exe（開発時はプロジェクトルート）の隣の config.json に保存する。
ファイルが無ければ空dictを返し、呼び出し側は従来動作（環境変数のみ）に倒れる。
"""
import json
import sys
from pathlib import Path

_BASE = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent.parent
CONFIG_PATH = _BASE / "config.json"


def load_config() -> dict:
    """config.json を読み込む。存在しない・壊れている場合は {} を返す。"""
    if not CONFIG_PATH.exists():
        return {}
    try:
        data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_config(ai_enabled: bool, api_key: str, key_name: str | None = None) -> None:
    """AI設定を config.json に保存する。

    key_name は管理用メモ（任意・動作には不使用）。None のときは
    既存の config.json の値を引き継ぐ（GUI保存で手書きのメモが消えない）。
    """
    if key_name is None:
        key_name = load_config().get("key_name", "")
    CONFIG_PATH.write_text(
        json.dumps(
            {
                "ai_enabled": bool(ai_enabled),
                "key_name": (key_name or "").strip(),
                "api_key": (api_key or "").strip(),
            },
            ensure_ascii=False, indent=2,
        ),
        encoding="utf-8",
    )
