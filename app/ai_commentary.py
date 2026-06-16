"""
KPIダッシュボード向け AIコメント生成（任意機能）

オフライン・単一exeという基本動作を壊さないため、この機能は完全にオプトインです:
  - GUIの「AIコメント設定」で有効化＋APIキーを保存（config.json）した場合に動作
  - CLI等でGUIを使わない場合は、環境変数 ANTHROPIC_API_KEY だけでも動作（後方互換）
  - キーが無い / 無効 / ネットが無い / API がエラー の場合は None を返すだけ
    （呼び出し側はコメント欄を出さず、従来通りのレポートを生成する）

モデルは claude-sonnet-4-6 を使用。
"""
import os

from app_config import load_config

MODEL = "claude-sonnet-4-6"

MONTH_NAMES = ["1月", "2月", "3月", "4月", "5月", "6月",
               "7月", "8月", "9月", "10月", "11月", "12月"]

_SYSTEM = (
    "あなたは小売業の売上データを分析するアナリストです。"
    "与えられた年間KPIをもとに、経営者向けの簡潔な日本語コメントを作成してください。\n"
    "要件:\n"
    "- 2〜3文、合計150字程度の文章で書く。\n"
    "- 傾向・注目点を述べ、最後に次の打ち手を1つ含める。\n"
    "- 数値は要点だけに絞り、冗長な説明や前置きは書かない。\n"
    "- 金額は与えられた『万円』単位の数値をそのまま使い、億・百万などへ単位変換しない。\n"
    "- 箇条書き・Markdown記法（#, *, -, ** 等）は使わず、本文のみを返す。"
)


def _man(value: int) -> str:
    """円 → 万円（四捨五入）に整形。大きな桁数による単位換算ミスを防ぐ。"""
    return f"{round(value / 10000):,}万円"


def _build_prompt(kpi: dict) -> str:
    monthly = "、".join(
        f"{MONTH_NAMES[i]}:{_man(v)}" for i, v in enumerate(kpi["monthly_sales"])
    )
    ranking = "、".join(
        f"{i + 1}位 {store}（{_man(sales)}）"
        for i, (store, sales) in enumerate(kpi["store_ranking"][:3])
    )
    mom = kpi["last_month_mom"]
    mom_str = f"{mom * 100:+.1f}%"
    return (
        "※金額はすべて『万円』単位。本文でもこの単位のまま使うこと。\n"
        f"対象年度: {kpi['year']}年\n"
        f"年間総売上（全店舗合計）: {_man(kpi['total_annual'])}\n"
        f"月別売上: {monthly}\n"
        f"最高月: {MONTH_NAMES[kpi['best_month'] - 1]}（{_man(kpi['best_sales'])}）\n"
        f"最低月: {MONTH_NAMES[kpi['worst_month'] - 1]}（{_man(kpi['worst_sales'])}）\n"
        f"月平均売上: {_man(kpi['avg_month'])}\n"
        f"店舗別年間売上ランキング: {ranking}\n"
        f"平均客単価: {kpi['avg_unit_price']:,}円\n"
        f"年間総客数: {kpi['total_customers']:,}人\n"
        f"直近の前月比（12月／11月）: {mom_str}\n"
    )


def generate_kpi_commentary(kpi: dict) -> str | None:
    """KPI辞書からAIコメントを生成して返す。

    APIキーが無い・anthropic未導入・通信失敗などの場合は None を返す
    （例外は送出せず、レポート生成を止めない）。
    """
    cfg = load_config()
    env_key = (os.environ.get("ANTHROPIC_API_KEY") or "").strip()
    key = (cfg.get("api_key") or "").strip() or env_key
    if not key:
        return None

    # 有効化判定: GUIトグルが優先。configが無い/未設定でも、環境変数があればCLI後方互換でオン。
    enabled = bool(cfg.get("ai_enabled", False)) or bool(env_key)
    if not enabled:
        return None

    try:
        import anthropic  # 遅延インポート: 未導入でもオフライン動作を壊さない

        client = anthropic.Anthropic(api_key=key)
        message = client.with_options(timeout=20.0, max_retries=1).messages.create(
            model=MODEL,
            max_tokens=512,
            system=_SYSTEM,
            messages=[{"role": "user", "content": _build_prompt(kpi)}],
        )
        text = next((b.text for b in message.content if b.type == "text"), "").strip()
        return text or None
    except Exception:
        # ネットワーク断・APIエラー・キー不正などはすべて無視してコメントなしに倒す
        return None
