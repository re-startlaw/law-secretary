"""送信者メアド確定分類（classify_by_sender_email）と
サブフォルダ選択リファクタの等価性テスト。実 Drive 非依存。

実行: venv/bin/python -m pytest tests/test_classify_sender.py
"""
import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import secretary as s


@pytest.fixture
def restore_sender_map():
    """SENDER_EMAIL_TO_CLIENT をテスト後に元へ戻す。"""
    saved = dict(s.SENDER_EMAIL_TO_CLIENT)
    yield
    s.SENDER_EMAIL_TO_CLIENT.clear()
    s.SENDER_EMAIL_TO_CLIENT.update(saved)


def _client(folder_name, client_name, path, case_type="unknown"):
    return {
        "folder_name": folder_name,
        "client_name": client_name,
        "path": path,
        "case_type": case_type,
    }


# ── classify_by_sender_email：送信者で依頼者が確定する ──────────────

def test_sender_email_resolves_ma(tmp_path):
    """martin.ma@letour.co.jp の非経理ファイル → ま_馬強（サブフォルダ語なし→直下）。"""
    ma = _client("ま_馬強", "ま_馬強", str(tmp_path / "ま_馬強"))
    os.makedirs(ma["path"], exist_ok=True)
    result = s.classify_by_sender_email(
        "260612_Angelina出生証明書.pdf",
        [ma],
        sender='"martin.ma letour.co.jp" <martin.ma@letour.co.jp>',
        subject="回复: 学校側対応について",
    )
    assert result is not None
    dest, _rel = result
    assert dest == ma["path"]


def test_sender_email_resolves_pham(tmp_path):
    """nao@teruiglobal.jp → ふ_ファム・ティ・フォン。"""
    pham = _client(
        "ふ_ファム・ティ・フォン", "ふ_ファム・ティ・フォン",
        str(tmp_path / "ふ_ファム・ティ・フォン"),
    )
    os.makedirs(pham["path"], exist_ok=True)
    result = s.classify_by_sender_email(
        "260616_嘆願書（塩野社長）.docx",
        [pham],
        sender='"行政書士 照井奈央" <nao@teruiglobal.jp>',
        subject="",
    )
    assert result is not None
    dest, _rel = result
    assert dest == pham["path"]


def test_unknown_sender_returns_none(tmp_path):
    """未知送信者 → None（送信者では確定しない）。"""
    ma = _client("ま_馬強", "ま_馬強", str(tmp_path / "ま_馬強"))
    result = s.classify_by_sender_email(
        "なにか.pdf", [ma], sender="unknown@example.com", subject="",
    )
    assert result is None


def test_empty_sender_returns_none(tmp_path):
    ma = _client("ま_馬強", "ま_馬強", str(tmp_path / "ま_馬強"))
    assert s.classify_by_sender_email("x.pdf", [ma], sender="", subject="") is None


# ── M-1：client_name ≠ folder_name（全角スペース依頼者）でも解決 ─────

def test_fullwidth_space_client_via_client_name(tmp_path, restore_sender_map):
    """手動値(client_name 表記) で全角スペース依頼者を解決できる。"""
    tam = _client("た　田村正宣", "田村正宣", str(tmp_path / "tam"), "criminal")
    os.makedirs(tam["path"], exist_ok=True)
    s.SENDER_EMAIL_TO_CLIENT["someone@example.com"] = "田村正宣"  # client_name 表記
    result = s.classify_by_sender_email(
        "260601_パスポート.pdf", [tam], sender="someone@example.com", subject="",
    )
    assert result is not None
    assert result[0] == tam["path"]


def test_fullwidth_space_client_via_folder_name(tmp_path, restore_sender_map):
    """学習値(folder_name 表記) で全角スペース依頼者を解決できる。"""
    tam = _client("た　田村正宣", "田村正宣", str(tmp_path / "tam"), "criminal")
    os.makedirs(tam["path"], exist_ok=True)
    s.SENDER_EMAIL_TO_CLIENT["someone@example.com"] = "た　田村正宣"  # folder_name 表記
    result = s.classify_by_sender_email(
        "260601_パスポート.pdf", [tam], sender="someone@example.com", subject="",
    )
    assert result is not None
    assert result[0] == tam["path"]


# ── 手動固定が学習を上書きしない（後勝ち）───────────────────────

def test_manual_overrides_learned(restore_sender_map):
    """学習で martin.ma が別依頼者に紐づいても、手動後勝ちで ま_馬強 が残る。"""
    # main() と同じ順序を再現
    learned = {"martin.ma@letour.co.jp": "い_インバウンドサポート学費訴訟"}
    s.SENDER_EMAIL_TO_CLIENT.update(learned)
    s.SENDER_EMAIL_TO_CLIENT.update(s._MANUAL_SENDER_EMAIL_TO_CLIENT)
    assert s.SENDER_EMAIL_TO_CLIENT["martin.ma@letour.co.jp"] == "ま_馬強"


# ── _select_case_subfolder：リファクタ等価性 ────────────────────

def test_subfolder_criminal(tmp_path):
    c = _client("た_x", "た_x", str(tmp_path), "criminal")
    assert s._select_case_subfolder(c, s._norm_text("260601_勾留請求")) == "02_身体拘束関係"


def test_subfolder_civil(tmp_path):
    c = _client("ま_y", "ま_y", str(tmp_path), "civil")
    assert s._select_case_subfolder(c, s._norm_text("260601_甲号証")) == "01甲号証"


def test_subfolder_unknown_existing_dir(tmp_path):
    """unknown 案件：実在するサブフォルダのみ採用（R-1）。"""
    c = _client("ま_馬強", "ま_馬強", str(tmp_path), "unknown")
    os.makedirs(os.path.join(c["path"], "06資料"), exist_ok=True)
    assert s._select_case_subfolder(c, s._norm_text("260601_資料")) == "06資料"


def test_subfolder_unknown_missing_dir(tmp_path):
    """unknown 案件：サブフォルダが実在しなければ None（直下行き）。"""
    c = _client("ま_馬強", "ま_馬強", str(tmp_path), "unknown")
    # 06資料 を作らない
    assert s._select_case_subfolder(c, s._norm_text("260601_資料")) is None


# ── classify_file 全体経路 ──────────────────────────────────────

def test_classify_file_sender_wins(tmp_path):
    ma = _client("ま_馬強", "ま_馬強", str(tmp_path / "ま_馬強"))
    os.makedirs(ma["path"], exist_ok=True)
    dest, _rel = s.classify_file(
        "260612_Angelina出生証明書.pdf", [ma],
        sender="<martin.ma@letour.co.jp>", subject="回复: 学校側対応について",
    )
    assert dest == ma["path"]


def test_classify_file_unknown_sender_no_match(tmp_path):
    ma = _client("ま_馬強", "ま_馬強", str(tmp_path / "ま_馬強"))
    os.makedirs(ma["path"], exist_ok=True)
    dest, rel = s.classify_file(
        "謎のファイル.pdf", [ma], sender="stranger@example.com", subject="",
    )
    assert dest == s.UNKNOWN_FOLDER
    assert "分からなかった" in rel


# ── _llm_case_dest_allowed：LLM事件記録誤割当ガード ──────────────

def _tamura_client(path="/x/た_田村正宣"):
    return _client("た_田村正宣", "田村正宣", path, case_type="criminal")


def test_guard_rejects_case_dest_when_client_absent(monkeypatch):
    """税金書類（田村名なし）を 田村/02 に振ろうとしたら拒否される。"""
    monkeypatch.setitem(s.CLIENT_ALIASES, "田村正宣", ["田村"])
    allowed = s._llm_case_dest_allowed(
        "01_事件記録/た_田村正宣/02_身体拘束関係",
        [_tamura_client()],
        "260629_あなたの特別区民税・都民税、森林環境税を.pdf 米谷尚起 特別区民税納税通知書",
    )
    assert allowed is False


def test_guard_allows_case_dest_when_client_mentioned(monkeypatch):
    """田村氏の謄写代領収証（本文に田村あり）は通す。"""
    monkeypatch.setitem(s.CLIENT_ALIASES, "田村正宣", ["田村"])
    allowed = s._llm_case_dest_allowed(
        "01_事件記録/た_田村正宣/12_精算？",
        [_tamura_client()],
        "260622_領収証.pdf 東京謄写センター 但 田村 正宣氏の謄写代として",
    )
    assert allowed is True


def test_guard_allows_client_matched_by_alias(monkeypatch):
    """別名（会社名等）でヒットする場合も通す。"""
    monkeypatch.setitem(s.CLIENT_ALIASES, "田村正宣", ["田村", "有限会社タナカ"])
    allowed = s._llm_case_dest_allowed(
        "01_事件記録/た_田村正宣/05_検察官証拠",
        [_tamura_client()],
        "260629_取引明細.pdf 有限会社タナカ 帳簿",
    )
    assert allowed is True


def test_guard_ignores_non_case_dest():
    """経理など事件記録以外の保存先は常に許可（対象外）。"""
    assert s._llm_case_dest_allowed(
        "03_経理/通信費", [_tamura_client()], "NTTファイナンス 米谷尚起",
    ) is True


def test_guard_passes_when_client_folder_unresolvable():
    """rel の依頼者フォルダが一覧に無いときは保守的に通す。"""
    assert s._llm_case_dest_allowed(
        "01_事件記録/ん_終了/ぐ_グエン/連絡文書", [_tamura_client()], "何か",
    ) is True
