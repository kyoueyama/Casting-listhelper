"""
Astream の特定キャンペーンの「リストアップ」Excel(.xlsx) を自動ダウンロードする。

  python download_astream.py --campaign "キャンペーン名"
  python download_astream.py --campaign "キャンペーン名" --headed   # ブラウザ表示
  python download_astream.py --campaign "キャンペーン名" --relogin  # 保存済みセッションを破棄

セレクタは画面のテキスト/役割ベース。UI が変わったら微調整する。
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv
from playwright.sync_api import (
    BrowserContext,
    Page,
    Playwright,
    TimeoutError as PWTimeout,
    sync_playwright,
)

ROOT = Path(__file__).resolve().parent
AUTH_STATE = ROOT / ".auth" / "state.json"
DOWNLOAD_DIR = ROOT / "downloads"


def safe_filename(name: str) -> str:
    return re.sub(r"[^\w\-.]+", "_", name).strip("_")


def login(page: Page, email: str, password: str, login_url: str) -> None:
    page.goto(login_url, wait_until="domcontentloaded")
    # ログイン画面以外（既にログイン済み）なら何もしない
    if "ログイン" not in page.content() and "login" not in page.url.lower():
        return
    page.get_by_placeholder("astream@astream.com").fill(email)
    page.get_by_placeholder(re.compile(r"\d+文字以上の半角英数字")).fill(password)
    page.get_by_role("button", name="ログイン").click()
    page.wait_for_load_state("networkidle")


def open_campaign(page: Page, campaign: str) -> None:
    """キャンペーン一覧で名前検索 → 該当行を開く"""
    # サイドナビの「キャンペーン」へ。link/button/任意のテキスト要素を順に試す
    page.wait_for_load_state("networkidle")
    nav_candidates = [
        page.get_by_role("link", name=re.compile("^キャンペーン$")),
        page.get_by_role("button", name=re.compile("^キャンペーン$")),
        page.locator("nav").get_by_text("キャンペーン", exact=True),
        page.get_by_text("キャンペーン", exact=True),
    ]
    clicked = False
    for cand in nav_candidates:
        try:
            cand.first.click(timeout=3000)
            clicked = True
            break
        except Exception:
            continue
    if not clicked:
        # 既に一覧にいる場合などはスキップ
        pass
    page.wait_for_load_state("networkidle")

    # キーワード検索ボックスにキャンペーン名を入れる
    search = page.get_by_placeholder(re.compile("キャンペーン名|会社名"))
    search.fill(campaign)
    # 検索 API のデバウンスを十分に待つ
    page.wait_for_timeout(1500)
    page.wait_for_load_state("networkidle")

    # 一覧から該当行のキャンペーン名リンク（テキスト一致）をクリック
    # ヘッダー行や別カラムの偶然マッチを避けるため、role=link/cell でテキスト一致を狙う
    target_candidates = [
        page.get_by_role("link", name=campaign, exact=True),
        page.get_by_role("cell", name=campaign, exact=True),
        page.get_by_text(campaign, exact=True),
    ]
    clicked = False
    for cand in target_candidates:
        try:
            cand.first.click(timeout=5000)
            clicked = True
            break
        except Exception:
            continue
    if not clicked:
        raise RuntimeError(
            f"キャンペーン '{campaign}' が一覧で見つかりませんでした。"
            f"検索結果0件か、表示名が完全一致していない可能性があります。"
        )

    # 詳細画面に遷移するまで待つ
    page.wait_for_url(re.compile(r"/campaign/detail/"), timeout=15_000)
    page.wait_for_load_state("networkidle")


def download_list_excel(page: Page, campaign: str) -> Path:
    """提案 > リストアップ で「⋮」→ リストダウンロード → エクセル形式 を実行"""

    # 「リストアップ」タブを選択（既に選択済みでも安全）
    try:
        page.get_by_role("tab", name="リストアップ").click(timeout=3000)
    except PWTimeout:
        pass

    # 三点メニュー: aria-label="more" + action-menu-button クラス
    page.locator("button.action-menu-button[aria-label='more']").first.click()

    # 「リストダウンロード」をホバー/クリックするとサブメニュー（エクセル形式 / csv形式）が出る
    page.get_by_role("menuitem", name=re.compile("リストダウンロード")).hover()
    page.wait_for_timeout(400)

    # サブメニューの「エクセル形式」を押すと、確認ダイアログが開く
    try:
        page.get_by_role("menuitem", name=re.compile(r"エクセル形式")).click()
    except Exception:
        page.get_by_text(re.compile(r"エクセル形式")).first.click()

    # 確認ダイアログ内の「ダウンロード」ボタンをクリックして実ダウンロードを発火
    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
    page.wait_for_timeout(400)
    with page.expect_download(timeout=180_000) as dl_info:
        page.get_by_role("button", name=re.compile(r"^ダウンロード$")).click()
    download = dl_info.value

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = DOWNLOAD_DIR / f"{safe_filename(campaign)}_{stamp}.xlsx"
    download.save_as(out)
    return out


def make_context(pw: Playwright, headed: bool, relogin: bool) -> BrowserContext:
    browser = pw.chromium.launch(headless=not headed)
    AUTH_STATE.parent.mkdir(parents=True, exist_ok=True)
    if AUTH_STATE.exists() and not relogin:
        return browser.new_context(storage_state=str(AUTH_STATE))
    return browser.new_context()


def main() -> int:
    load_dotenv(ROOT / ".env")
    parser = argparse.ArgumentParser()
    parser.add_argument("--campaign", required=True, help="ダウンロード対象のキャンペーン名")
    parser.add_argument("--headed", action="store_true", help="ブラウザを表示する")
    parser.add_argument("--relogin", action="store_true", help="保存済みログインを使わず再ログイン")
    args = parser.parse_args()

    email = os.environ.get("ASTREAM_EMAIL")
    password = os.environ.get("ASTREAM_PASSWORD")
    login_url = os.environ.get("ASTREAM_LOGIN_URL", "https://astream.jp/signin-cover")
    if not email or not password:
        print("ERROR: .env に ASTREAM_EMAIL / ASTREAM_PASSWORD を設定してください", file=sys.stderr)
        return 2

    with sync_playwright() as pw:
        ctx = make_context(pw, headed=args.headed, relogin=args.relogin)
        page = ctx.new_page()
        try:
            login(page, email, password, login_url)
            ctx.storage_state(path=str(AUTH_STATE))  # セッション保存
            open_campaign(page, args.campaign)
            out = download_list_excel(page, args.campaign)
            print(f"OK saved: {out}")
            return 0
        except Exception as e:
            # 失敗時は現在のページのスクショとHTMLを保存して原因特定しやすくする
            debug_dir = ROOT / "debug"
            debug_dir.mkdir(exist_ok=True)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            try:
                page.screenshot(path=str(debug_dir / f"error_{stamp}.png"), full_page=True)
                (debug_dir / f"error_{stamp}.html").write_text(page.content(), encoding="utf-8")
                print(f"DEBUG: saved screenshot/html under {debug_dir} (url={page.url})", file=sys.stderr)
            except Exception as inner:
                print(f"DEBUG: could not save debug artifacts: {inner}", file=sys.stderr)
            raise
        finally:
            ctx.close()


if __name__ == "__main__":
    sys.exit(main())
