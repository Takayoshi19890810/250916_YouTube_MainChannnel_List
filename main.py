# main.py

import os
import re
import requests
import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path

# ==== 設定 ====
API_KEY = os.environ.get("YOUTUBE_API_KEY")  # GitHub Actions Secrets から供給
EXCEL_FILE = "youtube_videos.xlsx"
YOUTUBE_API_URL = "https://www.googleapis.com/youtube/v3/"

# 監視対象チャンネル（名前, チャンネルID）
CHANNEL_DATA = [
    ["ワンソクTube", "UCo150kMjyLQDsLdudoyCqYg"],
    ["e-Carlife", "UCacmUS5IWcTzpI3b4ZkkSgw"],
    ["Lavecars TV", "UCtLo4nwb3ObCDZ4m8b8u7fA"],
    ["Driver channel", "UCup9EloQKxgDKvgJeKKZ85Q"],
    ["ベストカーweb", "UC7yk5_U7C0TuvYMqWzKyzkQ"],
    ["Ride now", "UC0P8fXzj-JxUsvDFEbpAkSg"],
    ["Kozzi TV", "UCaN_F80VfpzDN-vKn3IF4IQ"],  # 修正済み
]

# 取得対象期間（例：直近60日）
CUTOFF_DAYS = 60
# ショート等の除外（60秒未満は除外）
MIN_DURATION_SEC = 60


# ==== ユーティリティ ====
def parse_iso_duration(iso_duration: str) -> int:
    """ISO8601のPTxxHxxMxxSを秒に変換"""
    m = re.match(r"PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?", iso_duration or "")
    if not m:
        return 0
    h, mm, s = m.groups(default="0")
    return int(h) * 3600 + int(mm) * 60 + int(s)


def format_duration(seconds: int) -> str:
    """秒→HH:MM:SS"""
    h, s = divmod(int(seconds), 3600)
    m, s = divmod(s, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"


def calculate_engagement(views, likes, comments) -> str:
    """(いいね+コメント)/再生数 * 100 [%]"""
    try:
        v = int(views)
        l = int(likes)
        c = int(comments)
    except Exception:
        return ""
    if v <= 0:
        return ""
    return f"{((l + c) / v) * 100:.2f}%"


# ==== YouTube Data API ====
def get_uploads_playlist_id(channel_id: str) -> str | None:
    """チャンネルID→uploadsプレイリストID"""
    try:
        resp = requests.get(
            f"{YOUTUBE_API_URL}channels",
            params={"part": "contentDetails", "id": channel_id, "key": API_KEY},
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        items = data.get("items", [])
        if not items:
            print(f"❌ チャンネルID '{channel_id}' が見つかりません")
            return None
        return items[0]["contentDetails"]["relatedPlaylists"]["uploads"]
    except Exception as e:
        print(f"❌ チャンネル情報取得エラー: {e}")
        return None


def get_videos_from_playlist(playlist_id: str, cutoff_date: datetime) -> list[dict]:
    """
    uploadsプレイリストから動画の基礎情報を取得。
    cutoff_date より古い動画に達したら打ち切る。
    """
    results: list[dict] = []
    if not playlist_id:
        return results

    page_token = None
    while True:
        try:
            params = {
                "part": "snippet",
                "playlistId": playlist_id,
                "maxResults": 50,
                "key": API_KEY,
            }
            if page_token:
                params["pageToken"] = page_token

            resp = requests.get(f"{YOUTUBE_API_URL}playlistItems", params=params, timeout=30)
            resp.raise_for_status()
            data = resp.json()

            for item in data.get("items", []):
                snippet = item.get("snippet") or {}
                published_at_str = snippet.get("publishedAt")
                if not published_at_str:
                    continue

                # 例: '2025-09-01T12:34:56Z' → datetime
                published_at = datetime.fromisoformat(published_at_str.rstrip("Z"))
                if published_at < cutoff_date:
                    # 以降は古いと仮定して打ち切り
                    return results

                results.append(
                    {
                        "title": snippet.get("title", "タイトルなし"),
                        "videoId": (snippet.get("resourceId") or {}).get("videoId"),
                        "publishedAt": published_at_str,  # ISO文字列のまま保持
                    }
                )

            page_token = data.get("nextPageToken")
            if not page_token:
                break

        except Exception as e:
            print(f"❌ プレイリスト取得エラー: {e}")
            break

    return results


def get_video_details(video_ids: list[str]) -> dict[str, dict]:
    """動画ID群→ {videoId: {viewCount, likeCount, commentCount, durationSeconds}}"""
    details: dict[str, dict] = {}
    if not video_ids:
        return details

    for i in range(0, len(video_ids), 50):
        ids_chunk = list(filter(None, video_ids[i : i + 50]))
        if not ids_chunk:
            continue
        try:
            resp = requests.get(
                f"{YOUTUBE_API_URL}videos",
                params={
                    "part": "statistics,contentDetails",
                    "id": ",".join(ids_chunk),
                    "key": API_KEY,
                },
                timeout=30,
            )
            resp.raise_for_status()
            data = resp.json()

            for item in data.get("items", []):
                vid = item.get("id")
                stats = item.get("statistics", {}) or {}
                dur_iso = (item.get("contentDetails", {}) or {}).get("duration", "PT0S")
                details[vid] = {
                    "viewCount": stats.get("viewCount", 0),
                    "likeCount": stats.get("likeCount", 0),
                    "commentCount": stats.get("commentCount", 0),
                    "durationSeconds": parse_iso_duration(dur_iso),
                }
        except Exception as e:
            print(f"❌ 動画詳細取得エラー: {e}")

    return details


# ==== Excelの結合（既存ベースで新規だけ追記） ====
def append_to_excel_base_on_existing(base_xlsx: str, new_rows: pd.DataFrame) -> None:
    """
    既存Excelの内容をベースに new_rows を追記し、動画IDで重複排除して保存。
    公開日時があれば降順に並べ替え。
    """
    base_path = Path(base_xlsx)

    if base_path.exists():
        try:
            base_df = pd.read_excel(base_path)
        except Exception:
            base_df = pd.DataFrame()
    else:
        base_df = pd.DataFrame()

    # 列合わせ（欠けている列を補完）
    for col in set(new_rows.columns) - set(base_df.columns):
        base_df[col] = pd.NA
    for col in set(base_df.columns) - set(new_rows.columns):
        new_rows[col] = pd.NA

    merged = pd.concat([new_rows, base_df], ignore_index=True)  # 新着を先頭に
    if "動画ID" in merged.columns:
        merged.drop_duplicates(subset=["動画ID"], keep="first", inplace=True)

    # 公開日時の列候補
    for cand in ["投稿日", "publishedAt", "publishTime", "published_at"]:
        if cand in merged.columns:
            merged[cand] = pd.to_datetime(merged[cand], errors="coerce")
            merged.sort_values(cand, ascending=False, inplace=True)
            break

    merged.reset_index(drop=True, inplace=True)
    merged.to_excel(base_path, index=False)


# ==== メイン ====
def main():
    if not API_KEY:
        print("エラー: YouTube APIキーが設定されていません。")
        return

    # 既存Excelから既知の動画ID集合を作る（なければ空）
    try:
        df_existing = pd.read_excel(EXCEL_FILE)
        known_ids = set(df_existing.get("動画ID", pd.Series(dtype=str)).dropna().astype(str).tolist())
        print(f"既存Excel '{EXCEL_FILE}' を読み込みました。既知ID: {len(known_ids)}件")
    except FileNotFoundError:
        df_existing = pd.DataFrame(
            columns=[
                "チャンネル名",
                "タイトル",
                "投稿日",
                "動画ID",
                "再生時間",
                "再生数",
                "コメント数",
                "イイネ数",
                "エンゲージメント係数",
                "URL",
            ]
        )
        known_ids = set()
        print(f"'{EXCEL_FILE}' が見つからないため、新規作成します。")
    except Exception as e:
        print(f"Excel読み込みエラー: {e}")
        return

    cutoff = datetime.now() - timedelta(days=CUTOFF_DAYS)
    new_records = []

    for channel_name, channel_id in CHANNEL_DATA:
        print(f"▶ 処理中: {channel_name} ({channel_id})")
        uploads_id = get_uploads_playlist_id(channel_id)
        if not uploads_id:
            continue

        videos = get_videos_from_playlist(uploads_id, cutoff)
        if not videos:
            print(f"チャンネル '{channel_name}' に新着はありません。")
            continue

        ids = [v["videoId"] for v in videos if v.get("videoId")]
        details = get_video_details(ids)

        for v in videos:
            vid = v.get("videoId")
            if not vid or vid in known_ids:
                continue

            det = details.get(vid, {})
            dur = int(det.get("durationSeconds") or 0)
            if dur < MIN_DURATION_SEC:
                # 60秒未満は除外（ショート等）
                continue

            views = det.get("viewCount", 0)
            likes = det.get("likeCount", 0)
            comments = det.get("commentCount", 0)

            try:
                published_str = v.get("publishedAt", "")
                # 'YYYY-MM-DDTHH:MM:SSZ' → 'YYYY/MM/DD HH:MM:SS'
                published_fmt = (
                    datetime.fromisoformat(published_str.rstrip("Z")).strftime("%Y/%m/%d %H:%M:%S")
                    if published_str
                    else ""
                )
            except Exception:
                published_fmt = ""

            new_records.append(
                {
                    "チャンネル名": channel_name,
                    "タイトル": v.get("title", "タイトルなし"),
                    "投稿日": published_fmt,
                    "動画ID": vid,
                    "再生時間": format_duration(dur),
                    "再生数": int(views or 0),
                    "コメント数": int(comments or 0),
                    "イイネ数": int(likes or 0),
                    "エンゲージメント係数": calculate_engagement(views, likes, comments),
                    "URL": f"https://www.youtube.com/watch?v={vid}",
                }
            )

    if new_records:
        df_new = pd.DataFrame(new_records)
        append_to_excel_base_on_existing(EXCEL_FILE, df_new)
        print(f"✅ 新規 {len(df_new)} 件を追記しました。")
    else:
        print("新しい動画は見つかりませんでした。")


if __name__ == "__main__":
    main()
