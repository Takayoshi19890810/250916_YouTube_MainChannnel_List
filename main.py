# main.py

import os
import requests
import pandas as pd
from datetime import datetime, timedelta
import re

# GitHub ActionsのSecretsからAPIキーを取得
API_KEY = os.environ.get('YOUTUBE_API_KEY')

# スプレッドシートの代わりに、チャンネル情報をスクリプト内に定義
# ここに監視したいチャンネルの「チャンネル名」と「チャンネルID」を追加してください。
CHANNEL_DATA = [
    ['チャンネル名1', 'UCMuY6dM9l0c3o-C4R5pU3pQ'],
    ['チャンネル名2', 'UCg0s8PjYmD7U5R8tYQ1K5wA'],
    # 例：['Google Developers', 'UC_x5XG1OV2P6uZZ5FSM9Ttw'],
]

EXCEL_FILE = 'youtube_videos.xlsx'

# YouTube Data APIのベースURL
YOUTUBE_API_URL = 'https://www.googleapis.com/youtube/v3/'

def get_uploads_playlist_id(channel_id):
    """チャンネルIDからアップロード用プレイリストIDを取得する"""
    url = f'{YOUTUBE_API_URL}channels?part=contentDetails&id={channel_id}&key={API_KEY}'
    response = requests.get(url)
    response.raise_for_status()
    data = response.json()
    if 'items' not in data or not data['items']:
        raise ValueError(f"チャンネルID '{channel_id}' が見つかりません。")
    return data['items'][0]['contentDetails']['relatedPlaylists']['uploads']

def get_videos_from_playlist(playlist_id, cutoff_date):
    """プレイリストから指定期間内の動画を取得する"""
    videos = []
    page_token = ''
    while True:
        url = f'{YOUTUBE_API_URL}playlistItems?part=snippet&playlistId={playlist_id}&maxResults=50&pageToken={page_token}&key={API_KEY}'
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()

        for item in data.get('items', []):
            published_at = datetime.fromisoformat(item['snippet']['publishedAt'][:-1])
            if published_at < cutoff_date:
                return videos
            videos.append({
                'title': item['snippet']['title'],
                'videoId': item['snippet']['resourceId']['videoId'],
                'publishedAt': item['snippet']['publishedAt']
            })

        if 'nextPageToken' not in data:
            break
        page_token = data['nextPageToken']
    return videos

def get_video_details(video_ids):
    """動画IDリストから詳細情報を取得する"""
    details_map = {}
    for i in range(0, len(video_ids), 50):
        ids = ','.join(video_ids[i:i+50])
        url = f'{YOUTUBE_API_URL}videos?part=statistics,contentDetails&id={ids}&key={API_KEY}'
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()

        for item in data.get('items', []):
            video_id = item['id']
            stats = item.get('statistics', {})
            duration_iso = item.get('contentDetails', {}).get('duration', "PT0S")
            
            details_map[video_id] = {
                'viewCount': stats.get('viewCount', 0),
                'likeCount': stats.get('likeCount', 0),
                'commentCount': stats.get('commentCount', 0),
                'durationSeconds': parse_iso_duration(duration_iso)
            }
    return details_map

def parse_iso_duration(iso_duration):
    """ISO8601形式の再生時間を秒数に変換する"""
    match = re.match(r'PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?', iso_duration)
    if not match:
        return 0
    h, m, s = match.groups(default='0')
    return int(h) * 3600 + int(m) * 60 + int(s)

def format_duration(seconds):
    """秒数を HH:MM:SS 形式に変換する"""
    h, s = divmod(seconds, 3600)
    m, s = divmod(s, 60)
    return f'{h:02d}:{m:02d}:{s:02d}'

def calculate_engagement(views, likes, comments):
    """エンゲージメント係数を計算する"""
    views = int(views)
    likes = int(likes)
    comments = int(comments)
    if views == 0:
        return ''
    engagement = ((likes + comments) / views) * 100
    return f'{engagement:.2f}%'

def main():
    if not API_KEY:
        print("エラー: YouTube APIキーが設定されていません。")
        return

    # 既存のExcelファイルがあれば読み込み、なければ空のDataFrameを作成
    try:
        df_videos = pd.read_excel(EXCEL_FILE, index_col=None)
        existing_video_ids = set(df_videos['動画ID'])
        print(f"既存のExcelファイル '{EXCEL_FILE}' を読み込みました。")
    except FileNotFoundError:
        df_videos = pd.DataFrame(columns=['チャンネル名', 'タイトル', '投稿日', '動画ID', '再生時間', '再生数', 'コメント数', 'イイネ数', 'エンゲージメント係数', 'URL'])
        existing_video_ids = set()
        print(f"'{EXCEL_FILE}' が見つからないため、新しいファイルを作成します。")
    except Exception as e:
        print(f"Excelファイルの読み込み中に予期せぬエラーが発生しました: {e}")
        return

    # 過去2ヶ月間の動画を対象とする
    two_months_ago = datetime.now() - timedelta(days=60)
    new_rows = []

    for channel_name, channel_id in CHANNEL_DATA:
        print(f"▶ 処理中: {channel_name} ({channel_id})")
        try:
            uploads_id = get_uploads_playlist_id(channel_id)
            videos = get_videos_from_playlist(uploads_id, two_months_ago)

            if not videos:
                print(f"チャンネル '{channel_name}' に新しい動画が見つかりませんでした。")
                continue

            ids = [v['videoId'] for v in videos]
            details_map = get_video_details(ids)

            for v in videos:
                if v['videoId'] in existing_video_ids:
                    continue

                d = details_map.get(v['videoId'])
                if not d or d['durationSeconds'] < 60:
                    continue

                new_rows.append({
                    'チャンネル名': channel_name,
                    'タイトル': v['title'],
                    '投稿日': datetime.fromisoformat(v['publishedAt'][:-1]).strftime('%Y/%m/%d %H:%M:%S'),
                    '動画ID': v['videoId'],
                    '再生時間': format_duration(d['durationSeconds']),
                    '再生数': int(d['viewCount']),
                    'コメント数': int(d['commentCount']),
                    'イイネ数': int(d['likeCount']),
                    'エンゲージメント係数': calculate_engagement(d['viewCount'], d['likeCount'], d['commentCount']),
                    'URL': f'https://www.youtube.com/watch?v={v["videoId"]}'
                })
        except requests.exceptions.HTTPError as err:
            print(f"❌ API通信エラー [チャンネルID:{channel_id}]: {err}")
        except Exception as e:
            print(f"❌ 予期せぬエラー [チャンネルID:{channel_id}]: {e}")

    if new_rows:
        df_new = pd.DataFrame(new_rows)
        # 新しい動画を既存のデータフレームの先頭に追加
        df_videos = pd.concat([df_new, df_videos], ignore_index=True)
        try:
            df_videos.to_excel(EXCEL_FILE, index=False)
            print(f"{len(new_rows)}件の新しい動画をExcelファイルに追記しました。")
        except Exception as e:
            print(f"Excelファイルへの書き込み中にエラーが発生しました: {e}")
    else:
        print("新しい動画は見つかりませんでした。")

if __name__ == '__main__':
    main()
