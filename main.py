name: YouTube fetch hourly (Release)

on:
  schedule:
    - cron: "0 * * * *"   # 毎時00分（UTC基準。JSTでも毎時00分に走る）
  workflow_dispatch:

permissions:
  contents: write   # リリース作成・更新に必要

jobs:
  run:
    runs-on: ubuntu-latest

    steps:
      - name: Checkout repo
        uses: actions/checkout@v4

      - name: Set up Python
        uses: actions/setup-python@v5
        with:
          python-version: "3.11"

      - name: Install deps
        run: |
          python -m pip install --upgrade pip
          pip install -r requirements.txt

      # 直前のリリース資産(youtube_videos.xlsx)をローカルへ取得
      # 初回などリリースが無い場合はスキップ
      - name: Download latest release asset
        uses: dawidd6/action-download-release@v2
        with:
          tag: youtube-videos-latest
          file: youtube_videos.xlsx
          skip_if_no_release: true
          # not_found_behavior: warn  # v2ではskip_if_no_releaseで十分

      - name: Run Python script
        env:
          YOUTUBE_API_KEY: ${{ secrets.YOUTUBE_API_KEY }}
        run: |
          python main.py

      # 更新済みの youtube_videos.xlsx を「youtube-videos-latest」タグのリリースに上書き添付
      # リリースが無ければ作成、あれば上書き（overwrite: true）
      - name: Upload release asset
        uses: svenstaro/upload-release-action@v2
        with:
          repo_token: ${{ secrets.GITHUB_TOKEN }}
          tag: youtube-videos-latest
          release_name: "YouTube Videos (latest)"
          file: youtube_videos.xlsx
          overwrite: true
          body: "Hourly update: ${{ github.run_id }} / $(date -u +'%Y-%m-%d %H:%M:%SZ')"
