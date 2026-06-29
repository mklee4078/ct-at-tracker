"""
crawl_influencer.py
유튜브 브랜디드 / 인스타그램 릴스 인플루언서 데이터 수집
- 유튜브: 구독자수, 일반영상 평균조회수/댓글수(최근 7개), 숏폼 평균조회수/댓글수(최근 7개)
- 인스타: 팔로워수, 릴스 평균조회수/좋아요수/댓글수(최근 7개)
결과 → Google Sheets 업데이트

GitHub Actions 환경에서는 아래 2개를 환경변수(Secrets)로 주입받음:
  - YOUTUBE_API_KEY
  - GOOGLE_SERVICE_ACCOUNT_JSON (서비스 계정 키 파일 전체 내용, JSON 문자열)
"""

import os
import json
import re
import time
from datetime import datetime, timezone, timedelta

import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import instaloader

# ─────────────────────────────────────────
# 설정 (환경변수 우선, 없으면 로컬 기본값)
# ─────────────────────────────────────────
YOUTUBE_API_KEY = os.environ.get("YOUTUBE_API_KEY", "")
SPREADSHEET_ID  = os.environ.get("SPREADSHEET_ID", "1b7cXCVTTiSBpo-FzmS_R1QBatlAk32zGjxXO_xN6rO8")
SHEET_NAME      = os.environ.get("SHEET_NAME", "자동 수집")

# 서비스 계정 키: 환경변수(JSON 문자열) 또는 로컬 파일 경로
SERVICE_ACCOUNT_JSON_STR = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "")
SERVICE_ACCOUNT_FILE_LOCAL = r"G:\공유 드라이브\영업마케팅실\1.온라인\99.개인폴더\이민경\클로드\유튜버,인플루언서 트래킹\runrun-491210-51b9b77f8e0a.json"

# 컬럼 인덱스 (0-based)
COL_NAME      = 0   # A: 인플루언서
COL_TYPE      = 1   # B: 컨텐츠 유형
COL_INSTA_ID  = 2   # C: 인스타그램 아이디
COL_YT_URL    = 3   # D: 유튜브 채널 URL
COL_YT_SUBS   = 4   # E: 구독자수
COL_YT_VW     = 5   # F: 일반영상 평균조회수(최근7개)
COL_YT_CMT    = 6   # G: 일반영상 평균댓글수(최근7개)
COL_YT_SVW    = 7   # H: 숏폼 평균조회수(최근7개)
COL_YT_SCMT   = 8   # I: 숏폼 평균댓글수(최근7개)
COL_IG_FOL    = 9   # J: 팔로워수
COL_IG_VW     = 10  # K: 릴스 평균조회수(최근7개)
COL_IG_LIKE   = 11  # L: 릴스 평균좋아요수(최근7개)
COL_IG_CMT    = 12  # M: 릴스 평균댓글수(최근7개)
COL_UPDATED   = 13  # N: 업데이트 시각

DATA_START_ROW = 2  # 헤더=1행, 데이터는 2행부터
MAX_RECENT_ITEMS = 7  # "최근 7개" 기준

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


# ─────────────────────────────────────────
# Google Sheets 연결
# ─────────────────────────────────────────
def get_credentials():
    if SERVICE_ACCOUNT_JSON_STR:
        info = json.loads(SERVICE_ACCOUNT_JSON_STR)
        return Credentials.from_service_account_info(info, scopes=SCOPES)
    else:
        return Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE_LOCAL, scopes=SCOPES)


def get_sheets_service():
    creds = get_credentials()
    return build("sheets", "v4", credentials=creds)


def read_sheet_values(service):
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!A2:N1000",
    ).execute()
    return result.get("values", [])


def write_row(service, row_index, values):
    """row_index: 0-based offset from DATA_START_ROW"""
    sheet_row = DATA_START_ROW + row_index
    range_str = f"{SHEET_NAME}!E{sheet_row}:N{sheet_row}"
    body = {"values": [values]}
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=range_str,
        valueInputOption="RAW",
        body=body,
    ).execute()


# ─────────────────────────────────────────
# YouTube 유틸
# ─────────────────────────────────────────
def extract_channel_id(youtube, url):
    """URL에서 채널 ID 추출 (@핸들 / channel / user 형식 모두 지원)"""
    url = url.strip()

    channel_match = re.search(r"/channel/([a-zA-Z0-9_-]+)", url)
    if channel_match:
        return channel_match.group(1)

    handle_match = re.search(r"/@([^/?\s]+)", url)
    if handle_match:
        handle = handle_match.group(1)
        res = youtube.channels().list(part="id", forHandle=handle).execute()
        items = res.get("items", [])
        if items:
            return items[0]["id"]
        # fallback: search
        res = youtube.search().list(part="snippet", q=handle, type="channel", maxResults=1).execute()
        items = res.get("items", [])
        if items:
            return items[0]["snippet"]["channelId"]

    user_match = re.search(r"/user/([^/?\s]+)", url)
    if user_match:
        res = youtube.channels().list(part="id", forUsername=user_match.group(1)).execute()
        items = res.get("items", [])
        if items:
            return items[0]["id"]

    return None


def get_channel_stats(youtube, channel_id):
    res = youtube.channels().list(part="statistics,contentDetails", id=channel_id).execute()
    items = res.get("items", [])
    if not items:
        return None, None
    subs = int(items[0]["statistics"].get("subscriberCount", 0))
    uploads_playlist = items[0]["contentDetails"]["relatedPlaylists"]["uploads"]
    return subs, uploads_playlist


def get_recent_video_ids(youtube, uploads_playlist_id, max_items=30):
    res = youtube.playlistItems().list(
        part="contentDetails",
        playlistId=uploads_playlist_id,
        maxResults=max_items,
    ).execute()
    return [item["contentDetails"]["videoId"] for item in res.get("items", [])]


def get_videos_details(youtube, video_ids):
    """50개씩 배치로 contentDetails + statistics 조회"""
    details = []
    for i in range(0, len(video_ids), 50):
        batch = video_ids[i:i + 50]
        res = youtube.videos().list(
            part="contentDetails,statistics",
            id=",".join(batch),
        ).execute()
        details.extend(res.get("items", []))
    return details


def parse_duration_seconds(duration_str):
    """ISO 8601 duration (e.g. PT1M30S) -> seconds, without isodate dependency"""
    match = re.match(r"PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?", duration_str)
    if not match:
        return 0
    h, m, s = (int(x) if x else 0 for x in match.groups())
    return h * 3600 + m * 60 + s


def calc_youtube_metrics(youtube, channel_url):
    channel_id = extract_channel_id(youtube, channel_url)
    if not channel_id:
        return None

    subs, uploads_playlist = get_channel_stats(youtube, channel_id)
    if uploads_playlist is None:
        return None

    video_ids = get_recent_video_ids(youtube, uploads_playlist, max_items=30)
    details = get_videos_details(youtube, video_ids)

    shorts = []
    regular = []
    for item in details:
        duration = parse_duration_seconds(item["contentDetails"]["duration"])
        stats = item.get("statistics", {})
        views = int(stats.get("viewCount", 0))
        comments = int(stats.get("commentCount", 0))
        entry = {"views": views, "comments": comments}
        if duration <= 60:
            shorts.append(entry)
        else:
            regular.append(entry)

    def avg(lst, key, n=MAX_RECENT_ITEMS):
        recent = lst[:n]
        if not recent:
            return 0
        return round(sum(x[key] for x in recent) / len(recent))

    return {
        "subs": subs,
        "regular_avg_views": avg(regular, "views"),
        "regular_avg_comments": avg(regular, "comments"),
        "shorts_avg_views": avg(shorts, "views"),
        "shorts_avg_comments": avg(shorts, "comments"),
    }


# ─────────────────────────────────────────
# Instagram 유틸 (instaloader)
# ─────────────────────────────────────────
def calc_instagram_metrics(loader, username):
    try:
        profile = instaloader.Profile.from_username(loader.context, username)
    except Exception as e:
        print(f"  [IG] 프로필 조회 실패 ({username}): {e}")
        return None

    followers = profile.followers

    reels = []
    try:
        for post in profile.get_posts():
            if post.is_video and post.typename in ("GraphVideo", "GraphSidecar") or post.is_video:
                reels.append({
                    "views": post.video_view_count or 0,
                    "likes": post.likes or 0,
                    "comments": post.comments or 0,
                })
            if len(reels) >= MAX_RECENT_ITEMS:
                break
    except Exception as e:
        print(f"  [IG] 게시물 조회 실패 ({username}): {e}")

    def avg(key):
        if not reels:
            return 0
        return round(sum(r[key] for r in reels) / len(reels))

    return {
        "followers": followers,
        "reels_avg_views": avg("views"),
        "reels_avg_likes": avg("likes"),
        "reels_avg_comments": avg("comments"),
    }


# ─────────────────────────────────────────
# 메인
# ─────────────────────────────────────────
def main():
    print("=== 인플루언서 데이터 수집 시작 ===")
    service = get_sheets_service()
    rows = read_sheet_values(service)
    print(f"총 {len(rows)}개 행 로드")

    youtube = build("youtube", "v3", developerKey=YOUTUBE_API_KEY) if YOUTUBE_API_KEY else None
    ig_loader = instaloader.Instaloader(quiet=True, download_pictures=False,
                                         download_videos=False, download_video_thumbnails=False,
                                         save_metadata=False, compress_json=False)

    kst = timezone(timedelta(hours=9))

    for idx, row in enumerate(rows):
        name = row[COL_NAME] if len(row) > COL_NAME else ""
        if not name:
            continue

        yt_url = row[COL_YT_URL] if len(row) > COL_YT_URL else ""
        ig_id = row[COL_INSTA_ID] if len(row) > COL_INSTA_ID else ""

        print(f"\n[{idx+1}/{len(rows)}] {name}")

        yt_metrics = None
        if yt_url and youtube:
            try:
                yt_metrics = calc_youtube_metrics(youtube, yt_url)
                print(f"  YT 구독자: {yt_metrics['subs'] if yt_metrics else 'N/A'}")
            except Exception as e:
                print(f"  [YT] 오류: {e}")

        ig_metrics = None
        if ig_id:
            try:
                ig_metrics = calc_instagram_metrics(ig_loader, ig_id)
                print(f"  IG 팔로워: {ig_metrics['followers'] if ig_metrics else 'N/A'}")
                time.sleep(2)  # rate limit 완화
            except Exception as e:
                print(f"  [IG] 오류: {e}")

        values = [
            yt_metrics["subs"] if yt_metrics else "",
            yt_metrics["regular_avg_views"] if yt_metrics else "",
            yt_metrics["regular_avg_comments"] if yt_metrics else "",
            yt_metrics["shorts_avg_views"] if yt_metrics else "",
            yt_metrics["shorts_avg_comments"] if yt_metrics else "",
            ig_metrics["followers"] if ig_metrics else "",
            ig_metrics["reels_avg_views"] if ig_metrics else "",
            ig_metrics["reels_avg_likes"] if ig_metrics else "",
            ig_metrics["reels_avg_comments"] if ig_metrics else "",
            datetime.now(kst).strftime("%Y-%m-%d %H:%M:%S"),
        ]

        try:
            write_row(service, idx, values)
        except Exception as e:
            print(f"  [Sheets] 쓰기 오류: {e}")

        time.sleep(1)

    print("\n=== 완료 ===")


if __name__ == "__main__":
    main()
