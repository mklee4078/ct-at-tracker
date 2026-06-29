"""
crawl_influencer.py
유튜브 브랜디드 / 인스타그램 릴스 인플루언서 데이터 수집
공유 드라이브 지원 버전 (Google Sheets API 직접 호출)
"""

from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta, timezone
import isodate
import instaloader
import time
import re

# ─────────────────────────────────────────
# 설정
# ─────────────────────────────────────────
SERVICE_ACCOUNT_FILE = r"G:\공유 드라이브\영업마케팅실\1.온라인\99.개인폴더\이민경\클로드\유튜버,인플루언서 트래킹\runrun-491210-51b9b77f8e0a.json"
YOUTUBE_API_KEY      = "AIzaSyASxfKhoQMmNrwZd0JpAjNMSEYAubERhsE"
SPREADSHEET_ID       = "1aRm6LWTWOkPuMU6A82uYVePHD24uywiEmWXERmgF-tU"
SHEET_NAME           = "구독자수 체크"

COL_NAME      = 0
COL_TYPE      = 1
COL_INSTA_ID  = 2
COL_YT_URL    = 3
COL_YT_SUBS   = 4
COL_YT_VW     = 5
COL_YT_CMT    = 6
COL_YT_SVW    = 7
COL_YT_SCMT   = 8
COL_IG_FOL    = 9
COL_IG_VW     = 10
COL_IG_LIKE   = 11
COL_IG_CMT    = 12
COL_UPDATED   = 13

DATA_START_ROW = 2

# ─────────────────────────────────────────
# Google Sheets API 직접 호출 (공유 드라이브 지원)
# ─────────────────────────────────────────
def get_sheets_service():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
    return build("sheets", "v4", credentials=creds)

def read_sheet(service):
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!A:N",
        valueRenderOption="UNFORMATTED_VALUE",
    ).execute()
    return result.get("values", [])

def write_row(service, row_index, values):
    col_end = chr(ord('A') + len(values) - 1)
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!A{row_index}:{col_end}{row_index}",
        valueInputOption="USER_ENTERED",
        body={"values": [values]},
    ).execute()

def ensure_headers(service):
    rows = read_sheet(service)
    if rows and len(rows[0]) >= 5 and rows[0][4] == "구독자수":
        return
    headers = [
        "인플루언서", "컨텐츠", "인스타그램 아이디", "유튜브 채널 URL",
        "구독자수", "일반영상 평균조회수(최근7건)", "일반영상 평균댓글수(최근7건)",
        "숏폼 평균조회수(최근7건)", "숏폼 평균댓글수(최근7건)",
        "팔로워수", "릴스 평균조회수(최근7건)", "릴스 평균좋아요수(최근7건)", "릴스 평균댓글수(최근7건)",
        "업데이트 시각",
    ]
    write_row(service, 1, headers)

# ─────────────────────────────────────────
# YouTube
# ─────────────────────────────────────────
def extract_channel_id(youtube, url):
    handle_match = re.search(r"/@([^/?\s]+)", url)
    if handle_match:
        handle = handle_match.group(1)
        res = youtube.search().list(part="snippet", q=handle, type="channel", maxResults=1).execute()
        if res.get("items"):
            return res["items"][0]["snippet"]["channelId"]
    ch_match = re.search(r"/channel/([^/?\s]+)", url)
    if ch_match:
        return ch_match.group(1)
    return None

def is_short(duration_iso):
    try:
        return isodate.parse_duration(duration_iso).total_seconds() <= 60
    except:
        return False

def get_youtube_stats(url):
    youtube = build("youtube", "v3", developerKey=YOUTUBE_API_KEY)
    channel_id = extract_channel_id(youtube, url)
    if not channel_id:
        print(f"  ⚠ 채널 ID 추출 실패: {url}")
        return {}

    ch_res = youtube.channels().list(part="statistics", id=channel_id).execute()
    if not ch_res.get("items"):
        return {}
    subs = int(ch_res["items"][0]["statistics"].get("subscriberCount", 0))

    search_res = youtube.search().list(
        part="id", channelId=channel_id, type="video",
        maxResults=50, order="date"
    ).execute()

    video_ids = [item["id"]["videoId"] for item in search_res.get("items", [])]
    if not video_ids:
        return {"subs": subs, "vid_views": 0, "vid_comments": 0, "short_views": 0, "short_comments": 0}

    vid_res = youtube.videos().list(part="statistics,contentDetails", id=",".join(video_ids)).execute()

    normal_views, normal_cmts, short_views, short_cmts = [], [], [], []
    for v in vid_res.get("items", []):
        s = v["statistics"]
        vw  = int(s.get("viewCount", 0))
        cmt = int(s.get("commentCount", 0))
        if is_short(v["contentDetails"]["duration"]):
            if len(short_views) < 7:
                short_views.append(vw); short_cmts.append(cmt)
        else:
            if len(normal_views) < 7:
                normal_views.append(vw); normal_cmts.append(cmt)

    def avg(lst): return round(sum(lst) / len(lst)) if lst else 0

    return {
        "subs": subs,
        "vid_views": avg(normal_views), "vid_comments": avg(normal_cmts),
        "short_views": avg(short_views), "short_comments": avg(short_cmts),
    }

# ─────────────────────────────────────────
# Instagram
# ─────────────────────────────────────────
def get_instagram_stats(username):
    L = instaloader.Instaloader()
    try:
        profile = instaloader.Profile.from_username(L.context, username)
    except Exception as e:
        print(f"  ⚠ 인스타 프로필 로드 실패 ({username}): {e}")
        return {}

    followers = profile.followers
    seven_days_ago = datetime.now(timezone.utc) - timedelta(days=7)
    reel_views, reel_likes, reel_cmts = [], [], []

    try:
        for post in profile.get_posts():
            if len(reel_views) >= 7:
                break
            if post.is_video and post.typename == "GraphVideo":
                reel_views.append(post.video_view_count or 0)
                reel_likes.append(post.likes)
                reel_cmts.append(post.comments)
            time.sleep(1)
    except Exception as e:
        print(f"  ⚠ 인스타 포스트 수집 오류 ({username}): {e}")

    def avg(lst): return round(sum(lst) / len(lst)) if lst else 0

    return {
        "followers": followers,
        "reel_views": avg(reel_views), "reel_likes": avg(reel_likes), "reel_comments": avg(reel_cmts),
    }

# ─────────────────────────────────────────
# 메인
# ─────────────────────────────────────────
def main():
    print("=== 인플루언서 데이터 수집 시작 ===")
    service = get_sheets_service()
    ensure_headers(service)
    rows = read_sheet(service)
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

    for i, row in enumerate(rows[DATA_START_ROW - 1:], start=DATA_START_ROW):
        row = list(row) + [""] * (COL_UPDATED + 1 - len(row))

        name     = str(row[COL_NAME]).strip()
        ctype    = str(row[COL_TYPE]).strip()
        insta_id = str(row[COL_INSTA_ID]).strip()
        yt_url   = str(row[COL_YT_URL]).strip()

        if not name:
            continue

        print(f"\n[{i}] {name} ({ctype})")

        if ctype == "유튜브 브랜디드" and yt_url:
            print("  → 유튜브 수집 중...")
            yt = get_youtube_stats(yt_url)
            if yt:
                row[COL_YT_SUBS] = yt.get("subs", "")
                row[COL_YT_VW]   = yt.get("vid_views", "")
                row[COL_YT_CMT]  = yt.get("vid_comments", "")
                row[COL_YT_SVW]  = yt.get("short_views", "")
                row[COL_YT_SCMT] = yt.get("short_comments", "")
                print(f"  ✓ 구독자 {yt.get('subs'):,}")

        if ctype == "인스타그램 릴스" and insta_id:
            print("  → 인스타 수집 중...")
            ig = get_instagram_stats(insta_id)
            if ig:
                row[COL_IG_FOL]  = ig.get("followers", "")
                row[COL_IG_VW]   = ig.get("reel_views", "")
                row[COL_IG_LIKE] = ig.get("reel_likes", "")
                row[COL_IG_CMT]  = ig.get("reel_comments", "")
                print(f"  ✓ 팔로워 {ig.get('followers'):,}")

        row[COL_UPDATED] = now_str
        write_row(service, i, row[:COL_UPDATED + 1])
        time.sleep(1)

    print("\n=== 완료 ===")

if __name__ == "__main__":
    main()
