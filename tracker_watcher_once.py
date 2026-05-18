#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tracker_watcher_once.py
스케줄러/bat에서 호출 — 1회 실행 후 종료

구조:
  tracker_data.json       ← launchMap + stockMap + 날짜목록 (가벼움)
  snapshots/YYYY-MM-DD.json ← 날짜별 스냅샷 (날짜당 1파일)

장점: 새로고침시 오늘 파일만 로드 → 날짜 쌓여도 빠름
"""
import json, logging, re, sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    print('[오류] pip install pandas openpyxl'); sys.exit(1)

SCRIPT_DIR = Path(__file__).resolve().parent
JSON_PATH  = SCRIPT_DIR / 'tracker_data.json'
SNAP_DIR   = SCRIPT_DIR / 'snapshots'

PAT_LIKES   = re.compile(r'좋아요_리뷰수_\d{8}', re.I)
PAT_LINKAGE = re.compile(r'linkage_modify', re.I)
PAT_STOCK   = re.compile(r'창고별_가용재고', re.I)

logging.basicConfig(level=logging.INFO, format='%(asctime)s  %(message)s', datefmt='%H:%M:%S')
log = logging.getLogger(__name__)

# ── 유틸 ──────────────────────────────────────────
def parse_int(v):
    if v is None: return None
    s = re.sub(r'[^\d]', '', str(v))
    return int(s) if s else None

def extract_date(fn, ct=None):
    m = re.search(r'(\d{4})(\d{2})(\d{2})', fn)
    if m: return f'{m.group(1)}-{m.group(2)}-{m.group(3)}'
    if ct:
        m2 = re.search(r'(\d{4})-(\d{2})-(\d{2})', str(ct))
        if m2: return f'{m2.group(1)}-{m2.group(2)}-{m2.group(3)}'
    return datetime.now().strftime('%Y-%m-%d')

# ── JSON (메타데이터만) ───────────────────────────
def load_meta():
    base = {'launchMap': {}, 'stockMap': {}, 'snapshotDates': [], 'meta': {}}
    if JSON_PATH.exists():
        try:
            with open(JSON_PATH, 'r', encoding='utf-8') as f:
                d = json.load(f)
            # 구버전 호환 (snapshots 키 있으면 무시)
            for k in base:
                if k not in d:
                    d[k] = base[k]
            return d
        except Exception as e:
            log.warning(f'메타 로드 실패: {e}')
    return base

def save_meta(data):
    # snapshots 폴더 기준으로 날짜목록 자동 갱신
    SNAP_DIR.mkdir(exist_ok=True)
    data['snapshotDates'] = sorted(f.stem for f in SNAP_DIR.glob('*.json'))
    data.setdefault('meta', {})['lastUpdated'] = datetime.now().isoformat()
    # snapshots 키는 저장 안 함
    save_data = {k: v for k, v in data.items() if k != 'snapshots'}
    with open(JSON_PATH, 'w', encoding='utf-8') as f:
        json.dump(save_data, f, ensure_ascii=False, separators=(',', ':'))

# ── 스냅샷 (날짜별 파일) ─────────────────────────
def snap_exists(date_key):
    return (SNAP_DIR / f'{date_key}.json').exists()

def save_snap(date_key, rows):
    SNAP_DIR.mkdir(exist_ok=True)
    with open(SNAP_DIR / f'{date_key}.json', 'w', encoding='utf-8') as f:
        json.dump(rows, f, ensure_ascii=False, separators=(',', ':'))

# ── 파서 ─────────────────────────────────────────
def parse_likes(path):
    df = pd.read_excel(path, dtype=str).fillna('')
    ct = df['수집일시'].iloc[0] if '수집일시' in df.columns and len(df) else None
    dk = extract_date(path.name, ct)
    seen = set(); rows = []
    for _, r in df.iterrows():
        cd = str(r.get('상품코드', '')).strip()
        ml = str(r.get('쇼핑몰', '')).strip()
        if not cd: continue
        k = f'{cd}|{ml}'
        if k in seen: continue
        seen.add(k)
        rows.append({
            'ml': ml, 'br': str(r.get('브랜드', '')).strip(),
            'cd': cd, 'nm': str(r.get('상품명', '')).strip(),
            'pr': parse_int(r.get('가격', '')),
            'dc': parse_int(r.get('할인율', '')),
            'lk': parse_int(r.get('좋아요수', '')),
            'rv': parse_int(r.get('리뷰수', '')),
            'ur': str(r.get('URL', '')).strip()
        })
    return dk, rows

def parse_linkage(path):
    df = pd.read_excel(path)
    result = {}
    for _, r in df.iterrows():
        code = str(r.get('쇼핑몰상품코드', '') or '').strip()
        if not code or code in result: continue
        model = str(r.get('모델명', '') or '').strip()
        price = parse_int(r.get('판매가'))
        dr = str(r.get('발매일', '') or '').strip()
        ld = None
        if dr:
            if re.match(r'^\d{4}-\d{2}-\d{2}$', dr): ld = dr
            elif re.match(r'^\d{8}$', dr): ld = f'{dr[:4]}-{dr[4:6]}-{dr[6:8]}'
        result[code] = {'d': ld, 'p': price, 'm': model}
    return result

def parse_stock(path):
    all_r = pd.read_excel(path, header=None).values.tolist()
    hdr = next((i for i, row in enumerate(all_r[:5]) if '품번' in [str(x) for x in row]), None)
    if hdr is None: raise ValueError('"품번" 컬럼 없음')
    headers = [str(x) for x in all_r[hdr]]
    pi, qi = headers.index('품번'), headers.index('가용재고')
    acc = {}
    for row in all_r[hdr+1:]:
        pbn = str(row[pi] if pi < len(row) else '').strip()
        qty = row[qi] if qi < len(row) else 0
        if not pbn or pbn in ('None', 'nan'): continue
        try: acc[pbn] = acc.get(pbn, 0) + float(qty or 0)
        except: pass
    return {k: int(v) for k, v in acc.items()}

# ── 메인 ─────────────────────────────────────────
def main():
    log.info('1회 실행 시작')
    data = load_meta()
    changed = False
    today_str = datetime.now().strftime('%Y%m%d')

    all_files = sorted(list(SCRIPT_DIR.glob('*.xlsx')) + list(SCRIPT_DIR.glob('*.xls')))

    # ① 좋아요: 날짜별 그룹핑 → 없는 날짜 + 오늘만 처리
    date_files = defaultdict(list)
    for f in all_files:
        if PAT_LIKES.search(f.name):
            m = re.search(r'(\d{8})', f.name)
            if m: date_files[m.group(1)].append(f)

    for ds, files in sorted(date_files.items()):
        dk = f'{ds[:4]}-{ds[4:6]}-{ds[6:8]}'
        is_today = (ds == today_str)
        # 이미 있고 오늘 아니면 스킵
        if snap_exists(dk) and not is_today:
            continue
        fp = sorted(files, reverse=True)[0]  # 같은 날짜 중 최신 파일
        try:
            dk2, rows = parse_likes(fp)
            save_snap(dk2, rows)
            log.info(f'좋아요 {dk2}: {len(rows)}개 저장 ({fp.name})')
            changed = True
        except Exception as e:
            log.error(f'{fp.name}: {e}')

    # ② linkage: 최신 파일 1개
    lfiles = sorted([f for f in all_files if PAT_LINKAGE.search(f.name)], reverse=True)
    if lfiles:
        try:
            data['launchMap'] = parse_linkage(lfiles[0])
            log.info(f'linkage {lfiles[0].name}: {len(data["launchMap"])}개')
            changed = True
        except Exception as e:
            log.error(f'{lfiles[0].name}: {e}')

    # ③ 재고: 최신 파일 1개
    sfiles = sorted([f for f in all_files if PAT_STOCK.search(f.name)], reverse=True)
    if sfiles:
        try:
            data['stockMap'] = parse_stock(sfiles[0])
            log.info(f'재고 {sfiles[0].name}: {len(data["stockMap"])}개 품번')
            changed = True
        except Exception as e:
            log.error(f'{sfiles[0].name}: {e}')

    if changed:
        save_meta(data)
        snap_count = len(list(SNAP_DIR.glob('*.json'))) if SNAP_DIR.exists() else 0
        log.info(f'✓ 업데이트 완료 (스냅샷 {snap_count}일치)')
    else:
        log.info('변경 없음')

if __name__ == '__main__': main()
