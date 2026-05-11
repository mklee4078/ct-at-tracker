#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tracker_watcher_once.py
스케줄러/bat에서 호출 — 1회 실행 후 종료
(상시 감시가 아닌 스케줄러 트리거 방식)
"""
import hashlib, json, logging, re, sys
from datetime import datetime
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    print('[오류] pip install pandas openpyxl'); sys.exit(1)

SCRIPT_DIR  = Path(__file__).resolve().parent
JSON_PATH   = SCRIPT_DIR / 'tracker_data.json'
STATE_PATH  = SCRIPT_DIR / '.watcher_state.json'  # 처리된 파일 해시 기록

PAT_LIKES   = re.compile(r'좋아요_리뷰수_\d{8}', re.I)
PAT_LINKAGE = re.compile(r'linkage_modify', re.I)
PAT_STOCK   = re.compile(r'창고별_가용재고', re.I)

logging.basicConfig(level=logging.INFO, format='%(asctime)s  %(message)s', datefmt='%H:%M:%S')
log = logging.getLogger(__name__)

def fhash(p):
    h = hashlib.md5()
    with open(p,'rb') as f:
        for c in iter(lambda: f.read(65536), b''): h.update(c)
    return h.hexdigest()

def parse_int(v):
    if v is None: return None
    s = re.sub(r'[^\d]','',str(v))
    return int(s) if s else None

def extract_date(fn, ct=None):
    m = re.search(r'(\d{4})(\d{2})(\d{2})', fn)
    if m: return f'{m.group(1)}-{m.group(2)}-{m.group(3)}'
    if ct:
        m2 = re.search(r'(\d{4})-(\d{2})-(\d{2})', str(ct))
        if m2: return f'{m2.group(1)}-{m2.group(2)}-{m2.group(3)}'
    return datetime.now().strftime('%Y-%m-%d')

def load_json():
    if JSON_PATH.exists():
        try:
            with open(JSON_PATH,'r',encoding='utf-8') as f: return json.load(f)
        except: pass
    return {'snapshots':{},'launchMap':{},'stockMap':{},'meta':{}}

def save_json(data):
    data['meta']['lastUpdated'] = datetime.now().isoformat()
    with open(JSON_PATH,'w',encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, separators=(',',':'))

def load_state():
    if STATE_PATH.exists():
        try:
            with open(STATE_PATH,'r') as f: return json.load(f)
        except: pass
    return {}

def save_state(state):
    with open(STATE_PATH,'w') as f: json.dump(state, f)

def parse_likes(path):
    df = pd.read_excel(path, dtype=str).fillna('')
    dk = extract_date(path.name, df['수집일시'].iloc[0] if '수집일시' in df.columns and len(df) else None)
    seen = set(); rows = []
    for _, r in df.iterrows():
        cd = str(r.get('상품코드','')).strip()
        ml = str(r.get('쇼핑몰','')).strip()
        if not cd: continue
        k = f'{cd}|{ml}'
        if k in seen: continue
        seen.add(k)
        rows.append({'ml':ml,'br':str(r.get('브랜드','')).strip(),'cd':cd,
            'nm':str(r.get('상품명','')).strip(),'pr':parse_int(r.get('가격','')),
            'dc':parse_int(r.get('할인율','')),'lk':parse_int(r.get('좋아요수','')),
            'rv':parse_int(r.get('리뷰수','')),'ur':str(r.get('URL','')).strip()})
    return dk, rows

def parse_linkage(path, existing):
    df = pd.read_excel(path)
    result = dict(existing); added = 0
    for _, r in df.iterrows():
        code = str(r.get('쇼핑몰상품코드','') or '').strip()
        if not code or code in result: continue
        model = str(r.get('모델명','') or '').strip()
        price = parse_int(r.get('판매가'))
        dr = str(r.get('발매일','') or '').strip(); ld = None
        if dr:
            if re.match(r'^\d{4}-\d{2}-\d{2}$', dr): ld = dr
            elif re.match(r'^\d{8}$', dr): ld = f'{dr[:4]}-{dr[4:6]}-{dr[6:8]}'
        result[code] = {'d':ld,'p':price,'m':model}; added += 1
    return result, added

def parse_stock(path):
    all_r = pd.read_excel(path, header=None).values.tolist()
    hdr = next((i for i,row in enumerate(all_r[:5]) if '품번' in [str(x) for x in row]), None)
    if hdr is None: raise ValueError('"품번" 없음')
    headers = [str(x) for x in all_r[hdr]]
    pi, qi = headers.index('품번'), headers.index('가용재고')
    acc = {}
    for row in all_r[hdr+1:]:
        pbn = str(row[pi] if pi < len(row) else '').strip()
        qty = row[qi] if qi < len(row) else 0
        if not pbn or pbn in ('None','nan'): continue
        try: acc[pbn] = acc.get(pbn,0) + float(qty or 0)
        except: pass
    return {k: int(v) for k,v in acc.items()}

def main():
    log.info('1회 실행 시작')
    data = load_json()
    state = load_state()
    changed = False

    files = (list(SCRIPT_DIR.glob('*.xlsx')) + list(SCRIPT_DIR.glob('*.xls')))
    # 재고 파일은 최신 1개만 처리 (파일명 날짜 기준)
    stock_files = sorted([f for f in files if PAT_STOCK.search(f.name)], reverse=True)

    for fp in files:
        try: h = fhash(fp)
        except: continue
        if state.get(str(fp)) == h: continue  # 이미 처리한 파일 스킵

        fn = fp.name
        try:
            if PAT_LIKES.search(fn):
                dk, rows = parse_likes(fp)
                if dk in data['snapshots']:
                    log.info(f'좋아요 {dk} 이미 존재 — 스킵')
                else:
                    data['snapshots'][dk] = rows
                    log.info(f'좋아요 {dk}: {len(rows)}개 추가')
                    changed = True
                state[str(fp)] = h

            elif PAT_LINKAGE.search(fn):
                result, added = parse_linkage(fp, data['launchMap'])
                data['launchMap'] = result
                log.info(f'linkage {fn}: {added}개 추가')
                changed = True
                state[str(fp)] = h

            elif PAT_STOCK.search(fn) and fp == stock_files[0]:
                # 재고는 가장 최신 파일만
                sm = parse_stock(fp)
                data['stockMap'] = sm
                log.info(f'재고 {fn}: {len(sm)}개 품번 완전 교체')
                changed = True
                state[str(fp)] = h

        except Exception as e:
            log.error(f'{fn}: {e}')

    if changed:
        save_json(data)
        save_state(state)
        log.info('✓ tracker_data.json 업데이트 완료')
    else:
        log.info('변경 없음')

if __name__ == '__main__': main()
