#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CT/AT 트래커 — 폴더 감시 & tracker_data.json 자동 업데이트
============================================================
규칙:
  - 좋아요_리뷰수_*.xlsx : 날짜별 누적, 중복(코드+채널) 제거
  - linkage_modify_*.xlsx : 쇼핑몰상품코드 기준 중복 스킵
  - 창고별_가용재고_*.xlsx : 항상 최신 파일로 완전 교체

설치: pip install pandas openpyxl
실행: python tracker_watcher.py
"""
import hashlib, json, logging, re, sys, time
from datetime import datetime
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    print('[오류] pip install pandas openpyxl'); sys.exit(1)

SCRIPT_DIR  = Path(__file__).resolve().parent
JSON_PATH   = SCRIPT_DIR / 'tracker_data.json'
INTERVAL    = 30

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
            with open(JSON_PATH, 'r', encoding='utf-8') as f: return json.load(f)
        except Exception as e: log.warning(f'JSON 로드 실패: {e}')
    return {'snapshots': {}, 'launchMap': {}, 'stockMap': {}, 'meta': {}}

def save_json(data):
    data['meta']['lastUpdated'] = datetime.now().isoformat()
    with open(JSON_PATH, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, separators=(',',':'))

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
    if hdr is None: raise ValueError('"품번" 컬럼 없음')
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

class Watcher:
    def __init__(self):
        self.hashes = {}
        self.data = load_json()
        log.info('='*52)
        log.info('  CT/AT 트래커 감시 시작')
        log.info(f'  폴더: {SCRIPT_DIR}')
        log.info(f'  JSON: {JSON_PATH}')
        log.info(f'  주기: {INTERVAL}초 | Ctrl+C 종료')
        log.info('='*52)
        self._initial()

    def _initial(self):
        files = sorted(SCRIPT_DIR.glob('*.xlsx')) + sorted(SCRIPT_DIR.glob('*.xls'))
        changed = False
        for fp in files:
            self.hashes[str(fp)] = fhash(fp)
            if PAT_LIKES.search(fp.name):
                try:
                    dk, rows = parse_likes(fp)
                    if dk not in self.data['snapshots']:
                        self.data['snapshots'][dk] = rows
                        log.info(f'[초기] {dk}: {len(rows)}개')
                        changed = True
                except Exception as e: log.error(f'[초기] {fp.name}: {e}')
        if changed: save_json(self.data)
        log.info(f'초기 완료 — 스냅샷 {len(self.data["snapshots"])}일치')

    def scan(self):
        files = list(SCRIPT_DIR.glob('*.xlsx')) + list(SCRIPT_DIR.glob('*.xls'))
        changed = False
        for fp in files:
            try: h = fhash(fp)
            except: continue
            if self.hashes.get(str(fp)) == h: continue
            self.hashes[str(fp)] = h
            fn = fp.name; log.info(f'변경: {fn}')
            try:
                if PAT_LIKES.search(fn):
                    dk, rows = parse_likes(fp)
                    if dk in self.data['snapshots']:
                        log.info(f'  {dk} 이미 존재 — 스킵')
                    else:
                        self.data['snapshots'][dk] = rows
                        log.info(f'  좋아요 {dk}: {len(rows)}개 추가'); changed = True
                elif PAT_LINKAGE.search(fn):
                    result, added = parse_linkage(fp, self.data['launchMap'])
                    self.data['launchMap'] = result
                    log.info(f'  linkage: {added}개 추가'); changed = True
                elif PAT_STOCK.search(fn):
                    sm = parse_stock(fp)
                    self.data['stockMap'] = sm
                    log.info(f'  재고: {len(sm)}개 품번 완전 교체'); changed = True
            except Exception as e: log.error(f'  오류: {e}')
        if changed:
            save_json(self.data)
            log.info('✓ tracker_data.json 업데이트')

    def run(self):
        while True:
            try: self.scan(); time.sleep(INTERVAL)
            except KeyboardInterrupt: log.info('종료'); break
            except Exception as e: log.error(f'오류: {e}'); time.sleep(INTERVAL)

if __name__ == '__main__': Watcher().run()
