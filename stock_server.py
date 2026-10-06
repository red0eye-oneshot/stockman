#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
주식 수익률 트래커 - 로컬 서버 v1.1
────────────────────────────────────
실행: python stock_server.py
  또는 run_server.bat 더블클릭

브라우저가 자동으로 http://localhost:5555 로 열립니다.
같은 Wi-Fi 스마트폰: http://[이 PC의 IP]:5555

필요 패키지: requests  (pip install requests)
선택 패키지: pykrx     (pip install pykrx)  ← 더 정확한 데이터
"""

import os, sys, re, json, time, calendar, threading, socket, webbrowser, io, html
from concurrent.futures import ThreadPoolExecutor, as_completed, TimeoutError as FutureTimeoutError
from datetime import datetime, timezone, timedelta
from urllib.parse import unquote, quote
from http.server import HTTPServer, BaseHTTPRequestHandler
from socketserver import ThreadingMixIn

# pythonw.exe: stdout/stderr 가 None 이면 devnull 로 대체
if sys.stdout is None:
    sys.stdout = open(os.devnull, 'w')
if sys.stderr is None:
    sys.stderr = open(os.devnull, 'w')

# 로그 파일 (서버와 같은 폴더에 server_log.txt 생성)
_LOG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'server_log.txt')

def _log(msg: str):
    ts = datetime.now().strftime('%H:%M:%S')
    line = f"[{ts}] {msg}"
    print(line)
    try:
        with open(_LOG_PATH, 'a', encoding='utf-8') as f:
            f.write(line + '\n')
    except Exception:
        pass
try:
    import requests
except ImportError:
    print("\n[오류] requests 패키지가 없습니다.")
    print("       설치 후 다시 실행하세요: pip install requests\n")
    input("엔터를 누르면 종료...")
    sys.exit(1)

_saved_stdout = sys.stdout
_saved_stderr = sys.stderr
try:
    import io as _io
    sys.stdout = sys.stderr = _io.StringIO()
    from pykrx import stock as krx
    sys.stdout = _saved_stdout
    sys.stderr = _saved_stderr
    PYKRX = True
except Exception:
    sys.stdout = _saved_stdout
    sys.stderr = _saved_stderr
    PYKRX = False

try:
    import openpyxl
    OPENPYXL = True
except Exception:
    OPENPYXL = False

try:
    from reportlab.pdfgen import canvas as _rl_canvas
    from reportlab.lib.pagesizes import A4 as _RL_A4
    from reportlab.pdfbase import pdfmetrics as _rl_pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont as _RL_UnicodeCIDFont
    _rl_pdfmetrics.registerFont(_RL_UnicodeCIDFont('HYSMyeongJo-Medium'))
    REPORTLAB = True
except Exception:
    REPORTLAB = False

try:
    import docx as _docx
    DOCX_OK = True
except Exception:
    DOCX_OK = False

import smtplib
from email.message import EmailMessage

# ── 설정 ────────────────────────────────────────────────
PORT          = int(os.environ.get('PORT', 5555))   # Render는 환경변수로 PORT 지정
CACHE_TTL     = 60    # 장중 캐시 유효시간(초) — 빠른 실시간 갱신
CACHE_TTL_OFF = 1800  # 장외 캐시 유효시간(초) — 30분 (장외엔 가격 변동 없음)
TIMEOUT       = 5     # HTTP 요청 타임아웃(초) — 느린 서버 빠르게 포기
API_KEY   = os.environ.get('API_KEY', '')        # 비어있으면 인증 비활성 (로컬 개발용)

# ── 텔레그램 봇 (일일 리포트 발송) ──────────────────────────
# 보안: 토큰/챗ID를 소스에 직접 적지 않음 → 환경변수(Render 배포용) 우선,
# 없으면 로컬 telegram_config.json(.gitignore 처리, 깃허브에 올라가지 않음)에서 읽음.
TELEGRAM_TOKEN   = os.environ.get('TELEGRAM_BOT_TOKEN', '')
TELEGRAM_CHAT_ID = os.environ.get('TELEGRAM_CHAT_ID', '')
if not TELEGRAM_TOKEN or not TELEGRAM_CHAT_ID:
    try:
        with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'telegram_config.json'), 'r', encoding='utf-8') as _tf:
            _tg_cfg = json.load(_tf)
        TELEGRAM_TOKEN   = TELEGRAM_TOKEN   or _tg_cfg.get('bot_token', '')
        TELEGRAM_CHAT_ID = TELEGRAM_CHAT_ID or _tg_cfg.get('chat_id', '')
    except Exception:
        pass

# ── 이메일 (일일 분석 엑셀/PDF 자동 발송) ──────────────────
# 보안: 앱 비밀번호를 소스에 직접 적지 않음 → 환경변수(Render 배포용) 우선,
# 없으면 로컬 email_config.json(.gitignore 처리, 깃허브에 올라가지 않음)에서 읽음.
EMAIL_USER     = os.environ.get('EMAIL_USER', '')          # 보내는 계정 (Gmail 주소)
EMAIL_APP_PASS = os.environ.get('EMAIL_APP_PASSWORD', '')  # Gmail 앱 비밀번호(16자리)
EMAIL_TO       = os.environ.get('EMAIL_TO', '')            # 받는 주소
if not EMAIL_USER or not EMAIL_APP_PASS:
    try:
        with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'email_config.json'), 'r', encoding='utf-8') as _ef:
            _em_cfg = json.load(_ef)
        EMAIL_USER     = EMAIL_USER     or _em_cfg.get('smtp_user', '')
        EMAIL_APP_PASS = EMAIL_APP_PASS or _em_cfg.get('smtp_app_password', '')
        EMAIL_TO       = EMAIL_TO       or _em_cfg.get('to_email', '')
    except Exception:
        pass

# ── Gemini (무료 AI로 종목 뉴스 해석/요약) ──────────────────
# 보안: API 키를 소스에 직접 적지 않음 → 환경변수(Render 배포용) 우선,
# 없으면 로컬 gemini_config.json(.gitignore 처리, 깃허브에 올라가지 않음)에서 읽음.
GEMINI_API_KEY = os.environ.get('GEMINI_API_KEY', '')
GEMINI_MODEL   = os.environ.get('GEMINI_MODEL', 'gemini-2.5-flash')
if not GEMINI_API_KEY:
    try:
        with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'gemini_config.json'), 'r', encoding='utf-8') as _gf:
            _ge_cfg = json.load(_gf)
        GEMINI_API_KEY = GEMINI_API_KEY or _ge_cfg.get('api_key', '')
        GEMINI_MODEL   = _ge_cfg.get('model', GEMINI_MODEL)
    except Exception:
        pass

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Referer": "https://finance.naver.com",
    "Accept":  "*/*",
}

# ── HTTP 세션 (연결 재사용 - 매 요청마다 새 TLS 핸드셰이크 방지) ────
SESSION = requests.Session()
_adapter = requests.adapters.HTTPAdapter(pool_connections=50, pool_maxsize=50, max_retries=0)
SESSION.mount('https://', _adapter)
SESSION.mount('http://', _adapter)

# ── 캐시 ────────────────────────────────────────────────
_cache      = {}          # code -> {'data': dict, 'ts': float}
_cache_lock = threading.Lock()

# ── 공유 스레드풀 (요청마다 새로 만들지 않음) ─────────────────
# 주의: `with ThreadPoolExecutor(...) as ex:` 패턴은 블록을 빠져나갈 때
# ex.shutdown(wait=True)를 호출해서, 개별 future.result(timeout=N)이
# 타임아웃으로 "포기"한 뒤에도 실제로는 그 느린 작업이 끝날 때까지
# 함수 전체가 계속 블록되는 문제가 있었음(pykrx가 사내망에서 응답 없이
# 멈추는 경우 특히 치명적). 공유 풀을 쓰면 느린 작업은 백그라운드에서
# 계속 돌게 두고, 요청 처리는 제때 반환된다.
_shared_pool = ThreadPoolExecutor(max_workers=40)

# ── 브라우저 탭 닫으면 자동 종료 (로컬 PC 모드 전용) ──────────────
# 클라이언트가 탭을 열어두는 동안 주기적으로 /api/heartbeat 를 호출해서
# _last_heartbeat 를 갱신함. 일정 시간(HEARTBEAT_GRACE) 동안 신호가 없으면
# "탭이 닫혔다"고 보고 파이썬 프로세스를 스스로 종료함.
# 새로고침(F5) 정도의 짧은 공백은 grace 시간 안에 새 heartbeat가 다시 들어오므로
# 서버가 꺼지지 않음 — 진짜로 탭을 닫고 한동안 안 열었을 때만 종료됨.
_last_heartbeat = time.time()
# 90초는 너무 짧았음 — Chrome 등이 "다른 탭 보는 중"이거나 창을 최소화하면
# 백그라운드 탭의 setInterval을 1분 이상 단위로 강하게 스로틀링하는 경우가 있어,
# 탭을 안 닫았는데도 서버가 꺼지는 오탐(false positive)이 실제로 발생함.
# 진짜 목적(탭을 닫고 한참 방치했을 때만 종료)은 유지하면서 오탐을 없애기 위해
# 여유를 10분으로 크게 늘림.
HEARTBEAT_GRACE = 600  # 이 시간(초) 동안 heartbeat가 없으면 종료

def _heartbeat_watchdog():
    while True:
        time.sleep(10)
        idle = time.time() - _last_heartbeat
        if idle > HEARTBEAT_GRACE:
            # pyw(콘솔 없음)로 실행 중이면 print()는 아무도 못 보므로 반드시 _log()로도 남김
            _log(f"브라우저 탭이 닫힌 것으로 감지됨 ({idle:.0f}초간 응답 없음) → 서버 종료")
            os._exit(0)

# ── KRX 실패 이력 ───────────────────────────────────────
_krx_failed = set()
_inv_debug: dict = {}   # 수급 디버그 로그
PYKRX_INV_DEAD = True   # pykrx 투자자 API hang 방지 → naver HTML 직접 사용
PYKRX_CANDLES_DEAD = True   # pykrx 일봉(캔들) API도 같은 사유로 hang → 네이버 fchart 직접 사용 (속도 개선)

# ── 컨센서스 캐시 (1시간) ─────────────────────────────────
_cs_cache: dict = {}    # code -> {'data': dict, 'ts': float}
_cs_lock  = threading.Lock()
CS_TTL    = 3600        # 1시간

# ── 재무요약 캐시 (1시간) ─────────────────────────────────
_fin_cache: dict = {}   # code -> {'data': dict, 'ts': float}
_fin_lock  = threading.Lock()
FIN_TTL    = 3600       # 1시간

# ── 수급(투자자) 캐시 (15분) ────────────────────────────────
# fetch_investor_data 는 네이버를 2~3회 순차 스크래핑해 가장 느림.
# 수급 데이터는 분 단위로 바뀌지 않으므로 별도 TTL로 캐시해 반복 스크래핑을 줄인다.
_invd_cache: dict = {}   # code -> {'data': dict, 'ts': float}
_invd_lock  = threading.Lock()
INVD_TTL    = 900        # 15분


# ════════════════════════════════════════════════════════
# 데이터 조회
# ════════════════════════════════════════════════════════

def is_etf(code: str) -> bool:
    return bool(re.search(r'[A-Za-z]', code))


def fetch_candles_pykrx(code: str) -> list:
    """pykrx로 KRX 공식 일봉 조회 (3개월 OHLC)"""
    if code in _krx_failed:
        return []
    from_d = (datetime.now() - timedelta(days=95)).strftime('%Y%m%d')  # 3개월 여유
    to_d   = datetime.now().strftime("%Y%m%d")
    try:
        if is_etf(code):
            df = krx.get_etf_ohlcv_by_date(from_d, to_d, code)
        else:
            df = krx.get_market_ohlcv(from_d, to_d, code)
        if df is None or df.empty:
            _krx_failed.add(code)
            return []
        candles = []
        for idx, row in df.iterrows():
            d = idx.strftime("%Y-%m-%d")
            try:
                o = int(row.get('시가', row.get('Open', 0)))
                h = int(row.get('고가', row.get('High', 0)))
                l = int(row.get('저가', row.get('Low',  0)))
                c = int(row.get('종가', row.get('Close', 0)))
            except Exception:
                c = int(row.iloc[-2]) if len(row) > 1 else 0
                o = h = l = c
            if c > 0:
                candles.append({'date': d, 'open': o, 'high': h, 'low': l, 'close': c})
        return sorted(candles, key=lambda x: x['date'])
    except Exception:
        _krx_failed.add(code)
        return []


def fetch_candles_naver(code: str) -> list:
    """네이버 fchart XML로 일봉 조회 (fallback, 3개월 OHLC)"""
    url = (
        f"https://fchart.stock.naver.com/sise.nhn"
        f"?symbol={code}&timeframe=day&count=70&requestType=0"
    )
    r = SESSION.get(url, headers=HEADERS, timeout=TIMEOUT)
    r.raise_for_status()
    candles = []
    for m in re.finditer(
        r'data="(\d{8})\|([\d.]+)\|([\d.]+)\|([\d.]+)\|([\d.]+)\|', r.text
    ):
        d_str = m.group(1)
        o = round(float(m.group(2)))
        h = round(float(m.group(3)))
        l = round(float(m.group(4)))
        c = round(float(m.group(5)))
        if c > 0:
            candles.append({
                'date':  f"{d_str[:4]}-{d_str[4:6]}-{d_str[6:8]}",
                'open':  o,
                'high':  h,
                'low':   l,
                'close': c,
            })
    return sorted(candles, key=lambda x: x['date'])


def fetch_candles(code: str) -> list:
    """네이버 fchart(기본) → pykrx(PYKRX_CANDLES_DEAD=False일 때만 보조) 순서로 일봉 조회.
       pykrx는 사내망/프록시 환경에서 응답 없이 멈추는 사례가 있어 기본적으로 건너뜀
       (투자자 수급·재무 데이터와 동일한 이유로 이미 비활성화된 것과 같은 조치)."""
    if PYKRX and not PYKRX_CANDLES_DEAD:
        c = fetch_candles_pykrx(code)
        if c:
            return c
    try:
        return fetch_candles_naver(code)
    except Exception as e:
        _log(f"[WARN] fetch_candles_naver({code}) 실패: {e}")
        if PYKRX and PYKRX_CANDLES_DEAD:
            # naver도 실패하면 마지막 수단으로 pykrx 한 번 시도
            return fetch_candles_pykrx(code)
        return []


def fetch_price_naver(code: str):
    """
    현재가/전일종가/종목명 조회
    1) 네이버 모바일 API  2) finance.naver.com main.naver HTML fallback
    """
    px, prev, name = None, None, ''

    # ── 1) 네이버 모바일 API (원래 방식) ───────────────────
    try:
        url = f"https://m.stock.naver.com/api/stock/{code}/basic"
        r   = SESSION.get(url, headers=HEADERS, timeout=8)
        r.raise_for_status()
        d   = r.json()

        for key in ('stockName', 'itemName', 'name', 'stockEndName', 'title'):
            raw = d.get(key)
            if raw and isinstance(raw, str) and raw.strip():
                name = raw.strip(); break

        for key in ('currentPrice', 'stockEndPrice', 'closePrice', 'price'):
            raw = d.get(key)
            if raw is None: continue
            try:
                v = int(str(raw).replace(',', '').replace(' ', ''))
                if v > 0: px = v; break
            except (ValueError, TypeError): continue

        if px:
            for key in ('compareToPreviousClosePrice', 'changePrice', 'fluctuations'):
                raw = d.get(key)
                if raw is None: continue
                try:
                    chg  = int(str(raw).replace(',', '').replace('+', '').replace(' ', ''))
                    cand = px - chg
                    if cand > 0: prev = cand; break
                except (ValueError, TypeError): continue
            return px, prev, name
    except Exception:
        pass

    # ── 2) finance.naver.com main.naver HTML fallback ────────
    try:
        pc_h = {**HEADERS, "Accept": "text/html,*/*"}
        url2 = f"https://finance.naver.com/item/main.naver?code={code}"
        r2   = SESSION.get(url2, headers=pc_h, timeout=TIMEOUT)
        r2.raise_for_status()
        html = r2.text

        # 종목명
        if not name:
            nm_m = re.search(
                r'<title>\s*(.+?)\s*(?:\d{{4,6}})?\s*[-:]\s*네이버\s*금융',
                html, re.IGNORECASE
            )
            if nm_m:
                cand = nm_m.group(1).strip().rstrip('-').strip()
                if cand and not re.fullmatch(r'[\d\s\-:]+', cand):
                    name = cand

        # 현재가: <em id="_nowVal">62,700</em>
        m_px = re.search(r'id=["\']_nowVal["\'][^>]*>\s*([\d,]+)', html)
        if not m_px:
            m_px = re.search(r'_nowVal[^>]*>\s*<[^>]+>\s*([\d,]+)', html)
        if m_px:
            px = int(m_px.group(1).replace(',', ''))

        # 전일종가 추정: 변동금액 + 등락 방향
        if px:
            m_chg = re.search(r'id=["\']_change["\'][^>]*>([\s\S]{0,30}?)(\d[\d,]+)', html)
            if m_chg:
                try:
                    chg = int(m_chg.group(2).replace(',', ''))
                    ctx = m_chg.group(0)
                    if re.search(r'down|하락|nv02', ctx, re.I):
                        prev = px + chg   # 하락: 전일 = 현재 + 하락폭
                    else:
                        prev = px - chg   # 상승: 전일 = 현재 - 상승폭
                    if prev <= 0: prev = None
                except Exception:
                    prev = None
    except Exception:
        pass

    return px, prev, name


def past_price(candles: list, months: int):
    """N개월 전 말일 종가 (없으면 인접 거래일)"""
    now = datetime.now()
    mon = now.month - months
    yr  = now.year
    while mon <= 0:
        mon += 12
        yr  -= 1
    last = calendar.monthrange(yr, mon)[1]
    tgt  = f"{yr:04d}-{mon:02d}-{last:02d}"
    bef  = [c for c in candles if c['date'] <= tgt]
    aft  = [c for c in candles if c['date'] >  tgt]
    if bef: return bef[-1]['close']
    if aft: return aft[0]['close']
    return None


def is_market_open() -> bool:
    """KST 기준 장중 여부 (월~금 09:00~15:30)"""
    utc    = datetime.now(timezone.utc)
    kst_h  = (utc.hour + 9) % 24
    # UTC 15시 이후는 KST 익일 요일
    kst_wd = (utc.weekday() + (1 if utc.hour >= 15 else 0)) % 7
    total  = kst_h * 60 + utc.minute
    if kst_wd >= 5:   # 토·일
        return False
    return 540 <= total < 930   # 09:00 ~ 15:30


def _parse_int(v) -> int:
    """문자열/숫자 → int 변환"""
    if v is None:
        return 0
    try:
        return int(str(v).replace(',', '').replace('+', '').strip() or 0)
    except Exception:
        return 0


def fetch_investor_data(code: str) -> dict:
    """수급 데이터 캐시 래퍼 (TTL 15분) → 실제 조회는 _fetch_investor_data_raw"""
    with _invd_lock:
        cached = _invd_cache.get(code)
        if cached and (time.time() - cached['ts']) < INVD_TTL:
            return cached['data']
    result = _fetch_investor_data_raw(code)
    with _invd_lock:
        _invd_cache[code] = {'data': result, 'ts': time.time()}
    return result


def _fetch_investor_data_raw(code: str) -> dict:
    """
    투자자별 순매수(60일 누계) + 외국인 보유비중 조회
    시도 순서:
      A) pykrx  → KRX 공식 데이터 (가장 신뢰도 높음)
      B) 네이버 모바일 API (여러 엔드포인트 시도)
      C) 네이버 PC HTML 파싱
    """
    frgn_ratio = None
    dividend   = None
    div_yield  = None
    inv        = {}
    stock_name = ''
    _dbg       = []   # 디버그 로그 (서버 콘솔 + /api/debug 엔드포인트)

    # ════════════════════════════════
    # A) pykrx — KRX 공식 투자자 데이터
    #    한 번 실패하면 PYKRX_INV_DEAD=True로 표시해 같은 세션에서는 더 이상 시도 안 함
    # ════════════════════════════════
    global PYKRX_INV_DEAD
    if PYKRX and not PYKRX_INV_DEAD:
        from datetime import timedelta
        to_d   = datetime.now().strftime('%Y%m%d')
        from_d = (datetime.now() - timedelta(days=70)).strftime('%Y%m%d')
        etf = is_etf(code)

        # A-1) 종목별 투자자 거래실적
        try:
            if etf:
                df = krx.get_market_trading_value_by_investor(from_d, to_d, code, etf=True)
            else:
                df = krx.get_market_trading_value_by_investor(from_d, to_d, code)

            if df is not None and not df.empty:
                _dbg.append(f'pykrx A1 cols={list(df.columns)} idx={list(df.index)}')
                col = next((c for c in ['순매수','순매수금액','net'] if c in df.columns), None)
                if col:
                    for idx in df.index:
                        try:
                            inv[str(idx)] = int(df.loc[idx, col])
                        except Exception:
                            pass
                    _dbg.append(f'pykrx A1 OK: {len(inv)}건')
            else:
                _dbg.append('pykrx A1: empty df')
                PYKRX_INV_DEAD = True
        except Exception as e:
            _dbg.append(f'pykrx A1 ERR: {type(e).__name__}: {e}')
            PYKRX_INV_DEAD = True   # 다음 종목부터는 시도조차 안 함

    # ════════════════════════════════
    # B) 네이버 모바일 JSON API — /api/stock/{code}/integration
    #    (예전엔 /api/stock/{code}/investor 등을 시도했는데 2026년에 404로 전부 막혀서 제거했었음.
    #     integration은 가격 조회에 이미 쓰고 있는 /basic과 같은 계열 엔드포인트라 더 안정적일 것으로
    #     보고 새로 추가 — 외국인소진율(foreignRate)과 배당금(dividend)을 여기서 가져옴)
    # ════════════════════════════════
    try:
        r_int = SESSION.get(f"https://m.stock.naver.com/api/stock/{code}/integration",
                             headers=HEADERS, timeout=8)
        _dbg.append(f'mstock integration: HTTP {r_int.status_code}')
        if r_int.status_code == 200:
            j_int = r_int.json()
            rows = j_int.get('totalInfos') or []
            info = {}
            for row in rows:
                if isinstance(row, dict) and 'code' in row:
                    info[row['code']] = row.get('value')
            _dbg.append(f'  integration keys={list(info.keys())[:20]}')
            fr = _mkt_num(info.get('foreignRate'))
            if fr is not None:
                frgn_ratio = fr
                _dbg.append(f'  frgnRatio(integration)={frgn_ratio}')
            dv = _mkt_num(info.get('dividend'))
            if dv is not None:
                dividend = dv
                _dbg.append(f'  dividend(integration)={dividend}')
            dy = _mkt_num(info.get('dividendYieldRatio'))
            if dy is not None:
                div_yield = dy
    except Exception as e:
        _dbg.append(f'mstock integration ERR: {type(e).__name__}: {e}')

    # ════════════════════════════════
    # C) 네이버 PC HTML — frgnRatio + sise_investor 전체 수급
    # ════════════════════════════════
    pc_headers = {**HEADERS, "Accept": "text/html,*/*"}

    # C-1) frgn.naver → 외국인 보유비중 + 일별 기관/외국인 순매매
    # 테이블 헤더 문자열 의존 없이 날짜 행 직접 스캔
    inv_daily = []
    try:
        url_frgn = f"https://finance.naver.com/item/frgn.naver?code={code}"
        r2 = SESSION.get(url_frgn, headers=pc_headers, timeout=TIMEOUT)
        _dbg.append(f'naver frgn.naver: HTTP {r2.status_code}')
        if r2.status_code == 200:
            html = r2.text

            # 종목명 추출 (title 태그: "삼성전자 005930 : 네이버 금융" 등)
            if not stock_name:
                nm_m = re.search(
                    r'<title>\s*(.+?)\s*(?:\d{4,6})?\s*(?:[-:]|:)\s*네이버\s*금융',
                    html, re.IGNORECASE
                )
                if nm_m:
                    cand = nm_m.group(1).strip().rstrip('-').strip()
                    # 코드 번호·특수문자만 남은 경우 제외
                    if cand and not re.fullmatch(r'[\d\s\-:]+', cand):
                        stock_name = cand
                        _dbg.append(f'  name={stock_name!r}')

            # 외국인 소진율(보유율)
            if frgn_ratio is None:
                for pat in [
                    r'외국인소진율\(B/A\)[\s\S]{0,2000}?<td[^>]*>\s*<em[^>]*>\s*([\d.]+)\s*%',
                    r'외국인소진율[\s\S]{0,2000}?<em[^>]*>\s*([\d.]+)\s*%',
                    r'외국인\s*소진율[^>]*>\s*([\d.]+)\s*%',
                ]:
                    m = re.search(pat, html)
                    if m:
                        frgn_ratio = float(m.group(1))
                        _dbg.append(f'  frgnRatio={frgn_ratio}')
                        break
                if frgn_ratio is None:
                    _dbg.append('  frgnRatio: 패턴 실패')

            # 날짜가 포함된 모든 onMouseOver 행 스캔 (테이블 헤더 무관)
            all_rows = re.findall(r'<tr[^>]+onMouseOver[^>]*>[\s\S]*?</tr>', html)
            _dbg.append(f'  rows: {len(all_rows)}')
            for row in all_rows:
                date_m = re.search(r'(\d{4}\.\d{2}\.\d{2})', row)
                if not date_m:
                    continue
                sp = re.findall(
                    r'<span[^>]*class="tah[^"]*"[^>]*>\s*([-+]?[\d,]+(?:\.\d+)?%?)\s*</span>',
                    row
                )
                sp = [s.strip() for s in sp if s.strip()]
                n = len(sp)
                if n < 4:
                    continue
                try:
                    # 컬럼: 종가/전일비/등락률/거래량/기관/외국인/보유주수/보유율
                    if n >= 8:
                        oi, fi, hi, ri = 4, 5, 6, 7
                    elif n >= 6:
                        oi, fi, hi, ri = n-4, n-3, n-2, n-1
                    else:
                        oi, fi, hi, ri = 0, 1, 2, 3
                    entry = {
                        'date':  date_m.group(1),
                        'organ': _parse_int(sp[oi]),
                        'frgn':  _parse_int(sp[fi]),
                    }
                    if hi < n:
                        entry['frgnHold'] = _parse_int(sp[hi])
                    if ri < n:
                        try:
                            entry['frgnRate'] = float(
                                sp[ri].replace('%','').replace(',','').strip()
                            )
                        except Exception:
                            pass
                    inv_daily.append(entry)
                    if len(inv_daily) >= 7:
                        break
                except Exception:
                    continue
            _dbg.append(f'  invDaily: {len(inv_daily)}행, sp예시:{sp[:3] if all_rows else []}')
    except Exception as e:
        _dbg.append(f'frgn.naver ERR: {type(e).__name__}: {e}')

    # C-2) sise_investor.naver → 최근 7거래일 합계 순매수(주) + 개인 탭용 일별 수급
    # 컬럼 순서: 개인, 외국인, 기관합계, 금융투자, 보험, 투신, 사모, 은행, 기타금융, 연기금등, 국가, 기타법인
    indiv_daily = []   # [{'date':..., 'indiv':..., 'frgn':..., 'organ':...}, ...] 개인 탭 표시용
    if not inv:
        SI_COLS = ['개인','외국인','기관합계','금융투자','보험','투신','사모','은행','기타금융','연기금','국가','기타법인']
        try:
            url_si = f"https://finance.naver.com/item/sise_investor.naver?code={code}"
            r3 = SESSION.get(url_si, headers=pc_headers, timeout=TIMEOUT)
            _dbg.append(f'sise_investor: HTTP {r3.status_code}')
            if r3.status_code == 200:
                html_si = r3.text
                rows_si = re.findall(r'<tr[^>]*onMouseOver[^>]*>[\s\S]*?</tr>', html_si)
                _dbg.append(f'  rows: {len(rows_si)}')
                acc = {}   # 7일 누계
                days_ok = 0
                for row in rows_si[:10]:   # 최근 10거래일 (개인 탭은 조금 더 넉넉히)
                    date_m = re.search(r'(\d{4}\.\d{2}\.\d{2})', row)
                    spans = re.findall(
                        r'<span[^>]*class="tah[^"]*"[^>]*>\s*([-+0-9,\s]+)\s*</span>',
                        row
                    )
                    spans = [s.strip() for s in spans if s.strip()]
                    if len(spans) >= len(SI_COLS):
                        if days_ok < 7:
                            for i, col in enumerate(SI_COLS):
                                try:
                                    acc[col] = acc.get(col, 0) + _parse_int(spans[i])
                                except Exception:
                                    pass
                            days_ok += 1
                        if date_m:
                            indiv_daily.append({
                                'date':  date_m.group(1),
                                'indiv': _parse_int(spans[0]),
                                'frgn':  _parse_int(spans[1]),
                                'organ': _parse_int(spans[2]),
                            })
                if days_ok > 0:
                    inv.update(acc)
                    _dbg.append(f'  sise_investor OK: {days_ok}일 누계 {len(inv)}건, indivDaily {len(indiv_daily)}건')
                else:
                    _dbg.append('  sise_investor: 유효 행 없음')
        except Exception as e:
            _dbg.append(f'sise_investor ERR: {type(e).__name__}: {e}')

    # C-3) fallback: frgn.naver 60일 누계 (외국인/기관만)
    if not inv:
        try:
            url_frgn = f"https://finance.naver.com/item/frgn.naver?code={code}"
            r2b = SESSION.get(url_frgn, headers=pc_headers, timeout=TIMEOUT)
            if r2b.status_code == 200:
                tbl_m = re.search(r'외국인 기관 순매매 거래량[\s\S]*?</table>', r2b.text)
                if tbl_m:
                    rows_fb = re.findall(r'<tr\s+onMouseOver[\s\S]*?</tr>', tbl_m.group(0))
                    tot_o = tot_f = days = 0
                    for row in rows_fb[:60]:
                        sp = re.findall(r'<span[^>]*class="tah[^"]*"[^>]*>\s*([-+0-9,.\s]+)\s*</span>', row)
                        sp = [s.strip() for s in sp if s.strip()]
                        if len(sp) >= 7:
                            tot_o += _parse_int(sp[5]); tot_f += _parse_int(sp[6]); days += 1
                    if days:
                        inv['기관합계'] = tot_o; inv['외국인'] = tot_f
                        _dbg.append(f'  fallback 60일 누계 OK: {days}일')
        except Exception as e:
            _dbg.append(f'fallback ERR: {type(e).__name__}: {e}')

    # 디버그 로그 출력 (서버 콘솔)
    if not inv:
        print(f"  [수급 실패] {code}: " + " | ".join(_dbg))
    else:
        print(f"  [수급 OK]  {code}: 외국인={inv.get('외국인','?')} 기관={inv.get('기관합계','?')}")

    _inv_debug[code] = _dbg

    # ── 키 이름 정규화 ──
    alias = {
        '외국인합계': '외국인', 'FORN': '외국인', 'foreigner': '외국인',
        '기관': '기관합계', 'ORG': '기관합계', 'institution': '기관합계',
        '연기금등': '연기금', 'PENS': '연기금',
        '금융투자': '금융투자', 'FINV': '금융투자',
        '투신': '투신', 'INVM': '투신',
        '사모': '사모', 'PRIV': '사모',
        '보험': '보험', 'INSU': '보험',
        '은행': '은행', 'BANK': '은행',
        '개인': '개인', 'INDV': '개인', 'individual': '개인',
    }
    normalized = {}
    for k, v in inv.items():
        nk = alias.get(k, k)
        normalized[nk] = normalized.get(nk, 0) + v

    # 개인 실측 데이터를 못 가져왔으면(sise_investor.naver 404 등) 기관+외국인으로 추정치 계산
    # 원리: 개인 + 외국인 + 기관 + 기타법인 ≈ 0 (거래소 순매수 총합)
    if not indiv_daily and inv_daily:
        indiv_daily = [
            {
                'date': r['date'],
                'indiv': -(_parse_int(r.get('organ')) + _parse_int(r.get('frgn'))),
                'estimated': True,
            }
            for r in inv_daily
        ]

    return {'frgnRatio': frgn_ratio, 'investors': normalized, 'invDaily': inv_daily, 'indivDaily': indiv_daily,
            'name': stock_name, 'dividend': dividend, 'dividendYieldRatio': div_yield}



def fetch_financial_summary(code: str) -> dict:
    """
    EPS / PER 조회: pykrx → 네이버증권 main → coinfo 순서로 시도
    반환: {'eps': float|None, 'per': float|None}
    캐시 TTL: 1시간
    """
    empty = {'eps': None, 'per': None}

    with _fin_lock:
        cached = _fin_cache.get(code)
        if cached and (time.time() - cached['ts']) < FIN_TTL:
            return cached['data']

    def _save(result):
        with _fin_lock:
            _fin_cache[code] = {'data': result, 'ts': time.time()}
        _log(f"  [재무] {code}: EPS={result['eps']}, PER={result['per']}")
        return result

    # ① pykrx fundamental API는 사내 프록시에서 차단 → 건너뜀
    # (get_market_fundamental_by_date 반복 실패 → 속도 저하 원인)

    # ② 네이버증권 main/coinfo 스크래핑
    def to_f(s):
        if not s: return None
        s = s.strip().replace(',', '')
        try: return float(s)
        except: return None

    def parse_html(html):
        eps, per = None, None
        # id 방식: <em id="_eps">9,749</em>
        for pat, key in [
            (r'''id=["']_eps["'][^>]*>\s*([\d,.\-]+)\s*<''', 'eps'),
            (r'''id=["']_per["'][^>]*>\s*([\d,.\-]+)\s*<''', 'per'),
        ]:
            m = re.search(pat, html, re.I)
            if m:
                v = to_f(m.group(1))
                if key == 'eps': eps = v
                else: per = v
        # 텍스트 테이블 방식: EPS X,XXX원 / PER XX.XX배
        if eps is None:
            m = re.search(r'EPS[^\d]+([\d,]+)\s*원', html)
            if m: eps = to_f(m.group(1))
        if per is None:
            m = re.search(r'PER[^\d]+([\d,.]+)\s*배', html)
            if m: per = to_f(m.group(1))
        return eps, per

    urls = [
        f"https://finance.naver.com/item/main.naver?code={code}",
        f"https://finance.naver.com/item/coinfo.naver?code={code}",
    ]
    for url in urls:
        try:
            r = SESSION.get(url, headers={**HEADERS, "Accept": "text/html,*/*"}, timeout=TIMEOUT)
            if r.status_code != 200: continue
            eps, per = parse_html(r.text)
            if eps or per:
                return _save({'eps': eps, 'per': per})
        except Exception as e:
            _log(f"  [재무 ERR] {code} {url[-30:]}: {e}")

    _log(f"  [재무] {code}: 데이터 없음")
    return empty


def fetch_consensus_data(code: str) -> dict:
    """
    네이버 증권 투자의견 컨센서스 집계 스크래핑
    URL: https://finance.naver.com/item/coinfo.naver?code={code}&target=cnc
    반환:
      opinionScore    : float|None  투자의견 집계 점수 (1.00~5.00)
      targetPrice     : int|None    목표주가 평균(원)
      consensusEps    : float|None  EPS 컨센서스
      consensusPer    : float|None  PER 컨센서스
      institutionCount: int|None    추정기관수
      baseDate        : str|None    기준일 (YYYY.MM.DD)
      consensusAvgTarget: int|None  하위호환
      consensus       : []          하위호환 (빈 배열)
    캐시 TTL: 1시간
    """
    empty = {
        'opinionScore': None, 'targetPrice': None,
        'consensusEps': None, 'consensusPer': None,
        'institutionCount': None, 'baseDate': None,
        'consensusAvgTarget': None, 'consensus': [],
    }

    with _cs_lock:
        cached = _cs_cache.get(code)
        if cached and (time.time() - cached['ts']) < CS_TTL:
            return cached['data']

    pc_h = {**HEADERS, "Accept": "text/html,*/*"}
    url  = f"https://finance.naver.com/item/coinfo.naver?code={code}&target=cnc"

    try:
        r = SESSION.get(url, headers=pc_h, timeout=TIMEOUT)
        _log(f"  [컨센서스] {code} → HTTP {r.status_code} ({len(r.text)}bytes)")
        if r.status_code != 200:
            return empty
        html = r.text

        result = dict(empty)

        # ── 기준일 ──────────────────────────────────────────────
        dm = re.search(r'\[기준[:\s]*([\d]{4}\.[\d]{2}\.[\d]{2})\]', html)
        if dm:
            result['baseDate'] = dm.group(1)

        # ── 투자의견|목표주가 박스 파싱 ──────────────────────────
        # 현재 네이버 구조: <caption>투자의견</caption> ... 목표주가</th><td>
        #   <span class="f_up"><em>4.00</em>매수</span><span class="bar">|</span><em>141,050</em>
        # </td>  (데이터 없으면 <em>N/A</em>)
        def _to_f(s):
            if not s: return None
            s = str(s).strip().replace(',', '').replace('\xa0', '').replace(' ', '')
            if not s or s in ('-', '—', 'N/A'): return None
            try: return float(s)
            except: return None

        tbl_blk = re.search(
            r'투자의견\s*</caption>[\s\S]{0,300}?목표주가\s*</th>\s*<td[^>]*>([\s\S]{0,400}?)</td>',
            html
        )
        if not tbl_blk:
            idx = html.find('투자의견')
            if idx == -1:
                _log(f"  [컨센서스디버그] {code}: '투자의견' 텍스트 자체가 페이지에 없음")
            else:
                snippet = re.sub(r'\s+', ' ', html[idx:idx+700])
                _log(f"  [컨센서스디버그] {code} 투자의견 주변: {snippet}")
        if tbl_blk:
            ems = re.findall(r'<em>([^<]*)</em>', tbl_blk.group(1))
            ems = [re.sub(r'\s+', '', e) for e in ems]
            _log(f"  [컨센서스] {code} ems={ems}")
            if len(ems) >= 2:
                result['opinionScore'] = _to_f(ems[0])
                v1 = _to_f(ems[1])
                result['targetPrice']  = int(v1) if v1 else None
                result['consensusAvgTarget'] = result['targetPrice']

        if any(v is not None for v in [
            result['opinionScore'], result['targetPrice'], result['institutionCount']
        ]):
            with _cs_lock:
                _cs_cache[code] = {'data': result, 'ts': time.time()}
            _log(f"  [컨센서스OK] {code}: 의견={result['opinionScore']}, "
                 f"목표={result['targetPrice']}, 기관={result['institutionCount']}, "
                 f"기준={result['baseDate']}")
            return result

        _log(f"  [컨센서스] {code}: 데이터 없음")
    except Exception as e:
        _log(f"  [컨센서스ERR] {code}: {type(e).__name__}: {e}")

    return empty


# ── 시장 지수 (코스피/코스닥/환율/해외지수/비트코인) ─────────────
_mkt_cache: dict = {'ts': 0, 'data': None}
_mkt_lock  = threading.Lock()
_mkt_logged: set = set()   # 원본 응답 디버그 로그 1회만 찍기 위한 dedup
MKT_TTL = 30                # 30초 캐시

def _mkt_num(s):
    if s is None: return None
    try:
        return float(str(s).replace(',', '').replace(' ', '').replace('%', '').replace('원', ''))
    except (TypeError, ValueError):
        return None

def _mkt_log_once(key, msg):
    if key not in _mkt_logged:
        _mkt_logged.add(key)
        _log(msg)

def _clean_snippet(html, length=3000, anchor=None):
    """디버그용: <script>/<style> 걷어내고 실제 보이는 마크업만 남긴 스니펫.
       anchor가 주어지고 본문에서 발견되면 그 지점부터 잘라서 반환 (헤드/내비 낭비 방지)."""
    c = re.sub(r'<script[\s\S]*?</script>', '', html)
    c = re.sub(r'<style[\s\S]*?</style>', '', c)
    c = re.sub(r'\s+', ' ', c)
    if anchor:
        idx = c.find(anchor)
        if idx != -1:
            start = max(0, idx - 200)
            return c[start:start + length]
    return c[:length]

def _digit_spans_to_str(fragment):
    """<span class="noN">N</span> / <span class="jum">.</span> 자릿수 span 시퀀스를 숫자 문자열로 변환 (콤마=shim은 무시)"""
    parts = re.findall(r'<span class="no\d">(\d)</span>|<span class="jum">(\.)</span>', fragment)
    s = ''.join(a or b for a, b in parts)
    return s or None

def _parse_naver_digitspan_quote(html):
    """네이버 구형 '자릿수 span' 시세 위젯 파싱 (환율 페이지 등에서 사용, id=_nowVal 없음)
       예: <p class="no_today"> <em class="no_down"> <span class="no_down">
             <span class="no1">1</span><span class="shim">,</span><span class="no3">3</span>...
           </span> </em> ... </p>
    """
    val = chg = pct = None

    m = re.search(r'<p class="no_today">([\s\S]{0,500}?)</p>', html)
    if m:
        s = _digit_spans_to_str(m.group(1))
        if s:
            try: val = float(s)
            except ValueError: pass

    m2 = re.search(r'<p class="no_exday">([\s\S]{0,700}?)</p>', html)
    if m2:
        block = m2.group(1)
        em_m = re.search(r'<em class="no_(up|down)">([\s\S]{0,250}?)</em>', block)
        if em_m:
            sign = -1 if em_m.group(1) == 'down' else 1
            s = _digit_spans_to_str(em_m.group(2))
            if s:
                try: chg = sign * float(s)
                except ValueError: pass
        pm = re.search(r'parenthesis1[\s\S]{0,300}?parenthesis2', block)
        if pm:
            seg = pm.group(0)
            sign2 = -1 if 'minus' in seg else 1
            s2 = _digit_spans_to_str(seg)
            if s2:
                try: pct = sign2 * float(s2)
                except ValueError: pass

    return val, chg, pct

def _parse_generic_quote(html):
    """네이버 금융 공통 시세 위젯 패턴 여러 개를 시도 (id=_nowVal 계열 → class=num 계열 → 자릿수 span 계열)"""
    val = chg = pct = None

    m = re.search(r'id=["\']_nowVal["\'][^>]*>\s*(?:<[^>]*>\s*)?([\d,]+\.?\d*)', html)
    if not m:
        m = re.search(r'<span class="num">\s*([\d,]+\.?\d*)\s*</span>', html)
    if m:
        val = _mkt_num(m.group(1))

    if val is not None:
        m2 = re.search(r'id=["\']_change["\'][^>]*>([\s\S]{0,80}?)([\d,]+\.?\d*)', html)
        if m2:
            raw = _mkt_num(m2.group(2))
            if raw is not None:
                chg = -raw if re.search(r'down|하락|dn|_ico_dn|red|fall', m2.group(0), re.I) and \
                              not re.search(r'up|상승|rise', m2.group(0), re.I) else raw
        if chg is None:
            m3 = re.search(r'<span class="range">\s*([\-+]?[\d,]+\.?\d*)\s*</span>', html)
            if m3: chg = _mkt_num(m3.group(1))

        m4 = re.search(r'id=["\']_rate["\'][^>]*>([\s\S]{0,80}?)([\-+]?[\d,]+\.?\d*)\s*%', html)
        if m4:
            pct = _mkt_num(m4.group(2))
        if pct is None:
            m5 = re.search(r'<span class="rate">\s*\(?([\-+]?[\d,]+\.?\d*)%?\)?\s*</span>', html)
            if m5: pct = _mkt_num(m5.group(1))
        if pct is not None and chg is not None and ((pct > 0) != (chg > 0)) and chg != 0:
            pct = -pct  # 부호 불일치 보정
        return val, chg, pct

    # id=_nowVal / class=num 둘 다 실패 → 자릿수 span 템플릿 시도
    return _parse_naver_digitspan_quote(html)

def _mstock_num_from(j: dict, keys: tuple):
    for k in keys:
        v = j.get(k)
        if v is None: continue
        n = _mkt_num(v)
        if n is not None: return n
    return None

def _fetch_mstock_json(url: str, dbg_key: str):
    """m.stock.naver.com 모바일 API 공통 시도 (종목 시세 조회에서 이미 검증된 방식과 동일 패턴).
       성공 시 (val, chg, pct) / 실패 시 (None, None, None). 원본 JSON은 1회만 디버그 로그."""
    try:
        r = SESSION.get(url, headers=HEADERS, timeout=8)
        if r.status_code != 200:
            _mkt_log_once(dbg_key, f"  [지수디버그] mstock {dbg_key} HTTP {r.status_code} url={url}")
            return None, None, None
        j = r.json()
        _mkt_log_once(dbg_key + '_raw', f"  [지수디버그] mstock {dbg_key} raw={str(j)[:600]}")
        val = _mstock_num_from(j, ('closePrice','now','nowValue','indexValue','currentValue','tradePrice','stockEndPrice'))
        chg = _mstock_num_from(j, ('compareToPreviousClosePrice','changeValue','fluctuations','changePrice'))
        pct = _mstock_num_from(j, ('fluctuationsRatio','changeRate','risingRate','fluctuationsRate'))
        return val, chg, pct
    except Exception as e:
        _log(f"  [지수ERR] mstock {dbg_key}: {type(e).__name__}: {e}")
        return None, None, None

def _extract_realtime_datum(j):
    """polling.finance.naver.com 응답에서 실제 시세가 담긴 dict 하나를 최대한 유연하게 찾아냄.
       네이버가 종종 중첩 구조를 바꿔서(datas 최상위 / result.areas[0].datas / result.datas 등)
       여러 형태를 다 시도. closePrice류 키가 하나라도 있으면 그 dict를 실제 시세로 판단."""
    def _has_price_key(d):
        return isinstance(d, dict) and any(k in d for k in
            ('closePrice','now','nowValue','indexValue','currentValue','tradePrice','stockEndPrice'))
    candidates = []
    if isinstance(j, list) and j:
        candidates.append(j[0])
    if isinstance(j, dict):
        if isinstance(j.get('datas'), list) and j['datas']:
            candidates.append(j['datas'][0])
        res = j.get('result')
        if isinstance(res, dict):
            areas = res.get('areas')
            if isinstance(areas, list) and areas and isinstance(areas[0], dict):
                d0 = areas[0].get('datas')
                if isinstance(d0, list) and d0:
                    candidates.append(d0[0])
            if isinstance(res.get('datas'), list) and res['datas']:
                candidates.append(res['datas'][0])
            candidates.append(res)
        elif isinstance(res, list) and res:
            candidates.append(res[0])
        candidates.append(j)
    for c in candidates:
        if _has_price_key(c):
            return c
    return None

def _fetch_polling_realtime(url: str, dbg_key: str):
    """polling.finance.naver.com/api/realtime/... 공통 시도 (m.stock API가 세계지수/환율에서
       400/404로 막힌 뒤 대체용으로 추가한 네이버의 또 다른 실시간 시세 API).
       성공 시 (val, chg, pct) / 실패 시 (None, None, None)."""
    try:
        r = SESSION.get(url, headers=HEADERS, timeout=8)
        if r.status_code != 200:
            _mkt_log_once(dbg_key, f"  [지수디버그] polling {dbg_key} HTTP {r.status_code} url={url}")
            return None, None, None
        j = r.json()
        _mkt_log_once(dbg_key + '_raw', f"  [지수디버그] polling {dbg_key} raw={str(j)[:600]}")
        d = _extract_realtime_datum(j)
        if not d:
            return None, None, None
        val = _mstock_num_from(d, ('closePrice','now','nowValue','indexValue','currentValue','tradePrice','stockEndPrice'))
        chg = _mstock_num_from(d, ('compareToPreviousClosePrice','changeValue','fluctuations','changePrice'))
        pct = _mstock_num_from(d, ('fluctuationsRatio','changeRate','risingRate','fluctuationsRate'))
        return val, chg, pct
    except Exception as e:
        _log(f"  [지수ERR] polling {dbg_key}: {type(e).__name__}: {e}")
        return None, None, None

def fetch_domestic_index(code: str, key: str) -> dict:
    """코스피/코스닥 실시간 지수 (모바일 JSON API 우선 → 실패 시 HTML 페이지 백업)"""
    out = {key: {'value': None, 'change': None, 'changePct': None, 'ok': False}}

    val, chg, pct = _fetch_mstock_json(f"https://m.stock.naver.com/api/index/{code}/basic", key)
    if val is not None:
        out[key] = {'value': val, 'change': chg, 'changePct': pct, 'ok': True}
        return out

    try:
        url = f"https://finance.naver.com/sise/sise_index.naver?code={code}"
        r = SESSION.get(url, headers={**HEADERS, "Accept": "text/html,*/*"}, timeout=8)
        html = r.text
        val, chg, pct = _parse_generic_quote(html)
        if val is None:
            _mkt_log_once(key + '_html', f"  [지수디버그] domestic {key}({code}) HTTP {r.status_code} "
                                          f"len={len(html)} snippet={_clean_snippet(html)}")
        else:
            out[key] = {'value': val, 'change': chg, 'changePct': pct, 'ok': True}
    except Exception as e:
        _log(f"  [지수ERR] domestic {key}: {type(e).__name__}: {e}")
    return out

def _fetch_marketindex_prices(reuters_code: str, dbg_key: str):
    """m.stock.naver.com의 신형 front-api (2026년 리뉴얼 이후 환율 데이터가 옮겨간 것으로 보이는 경로).
       기존 /api/marketindex/exchange/... 가 계속 404가 나서 대체용으로 추가."""
    try:
        url = (f"https://m.stock.naver.com/front-api/v1/marketIndex/prices"
               f"?category=exchange&reutersCode={reuters_code}&page=1&pageSize=1")
        r = SESSION.get(url, headers=HEADERS, timeout=8)
        if r.status_code != 200:
            _mkt_log_once(dbg_key, f"  [지수디버그] marketIndex {dbg_key} HTTP {r.status_code} url={url}")
            return None, None, None
        j = r.json()
        _mkt_log_once(dbg_key + '_raw', f"  [지수디버그] marketIndex {dbg_key} raw={str(j)[:600]}")
        result = j.get('result')
        first = None
        if isinstance(result, list) and result:
            first = result[0]
        elif isinstance(result, dict):
            for v in result.values():
                if isinstance(v, list) and v and isinstance(v[0], dict):
                    first = v[0]; break
        if not isinstance(first, dict):
            return None, None, None
        val = _mstock_num_from(first, ('closePrice', 'value'))
        chg = _mstock_num_from(first, ('compareToPreviousClosePrice', 'fluctuations', 'changeValue'))
        pct = _mstock_num_from(first, ('fluctuationsRatio', 'changeRate'))
        return val, chg, pct
    except Exception as e:
        _log(f"  [지수ERR] marketIndex {dbg_key}: {type(e).__name__}: {e}")
        return None, None, None

def _fetch_frankfurter_fx():
    """ECB 기준 USD/KRW 환율 (frankfurter.app - 무료·인증 불필요·안정적인 공개 API).
       네이버 쪽 환율 경로가 전부(4가지나) 막혔을 때의 최종 대안. 네이버의 실시간 시장환율이
       아니라 ECB 기준환율(평일 16:00 CET 갱신, 주말엔 금요일자 값 유지)이라 몇 원 정도
       차이날 수 있지만, 최소한 "—"로 안 뜨고 대략적인 환율/등락은 보여줄 수 있음."""
    try:
        r = SESSION.get("https://api.frankfurter.app/latest?from=USD&to=KRW", timeout=8)
        if r.status_code != 200:
            _mkt_log_once('fx_frank', f"  [지수디버그] frankfurter FX HTTP {r.status_code}")
            return None, None, None
        j = r.json()
        today = j.get('date')
        val = _mkt_num((j.get('rates') or {}).get('KRW'))
        if val is None:
            val = _mkt_num(j.get('rate'))  # v2 스타일 응답 대비
        if val is None:
            return None, None, None
        chg = pct = None
        try:
            if today:
                prev_day = (datetime.strptime(today, '%Y-%m-%d') - timedelta(days=1)).strftime('%Y-%m-%d')
                r2 = SESSION.get(f"https://api.frankfurter.app/{prev_day}..{today}?from=USD&to=KRW", timeout=6)
                if r2.status_code == 200:
                    rates_map = (r2.json() or {}).get('rates') or {}
                    dates = sorted(rates_map.keys())
                    if len(dates) >= 2:
                        prev_val = _mkt_num((rates_map[dates[0]] or {}).get('KRW'))
                        if prev_val:
                            chg = val - prev_val
                            pct = chg / prev_val * 100
        except Exception:
            pass  # 전일 대비 계산 실패해도 현재값은 살림
        return val, chg, pct
    except Exception as e:
        _log(f"  [지수ERR] frankfurter FX: {type(e).__name__}: {e}")
        return None, None, None

def fetch_exchange_rate() -> dict:
    """달러/원 환율 (모바일 JSON API 우선 → 실패 시 HTML 페이지 백업).
       네이버 개편 이후 기존 /api/marketindex/exchange 경로가 404가 나서, 대체 경로 두 개를 더 시도한다."""
    out = {'usdkrw': {'value': None, 'change': None, 'changePct': None, 'ok': False}}

    val, chg, pct = _fetch_mstock_json("https://m.stock.naver.com/api/marketindex/exchange/FX_USDKRW", 'fx')
    if val is not None:
        out['usdkrw'] = {'value': val, 'change': chg, 'changePct': pct, 'ok': True}
        return out

    for fx_url, fx_key in (
        ("https://polling.finance.naver.com/api/realtime/marketindicator/FX_USDKRW", 'fx_poll'),
        ("https://polling.finance.naver.com/api/realtime/marketindicator/exchange/FX_USDKRW", 'fx_poll2'),
    ):
        val, chg, pct = _fetch_polling_realtime(fx_url, fx_key)
        if val is not None:
            out['usdkrw'] = {'value': val, 'change': chg, 'changePct': pct, 'ok': True}
            return out

    val, chg, pct = _fetch_marketindex_prices("FX_USDKRW", 'fx_mip')
    if val is not None:
        out['usdkrw'] = {'value': val, 'change': chg, 'changePct': pct, 'ok': True}
        return out

    try:
        url = "https://finance.naver.com/marketindex/exchangeDetail.naver?marketindexCd=FX_USDKRW"
        r = SESSION.get(url, headers={**HEADERS, "Accept": "text/html,*/*"}, timeout=8)
        html = r.text
        val, chg, pct = _parse_generic_quote(html)
        if val is None:
            _mkt_log_once('fx_html', f"  [지수디버그] FX HTTP {r.status_code} "
                                      f"len={len(html)} snippet={_clean_snippet(html)}")
        else:
            out['usdkrw'] = {'value': val, 'change': chg, 'changePct': pct, 'ok': True}
            return out
    except Exception as e:
        _log(f"  [지수ERR] FX: {type(e).__name__}: {e}")

    # 네이버 쪽 경로 4개가 전부 실패한 경우의 최종 대안 — 안정적인 무료 공개 API
    val, chg, pct = _fetch_frankfurter_fx()
    if val is not None:
        out['usdkrw'] = {'value': val, 'change': chg, 'changePct': pct, 'ok': True}
    return out

def fetch_btc() -> dict:
    """비트코인 원화 시세 (Upbit 공개 API)"""
    out = {}
    try:
        url = "https://api.upbit.com/v1/ticker?markets=KRW-BTC"
        r = SESSION.get(url, headers={"Accept": "application/json"}, timeout=8)
        j = r.json()
        if isinstance(j, list) and j:
            d = j[0]
            val = _mkt_num(d.get('trade_price'))
            chg = _mkt_num(d.get('signed_change_price'))
            pct = _mkt_num(d.get('signed_change_rate'))
            if pct is not None: pct *= 100
            out['btc'] = {'value': val, 'change': chg, 'changePct': pct, 'ok': val is not None}
        else:
            _log(f"  [지수ERR] BTC: 데이터 없음 raw={str(j)[:300]}")
    except Exception as e:
        _log(f"  [지수ERR] BTC: {type(e).__name__}: {e}")
    return out

# 해외지수 심볼 후보 (네이버 world 페이지 기준) + 실패 시 목록 페이지에서 자동 탐색할 검색어
WORLD_SYMS = {
    'nasdaq': {'symbols': ['NAS@IXIC'],                        'discover': '나스닥종합',
               'mstock': ['.IXIC', 'IXIC'],                    'polling': ['.IXIC', 'NAS@IXIC']},
    'sp500':  {'symbols': ['SPI@SPX'],                         'discover': 'S&amp;P500',
               'mstock': ['.INX', '.SPX', 'SPX'],              'polling': ['.INX', '.SPX', 'SPI@SPX']},
    'vix':    {'symbols': ['CBO@VIX', 'CBOE@VIX'],             'discover': 'VIX',
               'mstock': ['.VIX', 'VIX'],                      'stooq': '^vix',
               'polling': ['.VIX', 'CBO@VIX'],
               'tradingview': 'CBOE:VIX'},
}

def fetch_tradingview_index(symbol: str, key: str) -> dict:
    """TradingView 스캐너 API (네이버·Stooq 다 실패한 지수용 최종 폴백)
       symbol 예: 'CBOE:VIX' (거래소:심볼 형식)"""
    out = {key: {'value': None, 'change': None, 'changePct': None, 'ok': False}}
    try:
        url = "https://scanner.tradingview.com/america/scan"
        payload = {
            "symbols": {"tickers": [symbol], "query": {"types": []}},
            "columns": ["close", "change", "change_abs"],
        }
        r = SESSION.post(url, json=payload,
                          headers={**HEADERS, "Content-Type": "application/json", "Accept": "application/json"},
                          timeout=8)
        if r.status_code == 200:
            j = r.json()
            data = j.get('data') or []
            d = (data[0].get('d') if data else None) or []
            close = _mkt_num(d[0]) if len(d) > 0 else None
            pct   = _mkt_num(d[1]) if len(d) > 1 else None
            chg   = _mkt_num(d[2]) if len(d) > 2 else None
            if close is not None:
                out[key] = {'value': close, 'change': chg, 'changePct': pct, 'ok': True}
            else:
                _mkt_log_once(f'tv_{key}', f"  [지수디버그] tradingview {key}({symbol}) raw={str(j)[:500]}")
        else:
            _mkt_log_once(f'tv_{key}', f"  [지수디버그] tradingview {key}({symbol}) HTTP {r.status_code} body={r.text[:400]}")
    except Exception as e:
        _log(f"  [지수ERR] tradingview {key}: {type(e).__name__}: {e}")
    return out

def fetch_stooq_index(symbol: str, key: str) -> dict:
    """Stooq 무료 CSV 시세 (네이버에서 심볼을 못 찾은 지수용 최종 폴백 - 독립적인 별도 소스)"""
    out = {key: {'value': None, 'change': None, 'changePct': None, 'ok': False}}
    try:
        url = f"https://stooq.com/q/l/?s={symbol}&f=sd2t2ohlcv&h&e=csv"
        r = SESSION.get(url, timeout=8)
        lines = r.text.strip().splitlines()
        if len(lines) >= 2:
            header = [h.strip() for h in lines[0].split(',')]
            row    = [c.strip() for c in lines[1].split(',')]
            d = dict(zip(header, row))
            close = _mkt_num(d.get('Close'))
            openp = _mkt_num(d.get('Open'))
            if close is not None and close > 0:
                chg = (close - openp) if openp else None
                pct = (chg / openp * 100) if (chg is not None and openp) else None
                out[key] = {'value': close, 'change': chg, 'changePct': pct, 'ok': True}
            else:
                _mkt_log_once(f'stooq_{key}', f"  [지수디버그] stooq {key}({symbol}) body={r.text[:300]}")
        else:
            _mkt_log_once(f'stooq_{key}', f"  [지수디버그] stooq {key}({symbol}) HTTP {r.status_code} body={r.text[:300]}")
    except Exception as e:
        _log(f"  [지수ERR] stooq {key}: {type(e).__name__}: {e}")
    return out

_world_sym_cache: dict = {}  # 검색어 -> 발견한 심볼 (or None), 프로세스당 1회만 탐색

def _discover_world_symbol(search_text: str):
    """네이버 해외증시 목록 페이지(world/)에서 심볼 코드 자동 탐색 (하드코딩 후보가 다 실패했을 때)"""
    if search_text in _world_sym_cache:
        return _world_sym_cache[search_text]
    result = None
    try:
        url = "https://finance.naver.com/world/"
        r = SESSION.get(url, headers={**HEADERS, "Accept": "text/html,*/*"}, timeout=8)
        html = r.text
        idx = html.find(search_text)
        if idx != -1:
            window = html[max(0, idx - 600):idx + 200]
            m = re.search(r'symbol=([A-Za-z0-9%\.@]+)', window)
            if m:
                result = unquote(m.group(1))
        if result is None:
            _log(f"  [지수디버그] world목록 '{search_text}' 못찾음 len={len(html)} "
                 f"snippet={_clean_snippet(html, 1500)}")
        else:
            _log(f"  [지수] world목록 '{search_text}' → 심볼 발견: {result}")
    except Exception as e:
        _log(f"  [지수ERR] world목록 '{search_text}': {type(e).__name__}: {e}")
    _world_sym_cache[search_text] = result
    return result

def fetch_world_index(key: str, cfg: dict) -> dict:
    """해외지수: 모바일 JSON API 우선 시도 → 실패 시 world 시세 페이지 HTML 스크래핑(심볼 후보 순차 → 자동 탐색)"""
    out = {key: {'value': None, 'change': None, 'changePct': None, 'ok': False}}

    # 네이버 개편 이후 기존 /api/index/{code}/basic 이 세계지수에서 계속 400이 나서,
    # 폴링(실시간) API를 먼저 시도 — 이게 현재 네이버 앱이 실제로 쓰는 걸로 보이는 경로.
    for pcode in cfg.get('polling', []):
        val, chg, pct = _fetch_polling_realtime(f"https://polling.finance.naver.com/api/realtime/worldstock/index/{pcode}", f"{key}_poll")
        if val is not None:
            out[key] = {'value': val, 'change': chg, 'changePct': pct, 'ok': True}
            return out

    for mcode in cfg.get('mstock', []):
        val, chg, pct = _fetch_mstock_json(f"https://m.stock.naver.com/api/index/{mcode}/basic", f"{key}_m")
        if val is not None:
            out[key] = {'value': val, 'change': chg, 'changePct': pct, 'ok': True}
            return out

    symbols = list(cfg['symbols'])

    def _try(symbol):
        r = SESSION.get(f"https://finance.naver.com/world/sise.naver?symbol={symbol}",
                         headers={**HEADERS, "Accept": "text/html,*/*"}, timeout=8)
        html = r.text
        if '존재하지 않는 종목' in html or '존재하지 않습니다' in html:
            return None, r.status_code, html
        val, chg, pct = _parse_generic_quote(html)
        if val is None:
            val_m = re.search(r'<span class="num">\s*([\d,]+\.?\d*)\s*</span>', html)
            chg_m = re.search(r'<span class="range">\s*([\-+]?[\d,]+\.?\d*)\s*</span>', html)
            pct_m = re.search(r'<span class="rate">\s*\(?([\-+]?[\d,]+\.?\d*)%?\)?\s*</span>', html)
            if val_m:
                val = _mkt_num(val_m.group(1))
                chg = _mkt_num(chg_m.group(1)) if chg_m else chg
                pct = _mkt_num(pct_m.group(1)) if pct_m else pct
        return (val, chg, pct) if val is not None else None, r.status_code, html

    last_html, last_status, last_sym = '', None, symbols[0]
    for symbol in symbols:
        try:
            found, status, html = _try(symbol)
            last_html, last_status, last_sym = html, status, symbol
            if found:
                val, chg, pct = found
                out[key] = {'value': val, 'change': chg, 'changePct': pct, 'ok': True}
                return out
        except Exception as e:
            _log(f"  [지수ERR] world {key}({symbol}): {type(e).__name__}: {e}")

    # 하드코딩 후보 전부 실패 → 목록 페이지에서 자동 탐색해 한 번 더 시도
    discovered = _discover_world_symbol(cfg['discover'])
    if discovered and discovered not in symbols:
        try:
            found, status, html = _try(discovered)
            last_html, last_status, last_sym = html, status, discovered
            if found:
                val, chg, pct = found
                out[key] = {'value': val, 'change': chg, 'changePct': pct, 'ok': True}
                return out
        except Exception as e:
            _log(f"  [지수ERR] world {key}({discovered}): {type(e).__name__}: {e}")

    # 네이버 쪽이 전부 실패 → Stooq(독립 소스)로 시도
    if cfg.get('stooq'):
        stooq_out = fetch_stooq_index(cfg['stooq'], key)
        if stooq_out[key]['ok']:
            return stooq_out

    # Stooq도 실패 → TradingView 스캐너 API로 최종 폴백
    if cfg.get('tradingview'):
        tv_out = fetch_tradingview_index(cfg['tradingview'], key)
        if tv_out[key]['ok']:
            return tv_out

    _snip = _clean_snippet(last_html, 4000, anchor='id="content"')
    _mkt_log_once(key, f"  [지수디버그] world {key}(마지막시도={last_sym}) HTTP {last_status} "
                        f"len={len(last_html)} snippet={_snip}")
    return out

def fetch_market_indices() -> dict:
    """전체 시장 지수 묶음 조회 (30초 캐시)"""
    with _mkt_lock:
        cached = _mkt_cache.get('data')
        if cached and (time.time() - _mkt_cache['ts']) < MKT_TTL:
            return cached

    result = {}
    ex = _shared_pool  # 공유 풀 사용 (per-call 풀 생성 시 shutdown(wait=True)로 블록되는 문제 방지)
    futs = {
        ex.submit(fetch_domestic_index, 'KOSPI', 'kospi'):   'kospi',
        ex.submit(fetch_domestic_index, 'KOSDAQ', 'kosdaq'): 'kosdaq',
        ex.submit(fetch_exchange_rate):                      'fx',
        ex.submit(fetch_btc):                                'btc',
    }
    for key, syms in WORLD_SYMS.items():
        futs[ex.submit(fetch_world_index, key, syms)] = key

    # 주의: 예전엔 이 내부 대기시간이 프론트엔드 fetch의 abort 시간(15초)과 똑같았음.
    # 서버가 15초를 꽉 채워서 응답하면 그 사이 클라이언트가 이미 요청을 중단해버려서
    # 응답 자체를 못 받고(= renderMktTicker가 아예 호출 안 됨) 지수 바 전체가 사라지는
    # 문제가 있었음. 오늘 폴백 경로를 추가하면서 지연이 조금 더 늘어나 이 경합이 더 자주
    # 발생한 것으로 보임 — 내부 대기를 10초로 줄여서 프론트(20초)보다 확실히 먼저 끝나게 함.
    try:
        for fut in as_completed(futs, timeout=10):
            try:
                result.update(fut.result(timeout=1) or {})
            except Exception as e:
                _log(f"  [지수ERR] {futs.get(fut)}: {type(e).__name__}: {e}")
    except FutureTimeoutError:
        _log("  [지수ERR] market_indices 전체 10초 타임아웃 - 완료된 것만 반영")

    with _mkt_lock:
        _mkt_cache['data'] = result
        _mkt_cache['ts'] = time.time()
    return result


def search_stock_name(query: str) -> list:
    """
    종목명/코드 → 코드 조회 (시스템 프록시 사용, 빠른 fallback)
    반환: [{'code': '005930', 'name': '삼성전자'}, ...]
    """
    results = []

    # 6자리 코드 직접 입력 처리
    clean_q = query.strip()
    if re.fullmatch(r'\d{5,6}', clean_q):
        code = clean_q.zfill(6)
        try:
            url = f"https://m.stock.naver.com/api/stock/{code}/basic"
            r = SESSION.get(url, headers=HEADERS, timeout=4)
            if r.status_code == 200:
                j = r.json()
                name = j.get('stockName', j.get('name', ''))
                if name:
                    _log(f"  [검색] 코드 직접 조회 {code} → {name}")
                    return [{'code': code, 'name': name}]
        except Exception as e:
            _log(f"[WARN] code direct lookup({code}): {e}")

    # 1) ac.finance.naver.com autocomplete
    try:
        ac_url = (
            f"https://ac.finance.naver.com/ac"
            f"?q={quote(query)}&q_enc=UTF-8&st=111&frq=0&rc=10&r_lt=111"
        )
        ra = SESSION.get(ac_url, headers=HEADERS, timeout=3)
        _log(f"  [검색] ac autocomplete '{query}' → HTTP {ra.status_code} ({len(ra.text)}bytes)")
        if ra.status_code == 200 and ra.text.strip():
            try:
                ac_data = ra.json()
                for group in ac_data.get('items', []):
                    for item in group:
                        if not isinstance(item, list) or len(item) < 2:
                            continue
                        raw_code = str(item[0]).strip()
                        raw_name = str(item[1]).strip()
                        if re.fullmatch(r'\d{4,6}', raw_code) and raw_name:
                            code = raw_code.zfill(6)
                            if code not in [x['code'] for x in results]:
                                results.append({'code': code, 'name': raw_name})
            except Exception:
                for m in re.finditer(r'"(\d{6})","([^"]{1,30})"', ra.text):
                    code = m.group(1)
                    if code not in [x['code'] for x in results]:
                        results.append({'code': code, 'name': m.group(2)})
    except Exception as e:
        _log(f"[WARN] ac.finance search('{query}'): {type(e).__name__}")

    # 2) KRX 한국거래소 종목 검색 API (공식 API, 프록시 우호적)
    if not results:
        try:
            krx_url = "https://data.krx.co.kr/comm/bldAttendant/getJsonData.cmd"
            krx_data = {
                "bld": "dbms/comm/finder/finder_stkisu",
                "mktsel": "ALL",
                "searchText": query,
            }
            krx_h = {
                **HEADERS,
                "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
                "Referer": "https://data.krx.co.kr/contents/MDC/STAT/standard/MDCSTAT01901.cmd",
                "Origin": "https://data.krx.co.kr",
                "X-Requested-With": "XMLHttpRequest",
            }
            rk = SESSION.post(krx_url, data=krx_data, headers=krx_h, timeout=5)
            _log(f"  [검색] KRX finder '{query}' → HTTP {rk.status_code} ({len(rk.text)}bytes)")
            if rk.status_code == 200:
                jk = rk.json()
                # 여러 가능한 키 이름 시도
                rows = (jk.get('block1') or jk.get('OutBlock_1') or
                        jk.get('output') or jk.get('data') or [])
                for item in rows:
                    # KRX 실제 필드: short_code, codeName
                    code = str(item.get('short_code', item.get('short_isin_cd',
                              item.get('ISU_SRT_CD', item.get('code', ''))))).strip()
                    name = str(item.get('codeName', item.get('kor_isin_nm',
                              item.get('ISU_NM', item.get('name', ''))))).strip()
                    if re.fullmatch(r'\d{6}', code) and name:
                        if code not in [x['code'] for x in results]:
                            results.append({'code': code, 'name': name})
        except Exception as e:
            _log(f"[WARN] KRX search('{query}'): {type(e).__name__}: {e}")

    # 3) 네이버 모바일 주식 검색 (여러 엔드포인트 시도)
    if not results:
        mob_urls = [
            f"https://m.stock.naver.com/api/stock/search?query={quote(query)}&page=1&pageSize=10",
            f"https://stock.naver.com/api/search/autocomplete?keyword={quote(query)}",
            f"https://m.stock.naver.com/api/stocks?keyword={quote(query)}&page=1&pageSize=10",
        ]
        for mob_url in mob_urls:
            try:
                mob_h = {**HEADERS, "Referer": "https://m.stock.naver.com/"}
                rm = SESSION.get(mob_url, headers=mob_h, timeout=4)
                ep = mob_url.split('/')[-1].split('?')[0]
                _log(f"  [검색] naver/{ep} '{query}' → HTTP {rm.status_code} ({len(rm.text)}bytes)")
                if rm.status_code == 200:
                    jm = rm.json()
                    items = (jm.get('stocks') or jm.get('items') or jm.get('list') or
                             (jm.get('result', {}) or {}).get('stocks') or
                             (jm.get('data', {}) or {}).get('stocks') or
                             (jm if isinstance(jm, list) else []))
                    for item in (items or []):
                        code = str(item.get('itemCode', item.get('code', item.get('stockCode', '')))).zfill(6)
                        name = item.get('itemName', item.get('name', item.get('stockName', '')))
                        if re.fullmatch(r'\d{6}', code) and name:
                            if code not in [x['code'] for x in results]:
                                results.append({'code': code, 'name': name})
                if results:
                    break
            except Exception as e:
                _log(f"[WARN] naver/{ep} search('{query}'): {type(e).__name__}")

    # 5) 네이버 금융 검색 페이지 HTML (최후 수단)
    if not results:
        try:
            for search_url in [
                f"https://finance.naver.com/search/index.naver?query={quote(query)}",
                f"https://finance.naver.com/search/searchList.naver?query={quote(query)}",
            ]:
                pc_headers = {**HEADERS, "Accept": "text/html,*/*"}
                r = SESSION.get(search_url, headers=pc_headers, timeout=4)
                _log(f"  [검색] HTML '{query}' → HTTP {r.status_code}")
                if r.status_code == 200:
                    for m in re.finditer(
                        r'(?:main\.naver\?code=|code=)(\d{4,6})[^>]*>\s*([^<]{1,30})\s*</a>',
                        r.text
                    ):
                        code = m.group(1).zfill(6)
                        name = m.group(2).strip()
                        if name and code not in [x['code'] for x in results]:
                            results.append({'code': code, 'name': name})
                if results:
                    break
        except Exception as e:
            _log(f"[WARN] HTML search('{query}'): {type(e).__name__}")

    _log(f"  [검색] '{query}' → {len(results)}건: {[r['name'] for r in results[:3]]}")
    return results[:10]


def get_stock_data(code: str) -> dict:
    """종목 데이터 반환 (장중 1분 / 장외 5분 캐시, 장개시 시 프리마켓 캐시 자동 무효화)"""
    mopen = is_market_open()
    ttl   = CACHE_TTL if mopen else CACHE_TTL_OFF

    with _cache_lock:
        cached = _cache.get(code)
        if cached and (time.time() - cached['ts']) < ttl:
            # 장 시작 후 장외에서 만들어진 캐시는 즉시 무효화
            if mopen and not cached['data'].get('_mopen', False):
                pass   # fall through → 새로 조회
            else:
                return cached['data']

    print(f"  조회: {code}", end="", flush=True)
    try:
        # ── 5개 fetch를 동시에 병렬 실행 (컨센서스 목표가 포함) ────
        # 공유 스레드풀 사용: 느린 작업이 있어도 그 작업 때문에 이 요청 전체가
        # 붙잡히지 않음 (개별 result(timeout=N)만큼만 기다리고 반환).
        ex = _shared_pool
        f_candles = ex.submit(fetch_candles, code)
        f_price   = ex.submit(fetch_price_naver, code)
        f_inv     = ex.submit(fetch_investor_data, code)
        f_fin     = ex.submit(fetch_financial_summary, code)
        f_cons    = ex.submit(fetch_consensus_data, code)

        try:
            candles = f_candles.result(timeout=12)
        except Exception as e:
            _log(f"  [캔들 타임아웃] {code}: {type(e).__name__}")
            candles = []
        try:
            nv_px, nv_prev, nv_name = f_price.result(timeout=12)
        except Exception:
            nv_px, nv_prev, nv_name = None, None, ''
        try:
            inv_data = f_inv.result(timeout=12)
        except Exception:
            inv_data = {'frgnRatio': None, 'investors': {}, 'invDaily': [], 'indivDaily': [], 'name': '',
                        'dividend': None, 'dividendYieldRatio': None}
        try:
            fin_data = f_fin.result(timeout=12)
        except Exception:
            fin_data = {'eps': None, 'per': None}
        try:
            cons_data = f_cons.result(timeout=12)
        except Exception:
            cons_data = {
                'opinionScore': None, 'targetPrice': None,
                'consensusEps': None, 'consensusPer': None,
                'institutionCount': None, 'baseDate': None,
                'consensusAvgTarget': None, 'consensus': [],
            }

        today_str = datetime.now().strftime('%Y-%m-%d')

        # ── pykrx가 장중에 오늘 날짜 intraday 캔들을 포함해 반환하는 경우 처리 ──
        # candles[-1] = {date:'오늘', close:현재가(intraday)}
        # 이를 prev_close로 사용하면 등락률이 0%로 잘못 표시됨
        has_today    = bool(candles and candles[-1]['date'] == today_str)
        hist_candles = candles[:-1] if has_today else candles  # 오늘 제외 히스토리

        # last_close = 전일(직전 거래일) 종가
        last_close = hist_candles[-1]['close'] if hist_candles else None
        px, src    = last_close, ("pykrx" if PYKRX and code not in _krx_failed else "Naver fchart")
        prev_close = None   # 전일종가

        if nv_px:
            if mopen:
                px, src = nv_px, "네이버 실시간"
            elif nv_px != last_close:
                # 장 종료 후 pykrx보다 최신이면 덮어씀
                px, src = nv_px, "네이버"
            # nv_prev == nv_px 이면 changePrice=0 오류(장개시 직후 미체결) → 무시
            if nv_prev and nv_prev > 0 and nv_prev != nv_px:
                prev_close = nv_prev   # 가장 신뢰도 높은 전일종가

        # 전일종가 fallback: hist_candles[-1] = 어제(직전 거래일) 종가
        # (오늘 intraday 캔들을 제거했으므로 항상 전일 종가)
        if prev_close is None:
            prev_close = hist_candles[-1]['close'] if hist_candles else None

        p1 = past_price(hist_candles, 1)
        p2 = past_price(hist_candles, 2)
        p3 = past_price(hist_candles, 3)

        def ret(a, b):
            if a and b and b != 0:
                return round((a - b) / b, 6)
            return None

        # 캔들차트용 최근 65개 (3개월 일봉, OHLC 포함, 오늘 intraday 제외)
        spark = hist_candles[-65:] if hist_candles else []

        result = {
            'code': code, 'ok': bool(px), 'src': src,
            'name': inv_data.get('name') or nv_name,
            'px':        px,
            'prevClose': prev_close,
            'isMarketOpen': mopen,
            '_mopen':    mopen,   # 캐시 장외/장중 구분용 (내부)
            'p1':   p1,  'p2':  p2,  'p3':  p3,
            'r1':   ret(px, p1),
            'r2':   ret(px, p2),
            'r3':   ret(px, p3),
            'frgnRatio': inv_data.get('frgnRatio'),
            'dividend': inv_data.get('dividend'),
            'dividendYieldRatio': inv_data.get('dividendYieldRatio'),
            'investors': inv_data.get('investors', {}),
            'invDaily':  inv_data.get('invDaily', []),
            'indivDaily': inv_data.get('indivDaily', []),
            'candles': spark,
            'consensus':          cons_data.get('consensus', []),
            'consensusAvgTarget': cons_data.get('consensusAvgTarget'),
            'opinionScore':       cons_data.get('opinionScore'),
            'targetPrice':        cons_data.get('targetPrice'),
            'consensusEps':       cons_data.get('consensusEps'),
            'consensusPer':       cons_data.get('consensusPer'),
            'institutionCount':   cons_data.get('institutionCount'),
            'baseDate':           cons_data.get('baseDate'),
            'eps':  fin_data.get('eps'),
            'per':  fin_data.get('per'),
        }
        print(f"  → {px:,}원 ({src})" if px else "  → 실패")

    except Exception as e:
        result = {'code': code, 'ok': False, 'src': '조회실패', 'px': None, 'error': str(e)}
        _log(f"[ERR] get_stock_data({code}): {type(e).__name__}: {e}")

    with _cache_lock:
        # 실패한 결과는 캐시 안 함 (다음 요청 시 다시 시도)
        if result.get('ok'):
            _cache[code] = {'data': result, 'ts': time.time()}
    return result


def clear_cache(code: str = None):
    """캐시 초기화 (code=None이면 전체)"""
    global PYKRX_INV_DEAD
    with _cache_lock:
        if code:
            _cache.pop(code, None)
        else:
            _cache.clear()
            _krx_failed.clear()
            PYKRX_INV_DEAD = False  # 다음 갱신 시 pykrx 다시 시도
    with _invd_lock:
        if code:
            _invd_cache.pop(code, None)
        else:
            _invd_cache.clear()


# ════════════════════════════════════════════════════════
# HTTP 서버
# ════════════════════════════════════════════════════════

SCRIPT_DIR     = os.path.dirname(os.path.abspath(__file__))
_idx       = os.path.join(SCRIPT_DIR, 'index.html')
_stk       = os.path.join(SCRIPT_DIR, 'stock_tracker.html')
HTML_PATH  = _idx if os.path.exists(_idx) else _stk
MANIFEST_PATH  = os.path.join(SCRIPT_DIR, 'manifest.json')
SW_PATH        = os.path.join(SCRIPT_DIR, 'sw.js')
PORTFOLIO_PATH = os.path.join(SCRIPT_DIR, 'portfolio.json')

# ── 공유 파일함 (엑셀/PDF 등을 노트북↔데스크탑↔폰 간에 공유) ─────
SHARED_DIR = os.path.join(SCRIPT_DIR, 'shared_files')
os.makedirs(SHARED_DIR, exist_ok=True)
SHARED_ALLOWED_EXT = {'.xlsx', '.xls', '.csv', '.pdf', '.docx', '.txt', '.png', '.jpg', '.jpeg'}

def _safe_shared_name(name: str) -> str:
    """경로 조작(../ 등) 방지 — 파일명만 남기고 확장자 검증."""
    name = os.path.basename(unquote(name or '')).strip()
    if not name or name in ('.', '..'):
        raise ValueError('invalid filename')
    ext = os.path.splitext(name)[1].lower()
    if ext not in SHARED_ALLOWED_EXT:
        raise ValueError(f'허용되지 않는 확장자: {ext}')
    return name

def list_shared_files() -> list:
    out = []
    try:
        for fn in sorted(os.listdir(SHARED_DIR)):
            fp = os.path.join(SHARED_DIR, fn)
            if os.path.isfile(fp):
                st = os.stat(fp)
                out.append({'name': fn, 'size': st.st_size,
                            'mtime': datetime.fromtimestamp(st.st_mtime).strftime('%Y-%m-%d %H:%M')})
    except Exception as e:
        _log(f'list_shared_files error: {e}')
    return out

_SHARED_MIME = {
    '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    '.xls':  'application/vnd.ms-excel',
    '.csv':  'text/csv; charset=utf-8',
    '.pdf':  'application/pdf',
    '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    '.txt':  'text/plain; charset=utf-8',
    '.png':  'image/png',
    '.jpg':  'image/jpeg',
    '.jpeg': 'image/jpeg',
}

def _shared_files_page_html() -> bytes:
    files = list_shared_files()
    key_q = f'?key={quote(API_KEY)}' if API_KEY else ''
    rows_html = ''.join(
        f'<tr><td><a href="/files/{quote(f["name"])}{key_q}">{f["name"]}</a></td>'
        f'<td>{f["size"]//1024:,} KB</td><td>{f["mtime"]}</td></tr>'
        for f in files
    ) or '<tr><td colspan="3" style="color:#888">업로드된 파일이 없습니다.</td></tr>'
    html = f"""<!DOCTYPE html><html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>주식트래커 - 공유 파일함</title>
<style>
body{{font-family:-apple-system,'맑은 고딕',sans-serif;max-width:720px;margin:24px auto;padding:0 16px;color:#222}}
h2{{margin-bottom:4px}} table{{width:100%;border-collapse:collapse;margin-top:16px}}
td,th{{padding:8px;border-bottom:1px solid #eee;text-align:left;font-size:14px}}
a{{color:#1565C0;text-decoration:none}} a:hover{{text-decoration:underline}}
.up{{margin-top:20px;padding:16px;background:#f5f5f5;border-radius:10px}}
button{{background:#1565C0;color:#fff;border:none;padding:8px 16px;border-radius:6px;cursor:pointer}}
input[type=file]{{margin-right:8px}}
</style></head><body>
<h2>📁 공유 파일함</h2>
<p style="color:#666;font-size:13px">엑셀·PDF 등을 올려두면 노트북·데스크탑·폰 어디서든 이 페이지에서 받을 수 있습니다.</p>
<div class="up">
  <input type="file" id="f" accept=".xlsx,.xls,.csv,.pdf,.docx,.txt,.png,.jpg,.jpeg">
  <button onclick="upload()">업로드</button>
  <span id="msg" style="margin-left:10px;font-size:13px"></span>
</div>
<table><tr><th>파일명</th><th>크기</th><th>수정일시</th></tr>{rows_html}</table>
<script>
async function upload(){{
  const el = document.getElementById('f');
  const file = el.files[0];
  const msg = document.getElementById('msg');
  if(!file){{ msg.textContent='파일을 선택하세요'; return; }}
  msg.textContent='업로드 중...';
  try {{
    const res = await fetch('/api/files/upload?key={quote(API_KEY) if API_KEY else ""}&name='+encodeURIComponent(file.name), {{method:'POST', body: file}});
    const j = await res.json();
    if(j.ok){{ msg.textContent='업로드 완료'; location.reload(); }}
    else {{ msg.textContent='실패: '+(j.error||''); }}
  }} catch(e) {{ msg.textContent='오류: '+e; }}
}}
</script>
</body></html>"""
    return html.encode('utf-8')

# ── 포트폴리오 파일 읽기/쓰기 ────────────────────────────────
_portfolio_lock = threading.Lock()

def load_portfolio() -> dict:
    """portfolio.json 읽기. 없으면 빈 구조 반환."""
    try:
        with _portfolio_lock:
            with open(PORTFOLIO_PATH, 'r', encoding='utf-8') as f:
                return json.load(f)
    except FileNotFoundError:
        return {'stocks': [], 'bp': {}, 'qty': {}, 'dt': {}, 'memo': {}, 'sell': {}, 'watch': []}
    except Exception as e:
        _log(f'portfolio load error: {e}')
        return {'stocks': [], 'bp': {}, 'qty': {}, 'dt': {}, 'memo': {}, 'sell': {}, 'watch': []}

def save_portfolio(data: dict):
    """portfolio.json 저장."""
    try:
        with _portfolio_lock:
            with open(PORTFOLIO_PATH, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        _log(f'portfolio save error: {e}')


# ── 추천주 날짜 (My_Stock_Bundle_Status.xlsx 파싱) ─────────────
# 로컬 PC: 서버 폴더에 놓인 엑셀을 직접 파싱해 서빙.
# 클라우드(엑셀 없음): PC가 파싱해서 push한 recommend_dates.json 을 그대로 서빙.
EXCEL_PATH          = os.path.join(SCRIPT_DIR, 'My_Stock_Bundle_Status.xlsx')
RECOMMEND_JSON_PATH = os.path.join(SCRIPT_DIR, 'recommend_dates.json')
_recommend_lock  = threading.Lock()
_recommend_cache = {'data': None, 'mtime': None}

_REC_SKIP_HEADER_KW = ('종목', '수익률', '횟수')                     # 날짜 열이 아닌 헤더(집계표 등)
_REC_SKIP_NAME_KW   = ('특별판', '반등', '매도', '라이브', '참고', '주의')  # 종목명이 아닌 안내 문구

def _rec_clean_name(raw):
    """셀 값 → 종목명 정제. '004170 신세계' → '신세계'. 종목명이 아니면 None."""
    if raw is None:
        return None
    s = str(raw).strip()
    if not s or len(s) > 14:
        return None
    for kw in _REC_SKIP_NAME_KW:
        if kw in s:
            return None
    parts = s.split(' ', 1)
    if len(parts) == 2 and re.fullmatch(r'\d{4,6}', parts[0]):
        s = parts[1].strip()
    return s or None

def _rec_parse_date(header):
    """'2월 6일' → '2/6'. 날짜 형식이 아니면 None."""
    m = re.match(r'(\d{1,2})\s*월\s*(\d{1,2})\s*일', str(header or ''))
    if not m:
        return None
    return f"{int(m.group(1))}/{int(m.group(2))}"

def parse_recommend_excel() -> dict:
    """My_Stock_Bundle_Status.xlsx (시트별 = 월그룹) → {종목명: {월그룹: [날짜,...]}}"""
    if not OPENPYXL or not os.path.exists(EXCEL_PATH):
        return {}
    result = {}
    wb = None
    try:
        wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True, read_only=True)
        for ws in wb.worksheets:
            month_group = (ws.title or '').strip()
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                continue
            header = rows[0]
            for col_idx, h in enumerate(header):
                if h is None:
                    continue
                hs = str(h)
                if any(kw in hs for kw in _REC_SKIP_HEADER_KW):
                    continue
                date_str = _rec_parse_date(hs)
                if not date_str:
                    continue
                for r in rows[1:]:
                    if col_idx >= len(r):
                        continue
                    name = _rec_clean_name(r[col_idx])
                    if not name:
                        continue
                    grp = result.setdefault(name, {}).setdefault(month_group, [])
                    if date_str not in grp:
                        grp.append(date_str)
        for name in result:
            for grp in result[name]:
                result[name][grp].sort(key=lambda d: tuple(map(int, d.split('/'))))
        _log(f'[추천주] 엑셀 파싱 완료: {len(result)}개 종목')
    except Exception as e:
        _log(f'[추천주] 엑셀 파싱 오류: {type(e).__name__}: {e}')
        return {}
    finally:
        # read_only 워크북은 명시적으로 close()하지 않으면 파일 핸들(zip)이 계속 열려있어
        # 엑셀에서 이 파일을 저장하려 할 때 "공유 위반" 오류가 날 수 있음 → 반드시 닫아줌
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass
    return result

def load_recommend_json() -> dict:
    try:
        with open(RECOMMEND_JSON_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}

def save_recommend_json(data: dict):
    try:
        with open(RECOMMEND_JSON_PATH, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        _log(f'[추천주] recommend_dates.json 저장: {len(data)}개 종목')
    except Exception as e:
        _log(f'[추천주] 저장 오류: {e}')

def get_recommend_dates() -> dict:
    """로컬(엑셀 있음): mtime 기준 캐시 파싱. 클라우드(엑셀 없음): push받은 JSON 서빙."""
    if OPENPYXL and os.path.exists(EXCEL_PATH):
        try:
            mtime = os.path.getmtime(EXCEL_PATH)
        except OSError:
            mtime = None
        with _recommend_lock:
            if _recommend_cache['mtime'] != mtime or _recommend_cache['data'] is None:
                data = parse_recommend_excel()
                if data:
                    _recommend_cache['data']  = data
                    _recommend_cache['mtime'] = mtime
            return _recommend_cache['data'] or {}
    return load_recommend_json()


# ════════════════════════════════════════════════════════
# 텔레그램 일일 리포트 (관심종목 + 보유종목 동향/이슈 요약)
# ════════════════════════════════════════════════════════

def _strip_html(s: str) -> str:
    """HTML 태그 제거 + 엔티티 디코딩 (뉴스 제목 등)"""
    s = re.sub(r'<[^>]+>', '', s or '')
    return html.unescape(s).strip()


def fetch_stock_news(code: str, limit: int = 2) -> list:
    """네이버 종목뉴스 최신 헤드라인 조회 (실패 시 빈 리스트, 절대 예외를 밖으로 던지지 않음)"""
    try:
        url = f"https://finance.naver.com/item/news_news.naver?code={code}&page=1&sm=title_entity_id.basic"
        r = SESSION.get(url, headers={**HEADERS, "Accept": "text/html,*/*"}, timeout=TIMEOUT)
        r.encoding = 'euc-kr'
        body = r.text
        rows = re.findall(
            r'<td\s+class="title">\s*<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>.*?'
            r'<td\s+class="date">\s*([^<]+?)\s*</td>',
            body, re.S)
        out, seen = [], set()
        for href, title, date in rows:
            t = _strip_html(title)
            if not t or t in seen:
                continue
            seen.add(t)
            link = href if href.startswith('http') else f"https://finance.naver.com{href}"
            out.append({'title': t, 'date': date.strip(), 'link': link})
            if len(out) >= limit:
                break
        return out
    except Exception as e:
        _log(f"[뉴스ERR] fetch_stock_news({code}): {type(e).__name__}: {e}")
        return []


_gemini_sem = threading.Semaphore(3)  # 무료 티어 RPM 보호용 동시 호출 제한

def summarize_with_gemini(name: str, code: str, today_pct, pnl_pct, frgn, has_div: bool, news: list) -> str:
    """종목의 시세 요약 + 최근 뉴스 제목들을 Gemini에게 넘겨 1~2문장 해석/요약을 받아온다.
    API 키 미설정이거나 호출 실패 시 빈 문자열 반환 (절대 예외를 밖으로 던지지 않음)."""
    if not GEMINI_API_KEY:
        return ''
    if not news:
        return ''
    try:
        news_lines = "\n".join(f"- {n['title']} ({n['date']})" for n in news)
        div_txt = "배당 있음" if has_div else "배당 없음"
        frgn_txt = f"외국인 지분율 {frgn}%" if frgn is not None else "외국인 지분율 정보 없음"
        prompt = (
            "너는 한국 주식 애널리스트야. 아래 종목의 최근 뉴스 헤드라인과 시세 정보를 근거로, "
            "증권사 리포트 요약처럼 3~4문장 분량의 분석 코멘트를 한국어로 작성해줘.\n"
            "- 무슨 사업을 하는 회사인지(알고 있는 경우 한 문장), 최근 어떤 이슈·공시·실적·수주·증권사 의견이 있었는지, "
            "그것이 주가 흐름(상승/조정)에 어떤 의미인지, 향후 확인할 리스크나 체크포인트를 담아줘.\n"
            "- 뉴스에 없는 구체적 수치(목표가, 실적 등)는 지어내지 말고, 모르면 언급하지 마.\n"
            "- '사세요/파세요' 같은 직접적 매매 권유는 하지 마.\n\n"
            f"종목: {name}({code})\n"
            f"오늘 등락률: {today_pct:+.1f}%\n"
            f"누적 손익률: {pnl_pct if pnl_pct is not None else '정보없음'}\n"
            f"{frgn_txt}, {div_txt}\n"
            f"최근 뉴스:\n{news_lines}\n\n"
            "답변은 분석 문장만, 머리말·따옴표·마크다운 없이 바로 작성해줘."
        )
        with _gemini_sem:
            url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent?key={GEMINI_API_KEY}"
            r = requests.post(url, json={
                "contents": [{"parts": [{"text": prompt}]}],
                "generationConfig": {"temperature": 0.3, "maxOutputTokens": 600, "thinkingConfig": {"thinkingBudget": 0}},
            }, timeout=20)
        if r.status_code != 200:
            _log(f"[Gemini ERR] {code}: HTTP {r.status_code} {r.text[:200]}")
            return ''
        j = r.json()
        cands = j.get('candidates') or []
        if not cands:
            return ''
        parts = (cands[0].get('content') or {}).get('parts') or []
        text = "".join(p.get('text', '') for p in parts).strip()
        return text
    except Exception as e:
        _log(f"[Gemini ERR] {code}: {type(e).__name__}: {e}")
        return ''


def build_daily_report() -> str:
    """관심종목 + 보유종목 전체를 훑어 일일 리포트 텍스트 생성.
    시세/뉴스 조회는 공유 스레드풀로 병렬 처리해 종목이 많아도 전체 소요시간을 줄인다."""
    pf = load_portfolio()
    stocks = pf.get('stocks', []) or []
    watch  = pf.get('watch', []) or []
    bp_map, qty_map = pf.get('bp', {}) or {}, pf.get('qty', {}) or {}

    entries = []  # (code, name_hint, brkr, bp, qty)
    seen = set()
    for s in stocks:
        code = s.get('realCode') or s.get('code')
        entries.append((code, s.get('name', ''), s.get('brkr', ''),
                         bp_map.get(s.get('code')), qty_map.get(s.get('code'))))
        seen.add(code)
    watch_codes = [c for c in watch if c not in seen]

    all_codes = list({e[0] for e in entries} | set(watch_codes))
    ex = _shared_pool
    d_futs = {c: ex.submit(get_stock_data, c) for c in all_codes}
    n_futs = {c: ex.submit(fetch_stock_news, c, 2) for c in all_codes}
    d_map, n_map = {}, {}
    for c, f in d_futs.items():
        try:
            d_map[c] = f.result(timeout=15)
        except Exception as e:
            _log(f"[리포트] get_stock_data({c}) 타임아웃/오류: {e}")
            d_map[c] = {'ok': False, 'code': c}
    for c, f in n_futs.items():
        try:
            n_map[c] = f.result(timeout=10)
        except Exception:
            n_map[c] = []

    def block(code, name_hint='', brkr='', bp=None, qty=None) -> str:
        d = d_map.get(code) or {'ok': False}
        name = d.get('name') or name_hint or code
        if not d.get('ok') or d.get('px') is None:
            return f"■ {name} ({code})\n  ⚠ 시세 조회 실패\n"

        px, prev = d['px'], d.get('prevClose')
        chg_r = ((px - prev) / prev * 100) if prev else None
        chg_txt = f"{chg_r:+.2f}%" if chg_r is not None else "집계중"
        lines = [f"■ {name} ({code}){' · '+brkr if brkr else ''}",
                 f"  현재가 {px:,.0f}원 ({chg_txt})"]

        if bp and qty:
            gain_r = (px - bp) / bp * 100
            lines.append(f"  보유 {qty:,}주 · 매수단가 {bp:,.0f}원 · 평가손익 {gain_r:+.2f}%")

        target = d.get('targetPrice')
        if target:
            gap = (target - px) / px * 100
            lines.append(f"  컨센서스 목표가 {target:,.0f}원 (현재가 대비 {gap:+.1f}%, 증권사 {d.get('institutionCount') or '?'}곳)")

        inv_daily = d.get('invDaily') or []
        if inv_daily:
            last3 = inv_daily[-3:]
            organ_sum = sum((r.get('organ') or 0) for r in last3)
            frgn_sum  = sum((r.get('frgn') or 0) for r in last3)
            organ_lbl = '순매수' if organ_sum > 0 else ('순매도' if organ_sum < 0 else '중립')
            frgn_lbl  = '순매수' if frgn_sum > 0 else ('순매도' if frgn_sum < 0 else '중립')
            lines.append(f"  최근 3일 수급: 기관 {organ_lbl} · 외국인 {frgn_lbl}")

        news = n_map.get(code) or []
        if news:
            for n in news:
                lines.append(f"  📰 {n['title']} ({n['date']})")
        else:
            lines.append("  📰 최근 뉴스 없음")
        return "\n".join(lines) + "\n"

    today = datetime.now().strftime('%Y-%m-%d (%a)')
    parts = [f"📊 주식트래커 일일 리포트\n{today} 장마감 기준\n" + "─"*22]

    if entries:
        parts.append(f"\n【 보유 종목 ({len(entries)}) 】\n")
        for code, name_hint, brkr, bp, qty in entries:
            parts.append(block(code, name_hint, brkr, bp, qty))
    if watch_codes:
        parts.append(f"\n【 관심 종목 ({len(watch_codes)}) 】\n")
        for code in watch_codes:
            parts.append(block(code))
    if not entries and not watch_codes:
        parts.append("\n등록된 보유/관심 종목이 없습니다.\n")

    tips = []
    for code, name_hint, brkr, bp, qty in entries:
        d = d_map.get(code) or {}
        px, target = d.get('px'), d.get('targetPrice')
        if px and target and px >= target:
            tips.append(f"• {name_hint or code}: 목표가 도달 → 익절 검토")
    if tips:
        parts.append("\n【 종합 제안 】\n" + "\n".join(tips) + "\n")

    return "\n".join(parts)


# ════════════════════════════════════════════════════════
# 일일 종합분석 엑셀/PDF 자동 생성 + 이메일 발송
# ════════════════════════════════════════════════════════

_CODE_ALIAS = {}  # 예: '060280A' -> '060280' (실제 코드가 다른 별칭 보유분)

def _score_row(today_pct, frgn, pnl_pct, has_div) -> float:
    s = 0.0
    if today_pct is None: today_pct = 0
    if today_pct >= 3: s += 2
    elif today_pct >= 0: s += 1
    elif today_pct > -3: s -= 1
    else: s -= 2
    if frgn is not None:
        if frgn >= 20: s += 1
        elif frgn >= 5: s += 0.5
        else: s -= 0.5
    if has_div: s += 0.5
    if pnl_pct is not None:
        if pnl_pct <= -40: s -= 1.5
        elif pnl_pct <= -15: s -= 0.5
        elif pnl_pct >= 0: s += 1
    return s

def _bucket_of(score: float) -> str:
    if score >= 2: return "상승 흐름 기대"
    if score >= 0: return "유지/관망"
    if score >= -2: return "리스크 관리 필요"
    return "보수적 접근 권고"

def collect_analysis_rows():
    """보유종목(bp/qty 기준 실제 잔량) + 관심종목을 조회해 종합분석 행 데이터를 만든다."""
    pf = load_portfolio()
    stocks = pf.get('stocks', []) or []
    bp_map = pf.get('bp', {}) or {}
    qty_map = pf.get('qty', {}) or {}
    watch = pf.get('watch', []) or []

    name_map, brkr_map = {}, {}
    for s in stocks:
        name_map[s['code']] = s.get('name', s['code'])
        brkr_map[s['code']] = s.get('brkr', '')
        if s.get('realCode'):
            _CODE_ALIAS[s['code']] = s['realCode']

    held_codes = [c for c in bp_map.keys() if qty_map.get(c)]
    real_codes = {_CODE_ALIAS.get(c, c) for c in held_codes}
    watch_only = [c for c in watch if c not in real_codes and c not in held_codes]

    all_real = list(real_codes | set(watch_only))
    ex = _shared_pool

    # 동시 요청이 너무 많으면 네이버가 차단/지연시켜 전부 실패할 수 있어 소규모 풀로 제한 + 재시도
    def _fetch_with_retry(c):
        last = {'ok': False}
        for attempt in range(3):
            try:
                d = get_stock_data(c)
                if d and d.get('px') is not None:
                    return d
                last = d or last
            except Exception as e:
                _log(f"[종합분석] get_stock_data({c}) 시도{attempt+1} 실패: {type(e).__name__}: {e}")
            time.sleep(1.0 + attempt)
        return last

    data_map = {}
    with ThreadPoolExecutor(max_workers=5) as small_ex:
        sfuts = {c: small_ex.submit(_fetch_with_retry, c) for c in all_real}
        for c, f in sfuts.items():
            try:
                data_map[c] = f.result(timeout=90)
            except Exception as e:
                _log(f"[종합분석] get_stock_data({c}) 최종 실패: {e}")
                data_map[c] = {'ok': False}
    _ok_cnt = sum(1 for d in data_map.values() if d.get('px') is not None)
    _log(f"[종합분석] 시세 조회 {_ok_cnt}/{len(all_real)} 성공 (보유코드 {len(held_codes)}개, 관심 {len(watch_only)}개)")

    news_futs = {c: ex.submit(fetch_stock_news, c, 5) for c in all_real}
    news_map = {}
    for c, f in news_futs.items():
        try:
            news_map[c] = f.result(timeout=10)
        except Exception:
            news_map[c] = []

    rows = []
    for code in held_codes:
        buy = bp_map.get(code)
        qty = qty_map.get(code)
        if not buy or not qty:
            continue
        real = _CODE_ALIAS.get(code, code)
        d = data_map.get(real, {}) or {}
        px = d.get('px')
        if px is None:
            continue
        prev = d.get('prevClose') or px
        pnl_pct = (px - buy) / buy * 100
        today_pct = (px - prev) / prev * 100 if prev else 0
        frgn = d.get('frgnRatio')
        has_div = bool(d.get('dividend'))
        score = _score_row(today_pct, frgn, pnl_pct, has_div)
        rows.append({
            'name': name_map.get(code, real), 'code': real, 'brkr': brkr_map.get(code, ''),
            'buy': buy, 'qty': qty, 'px': px, 'pnl_pct': pnl_pct, 'today_pct': today_pct,
            'frgn': frgn, 'has_div': has_div, 'score': score, 'bucket': _bucket_of(score),
            'news': news_map.get(real, []),
        })
    rows.sort(key=lambda r: -r['score'])

    wrows = []
    for code in watch_only:
        d = data_map.get(code, {}) or {}
        px = d.get('px')
        if px is None:
            continue
        prev = d.get('prevClose') or px
        today_pct = (px - prev) / prev * 100 if prev else 0
        frgn = d.get('frgnRatio')
        has_div = bool(d.get('dividend'))
        score = _score_row(today_pct, frgn, None, has_div)
        wrows.append({
            'name': d.get('name', code), 'code': code, 'px': px, 'today_pct': today_pct,
            'frgn': frgn, 'has_div': has_div, 'score': score, 'bucket': _bucket_of(score),
            'news': news_map.get(code, []),
        })
    wrows.sort(key=lambda r: -r['score'])

    # Gemini AI 해석/요약 (API 키 설정된 경우에만 동작, 뉴스가 있는 종목만 대상)
    if GEMINI_API_KEY:
        all_entries = rows + wrows
        insight_futs = {}
        for r in all_entries:
            if r.get('news'):
                insight_futs[id(r)] = (r, ex.submit(
                    summarize_with_gemini, r['name'], r['code'], r['today_pct'],
                    r.get('pnl_pct'), r['frgn'], r['has_div'], r['news']))
        for key, (r, f) in insight_futs.items():
            try:
                r['insight'] = f.result(timeout=25)
            except Exception as e:
                _log(f"[Gemini ERR] {r['code']} 결과 수신 실패: {e}")
                r['insight'] = ''
        for r in all_entries:
            r.setdefault('insight', '')
    else:
        for r in rows + wrows:
            r['insight'] = ''

    return rows, wrows


def generate_daily_xlsx(rows, wrows, date_str: str) -> str:
    if not OPENPYXL:
        raise RuntimeError('openpyxl 미설치')
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '종합분석'
    head_fill = openpyxl.styles.PatternFill('solid', fgColor='1565C0')
    head_font = openpyxl.styles.Font(color='FFFFFF', bold=True)
    headers = ['구분', '종목명', '코드', '증권사', '매입가', '수량', '현재가', '오늘%', '누적%', '외국인%', '배당', '분석 코멘트', '최근 뉴스']
    ws.append(headers)
    for c in ws[1]:
        c.fill = head_fill; c.font = head_font
    def news_txt(r):
        news = r.get('news') or []
        return ' / '.join(f"{n['title']}({n['date']})" for n in news) if news else ''
    for r in rows:
        ws.append([r['bucket'], r['name'], r['code'], r['brkr'], r['buy'], r['qty'], r['px'],
                   round(r['today_pct'], 2), round(r['pnl_pct'], 2), r['frgn'], 'Y' if r['has_div'] else '',
                   _commentary(r), news_txt(r)])
    ws.append([])
    ws.append(['── 관심종목(미보유) ──'])
    for r in wrows:
        ws.append([r['bucket'], r['name'], r['code'], '', '', '', r['px'],
                   round(r['today_pct'], 2), '', r['frgn'], 'Y' if r['has_div'] else '',
                   _commentary(r), news_txt(r)])
    for i, w in enumerate([14, 16, 10, 10, 10, 8, 10, 8, 8, 9, 6, 45, 40], start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = openpyxl.styles.Alignment(wrap_text=True, vertical='top')
    fpath = os.path.join(SHARED_DIR, f'종합분석_{date_str}.xlsx')
    wb.save(fpath)
    return fpath


def _portfolio_totals(rows):
    buy = sum(r['buy'] * r['qty'] for r in rows)
    ev = sum(r['px'] * r['qty'] for r in rows)
    pl = ev - buy
    return buy, ev, pl, (pl / buy * 100 if buy else 0.0)


def _commentary(r) -> str:
    """종목별 분석 코멘트. Gemini 해석이 있으면 그것을, 없으면 뉴스+수급 기반 규칙형 문장을 반환."""
    ins = (r.get('insight') or '').strip()
    if ins:
        return ins
    parts = []
    news = r.get('news') or []
    if news:
        parts.append("최근 뉴스: " + " / ".join(f"{n['title']}({n['date']})" for n in news[:3]) + ".")
    t = r.get('today_pct') or 0
    if t >= 3:
        parts.append(f"오늘 {t:+.1f}% 강세로 단기 수급이 몰리는 모습.")
    elif t <= -3:
        parts.append(f"오늘 {t:+.1f}% 약세로 단기 매도 압력이 확인됨.")
    f = r.get('frgn')
    if f is not None:
        if f >= 20:
            parts.append(f"외국인 지분율 {f}%로 수급 기반이 탄탄한 편.")
        elif f < 5:
            parts.append(f"외국인 지분율 {f}%로 낮아 외부 수급 뒷받침은 약한 편.")
    p = r.get('pnl_pct')
    if p is not None:
        if p <= -40:
            parts.append(f"누적 {p:+.1f}%로 손실 폭이 커 비중 점검이 필요.")
        elif p >= 20:
            parts.append(f"누적 {p:+.1f}%로 수익 구간, 일부 차익실현 여부 점검 가능.")
    return " ".join(parts) if parts else "특이 뉴스·수급 변화 없음."


def _tag_text(r, held=True) -> str:
    bits = []
    if held and r.get('pnl_pct') is not None:
        bits.append(f"누적 {r['pnl_pct']:+.1f}%")
    bits.append(f"오늘 {r['today_pct']:+.1f}%")
    if r.get('frgn') is not None:
        bits.append(f"외인 {r['frgn']}%")
    if r.get('has_div'):
        bits.append("배당")
    return "[" + ", ".join(bits) + "]"


def _head_text(r, held=True) -> str:
    if held:
        return (f"{r['name']}({r['code']}, {r['brkr']}) - 매입 {r['buy']:,.0f}원 x {r['qty']}주 "
                f"→ 현재 {r['px']:,.0f}원 {_tag_text(r)}")
    return f"{r['name']}({r['code']}) - 현재 {r['px']:,.0f}원 {_tag_text(r, held=False)}"


_BUCKET_ORDER = ['상승 흐름 기대', '유지/관망', '리스크 관리 필요', '보수적 접근 권고']
_BUCKET_RGB = {'상승 흐름 기대': (0.1, 0.55, 0.25), '유지/관망': (0.75, 0.6, 0.0),
               '리스크 관리 필요': (0.85, 0.45, 0.0), '보수적 접근 권고': (0.8, 0.15, 0.15)}
_NOTICE = ("오늘 모멘텀, 외국인 지분율, 배당, 누적손익 점수로 1차 분류하고, 종목별 최근 뉴스·공시를 함께 해석해 "
           "근거를 적었습니다. 투자 자문이 아닌 참고 자료이며, 실제 매매 판단은 본인 책임 하에 신중히 내리시기 바랍니다.")


def _pdf_clean(text: str) -> str:
    """HYSMyeongJo CID 폰트가 못 그리는 문자(가운뎃점, 줄임표, 긴 대시, 이모지 등)를 안전한 문자로 치환."""
    for a, b in (('·', ','), ('ㆍ', ','), ('•', '-'), ('…', '...'), ('—', '-'), ('–', '-'),
                 ('‘', "'"), ('’', "'"), ('“', '"'), ('”', '"'), (' ', ' ')):
        text = text.replace(a, b)
    return ''.join(ch for ch in text if ord(ch) < 0x1F000 and not (0x2600 <= ord(ch) <= 0x27BF))


def _wrap_pdf(text, font, size, maxw):
    text = _pdf_clean(text)
    lines, cur = [], ''
    for ch in text:
        if ch == '\n':
            lines.append(cur); cur = ''
            continue
        if _rl_pdfmetrics.stringWidth(cur + ch, font, size) > maxw:
            lines.append(cur); cur = ch.lstrip()
        else:
            cur += ch
    if cur:
        lines.append(cur)
    return lines


def generate_daily_pdf(rows, wrows, date_str: str) -> str:
    """상세 분석 PDF: 요약 + 등급별 종목 + 종목별 분석 코멘트 + 번외(관심종목)."""
    if not REPORTLAB:
        raise RuntimeError('reportlab 미설치')
    fpath = os.path.join(SHARED_DIR, f'종합분석_{date_str}.pdf')
    c = _rl_canvas.Canvas(fpath, pagesize=_RL_A4)
    width, height = _RL_A4
    FONT = 'HYSMyeongJo-Medium'
    LM, RM = 40, 40
    maxw = width - LM - RM
    y = height - 50

    def text_block(text, size=9, dy=13, color=(0, 0, 0), indent=0):
        nonlocal y
        for ln in _wrap_pdf(text, FONT, size, maxw - indent):
            if y < 50:
                c.showPage(); y = height - 50
            c.setFont(FONT, size)
            c.setFillColorRGB(*color)
            c.drawString(LM + indent, y, ln)
            y -= dy

    text_block("보유 · 관심 종목 종합분석", size=17, dy=24)
    text_block(f"기준일: {date_str} (뉴스·공시 해석 반영)", size=10, dy=16, color=(0.3, 0.3, 0.3))
    text_block("※ " + _NOTICE, size=8, dy=12, color=(0.4, 0.4, 0.4))
    y -= 8

    if rows:
        b, e, p, pp = _portfolio_totals(rows)
        text_block("포트폴리오 요약", size=12, dy=18, color=(0.08, 0.4, 0.75))
        text_block(f"- 총 매입 {b:,.0f}원 → 총 평가 {e:,.0f}원 (총손익 {p:+,.0f}원, {pp:+.1f}%)", size=10, dy=15)
        y -= 6

    for bk in _BUCKET_ORDER:
        grp = [r for r in rows if r['bucket'] == bk]
        if not grp:
            continue
        y -= 6
        text_block(f"● {bk} ({len(grp)}종목)", size=12, dy=18, color=_BUCKET_RGB.get(bk, (0, 0, 0)))
        for r in grp:
            text_block("- " + _head_text(r), size=9.5, dy=14)
            text_block(_commentary(r), size=9, dy=13, color=(0.25, 0.25, 0.25), indent=14)
            y -= 5

    if wrows:
        y -= 8
        text_block("[번외] 관심종목 (미보유)", size=12, dy=18, color=(0, 0.5, 0.55))
        text_block("매입 이력이 없어 손익은 제외하고, 모멘텀·수급·뉴스만으로 분류했습니다.", size=8.5, dy=13, color=(0.4, 0.4, 0.4))
        for r in wrows:
            text_block("- " + _head_text(r, held=False) + f" / {r['bucket']}", size=9.5, dy=14)
            text_block(_commentary(r), size=9, dy=13, color=(0.25, 0.25, 0.25), indent=14)
            y -= 5

    y -= 10
    text_block("참고", size=11, dy=16, color=(0.08, 0.4, 0.75))
    text_block("- 공시·뉴스 내용은 조사 시점 기준이며, 이후 추가 공시나 시황 변화로 상황이 달라질 수 있습니다.", size=8.5, dy=12)
    text_block("- 이 자료는 투자 자문이 아닌 참고 자료이며, 실제 매수·매도 결정은 본인 판단과 책임 하에 하시기 바랍니다.", size=8.5, dy=12)
    c.save()
    return fpath


def generate_daily_docx(rows, wrows, date_str: str) -> str:
    """상세 분석 Word 문서 (샘플 '보유종목_관심종목_종합분석_상세.docx' 구성과 동일)."""
    if not DOCX_OK:
        raise RuntimeError('python-docx 미설치')
    from docx.shared import Pt, RGBColor
    from docx.oxml.ns import qn
    doc = _docx.Document()
    st = doc.styles['Normal']
    st.font.name = '맑은 고딕'
    st.font.size = Pt(10)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')

    def para(text, size=10, bold=False, color=None, indent=0):
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.font.size = Pt(size)
        run.bold = bold
        run.font.name = '맑은 고딕'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')
        if color:
            run.font.color.rgb = RGBColor(*color)
        if indent:
            p.paragraph_format.left_indent = Pt(indent)
        p.paragraph_format.space_after = Pt(3)
        return p

    para("보유 · 관심 종목 종합분석 (뉴스·공시 해석 반영판)", size=18, bold=True)
    para(f"기준일: {date_str}", size=10, color=(90, 90, 90))
    para("※ " + _NOTICE, size=9, color=(110, 110, 110))

    if rows:
        b, e, p, pp = _portfolio_totals(rows)
        para("📊 포트폴리오 요약", size=13, bold=True, color=(21, 101, 192))
        para(f"총 매입 {b:,.0f}원 → 총 평가 {e:,.0f}원 (총손익 {p:+,.0f}원, {pp:+.1f}%)", size=10.5)

    emoji = {'상승 흐름 기대': '🟢', '유지/관망': '🟡', '리스크 관리 필요': '🟠', '보수적 접근 권고': '🔴'}
    for bk in _BUCKET_ORDER:
        grp = [r for r in rows if r['bucket'] == bk]
        if not grp:
            continue
        para(f"{emoji.get(bk, '')} {bk} ({len(grp)}종목)", size=13, bold=True)
        for r in grp:
            para(_head_text(r), size=10, bold=True)
            para(_commentary(r), size=9.5, color=(60, 60, 60), indent=14)

    if wrows:
        para("⭐ 번외 종목 (관심종목 — 미보유)", size=13, bold=True, color=(0, 128, 140))
        para("매입 이력이 없어 손익은 제외하고, 모멘텀·수급·뉴스만으로 분류했습니다.", size=9, color=(110, 110, 110))
        for r in wrows:
            para(_head_text(r, held=False) + f" / {r['bucket']}", size=10, bold=True)
            para(_commentary(r), size=9.5, color=(60, 60, 60), indent=14)

    para("📌 참고", size=12, bold=True, color=(21, 101, 192))
    para("공시·뉴스 내용은 조사 시점 기준이며, 이후 추가 공시나 시황 변화로 상황이 달라질 수 있습니다.", size=9)
    para("이 자료는 투자 자문이 아닌 참고 자료이며, 실제 매수·매도 결정은 본인 판단과 책임 하에 하시기 바랍니다.", size=9)
    fpath = os.path.join(SHARED_DIR, f'종합분석_{date_str}.docx')
    doc.save(fpath)
    return fpath


def send_email_with_attachments(subject: str, body: str, attachment_paths: list) -> bool:
    if not (EMAIL_USER and EMAIL_APP_PASS and EMAIL_TO):
        _log("[이메일ERR] EMAIL_USER/EMAIL_APP_PASSWORD/EMAIL_TO 미설정")
        return False
    try:
        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = EMAIL_USER
        msg['To'] = EMAIL_TO
        msg.set_content(body)
        for fp in attachment_paths:
            if not fp or not os.path.exists(fp):
                continue
            with open(fp, 'rb') as f:
                data = f.read()
            ext = os.path.splitext(fp)[1].lower()
            maintype, subtype = ('application', 'octet-stream')
            if ext == '.pdf': maintype, subtype = 'application', 'pdf'
            elif ext == '.xlsx': maintype, subtype = 'application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=os.path.basename(fp))
        with smtplib.SMTP('smtp.gmail.com', 587, timeout=20) as s:
            s.starttls()
            s.login(EMAIL_USER, EMAIL_APP_PASS)
            s.send_message(msg)
        _log(f"[이메일] 발송 완료 → {EMAIL_TO} ({len(attachment_paths)}개 첨부)")
        return True
    except Exception as e:
        _log(f"[이메일ERR] 발송 실패: {type(e).__name__}: {e}")
        return False


_sched_lock = threading.Lock()
_sched_last_fired = {'pre': None, 'post': None}  # 마지막으로 발송 성공한 날짜(YYYY-MM-DD, KST)

def _kst_now() -> datetime:
    return datetime.now(timezone.utc) + timedelta(hours=9)

def _fire_scheduled_job(label: str):
    try:
        _log(f"[스케줄러] {label} 자동 리포트 실행 시작")
        rows, wrows = collect_analysis_rows()
        result = _run_daily_files_job(rows, wrows)
        text = build_telegram_analysis_text(rows, wrows)
        send_telegram_message(f"[{label}]\n" + text)
        _log(f"[스케줄러] {label} 완료: {result}")
    except Exception as e:
        _log(f"[스케줄러ERR] {label} 실패: {type(e).__name__}: {e}")

def _daily_schedule_worker():
    """평일 장시작 전(08:50)·장마감 후(15:40) KST 기준으로 하루 두 번 엑셀/PDF 생성+이메일+텔레그램 발송."""
    _log("[스케줄러] 장전/장마감 자동 리포트 스케줄러 시작 (평일 08:50, 15:40 KST)")
    while True:
        try:
            kst = _kst_now()
            if kst.weekday() < 5:  # 월~금만
                date_str = kst.strftime('%Y-%m-%d')
                hm = kst.hour * 60 + kst.minute
                with _sched_lock:
                    if 530 <= hm <= 535 and _sched_last_fired['pre'] != date_str:
                        _sched_last_fired['pre'] = date_str
                        threading.Thread(target=_fire_scheduled_job, args=('장시작전',), daemon=True).start()
                    if 940 <= hm <= 945 and _sched_last_fired['post'] != date_str:
                        _sched_last_fired['post'] = date_str
                        threading.Thread(target=_fire_scheduled_job, args=('장마감후',), daemon=True).start()
        except Exception as e:
            _log(f"[스케줄러ERR] 루프 오류: {type(e).__name__}: {e}")
        time.sleep(60)


def build_telegram_analysis_text(rows: list, wrows: list) -> str:
    """종합분석 상세 내용을 텔레그램 텍스트로 변환 (요약 + 등급별 종목 + 종목별 분석 코멘트)."""
    today = datetime.now().strftime('%Y-%m-%d (%a)')
    parts = [f"📊 보유·관심종목 종합분석\n기준일 {today}\n" + "─" * 20]
    if rows:
        b, e, p, pp = _portfolio_totals(rows)
        parts.append(f"\n총 매입 {b:,.0f}원 → 총 평가 {e:,.0f}원\n총손익 {p:+,.0f}원 ({pp:+.1f}%)")

    emoji = {'상승 흐름 기대': '🟢', '유지/관망': '🟡', '리스크 관리 필요': '🟠', '보수적 접근 권고': '🔴'}
    for bk in _BUCKET_ORDER:
        grp = [r for r in rows if r['bucket'] == bk]
        if not grp:
            continue
        parts.append(f"\n{emoji.get(bk, '•')} {bk} ({len(grp)}종목)")
        for r in grp:
            parts.append(f"• {_head_text(r)}")
            parts.append(f"   {_commentary(r)}")

    if wrows:
        parts.append("\n⭐ 번외 종목 (관심종목, 미보유)")
        for r in wrows:
            parts.append(f"• {_head_text(r, held=False)} / {r['bucket']}")
            parts.append(f"   {_commentary(r)}")

    parts.append("\n※ 투자 자문이 아닌 참고 자료이며, 실제 매매 판단은 본인 책임 하에 신중히 내리시기 바랍니다.")
    return "\n".join(parts)


def _run_daily_files_job(rows=None, wrows=None) -> dict:
    """종합분석 엑셀/PDF를 날짜별로 생성하고 이메일로 발송. /api/send_report, /api/send_daily_files 공용."""
    date_str = datetime.now().strftime('%Y-%m-%d')
    if rows is None or wrows is None:
        rows, wrows = collect_analysis_rows()
    if not rows and not wrows:
        _pf = load_portfolio()
        return {'ok': False, 'error': '보유/관심 종목 데이터 없음',
                'detail': f"portfolio bp={len(_pf.get('bp', {}) or {})}개, qty={len(_pf.get('qty', {}) or {})}개, watch={len(_pf.get('watch', []) or [])}개 (0이면 로컬앱 미동기화, 아니면 시세조회 실패)"}

    out = {'ok': False, 'xlsx': None, 'pdf': None, 'emailed': False}
    paths = []
    try:
        xp = generate_daily_xlsx(rows, wrows, date_str)
        out['xlsx'] = os.path.basename(xp)
        paths.append(xp)
    except Exception as e:
        _log(f"[종합분석ERR] xlsx 생성 실패: {e}")
    try:
        pp = generate_daily_pdf(rows, wrows, date_str)
        out['pdf'] = os.path.basename(pp)
        paths.append(pp)
    except Exception as e:
        _log(f"[종합분석ERR] pdf 생성 실패: {e}")
    try:
        dp = generate_daily_docx(rows, wrows, date_str)
        out['docx'] = os.path.basename(dp)
        paths.append(dp)
    except Exception as e:
        _log(f"[종합분석ERR] docx 생성 실패: {e}")

    if paths:
        subject = f"[주식트래커] {date_str} 보유·관심종목 종합분석"
        body = (f"{date_str} 기준 보유·관심종목 종합분석 파일을 첨부합니다.\n"
                f"보유 {len(rows)}건 · 관심 {len(wrows)}건\n\n"
                f"※ 투자 자문이 아닌 참고 자료입니다.")
        out['emailed'] = send_email_with_attachments(subject, body, paths)
        out['ok'] = True
    return out


def send_telegram_message(text: str) -> bool:
    """텔레그램으로 메시지 발송 (4096자 제한 → 줄 단위로 안전하게 분할 전송)"""
    if not TELEGRAM_TOKEN or not TELEGRAM_CHAT_ID:
        _log("[텔레그램ERR] TELEGRAM_TOKEN/CHAT_ID 미설정")
        return False
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    chunks, cur = [], ''
    for line in text.split('\n'):
        if len(cur) + len(line) + 1 > 3500:
            chunks.append(cur)
            cur = ''
        cur += line + '\n'
    if cur:
        chunks.append(cur)

    ok_all = True
    for chunk in chunks:
        try:
            r = requests.post(url, json={'chat_id': TELEGRAM_CHAT_ID, 'text': chunk}, timeout=20)
            if not r.ok:
                _log(f"[텔레그램ERR] sendMessage 실패: {r.status_code} {r.text[:200]}")
                ok_all = False
        except Exception as e:
            _log(f"[텔레그램ERR] sendMessage 예외: {type(e).__name__}: {e}")
            ok_all = False
    return ok_all


class Handler(BaseHTTPRequestHandler):

    def _check_auth(self) -> bool:
        if not API_KEY:
            return True
        provided = self.headers.get('X-API-Key', '')
        if not provided:
            m = re.search(r'[?&]key=([^&]+)', self.path)
            if m:
                provided = m.group(1)
        return provided == API_KEY

    def do_GET(self):
        path = self.path.split('?')[0].rstrip('/')

        public_paths = ('', '/index.html', '/stock_tracker.html',
                        '/manifest.json', '/sw.js', '/api/ping', '/api/heartbeat', '/files')

        if path not in public_paths and not self._check_auth():
            self.send_error(401, 'Unauthorized - invalid or missing API key')
            return

        if path in ('', '/index.html', '/stock_tracker.html'):
            self._serve_file(HTML_PATH, 'text/html; charset=utf-8', no_store=True)

        elif path == '/manifest.json':
            self._serve_file(MANIFEST_PATH, 'application/manifest+json; charset=utf-8')

        elif path == '/sw.js':
            self._serve_file(SW_PATH, 'application/javascript; charset=utf-8', no_store=True)

        elif path in ('/icon-192.png', '/icon-512.png'):
            icon_path = os.path.join(SCRIPT_DIR, path.lstrip('/'))
            self._serve_file(icon_path, 'image/png')

        elif path.startswith('/api/stock/'):
            code = path.split('/')[-1]
            if re.fullmatch(r'[A-Za-z0-9]{4,8}', code):
                self._serve_json(get_stock_data(code))
            else:
                self.send_error(400, 'Invalid code')

        elif path == '/api/ping':
            self._serve_json({
                'ok':    True,
                'pykrx': PYKRX,
                'time':  datetime.now().strftime('%H:%M:%S'),
                'market': is_market_open(),
            })

        elif path == '/api/heartbeat':
            # PC 브라우저 탭이 열려있는 동안 주기적으로 호출 → 마지막 신호 시각 갱신.
            # (로컬 모드에서만 의미 있음. 클라우드에선 자동 종료 감시 스레드 자체가 안 돎)
            global _last_heartbeat
            _last_heartbeat = time.time()
            self._serve_json({'ok': True})

        elif path == '/api/portfolio':
            self._serve_json(load_portfolio())

        elif path == '/api/send_report':
            # 종합분석(상승 흐름 기대/유지관망/리스크관리/보수적 접근) 내용을 텔레그램 + 이메일(엑셀/PDF)로 발송
            # (예약 작업이 매일 호출. 예전의 개별 종목 뉴스 나열 리포트는 더 이상 사용하지 않음)
            try:
                rows, wrows = collect_analysis_rows()
                text = build_telegram_analysis_text(rows, wrows)
                sent = send_telegram_message(text)
                result = {'ok': sent, 'chars': len(text)}
            except Exception as e:
                _log(f"[리포트ERR] /api/send_report: {type(e).__name__}: {e}")
                result = {'ok': False, 'error': str(e)}
                rows = wrows = None
            try:
                result['files'] = _run_daily_files_job(rows, wrows)
            except Exception as e:
                _log(f"[종합분석ERR] send_report 연계 실패: {type(e).__name__}: {e}")
                result['files'] = {'ok': False, 'error': str(e)}
            self._serve_json(result)

        elif path == '/api/send_daily_files':
            # 종합분석 엑셀/PDF만 즉시 생성+이메일 발송 (수동 테스트용)
            try:
                self._serve_json(_run_daily_files_job())
            except Exception as e:
                _log(f"[종합분석ERR] /api/send_daily_files: {type(e).__name__}: {e}")
                self._serve_json({'ok': False, 'error': str(e)})

        elif path.startswith('/api/search/'):
            query = unquote(path.split('/api/search/', 1)[-1].strip('/'))
            if query:
                self._serve_json(search_stock_name(query))
            else:
                self._serve_json([])

        elif path == '/api/recommend_dates':
            self._serve_json(get_recommend_dates())

        elif path == '/api/market_indices':
            self._serve_json(fetch_market_indices())

        elif path == '/api/reload':
            clear_cache()
            self._serve_json({'ok': True, 'msg': '캐시 초기화 완료'})

        elif path == '/files':
            self._send(200, _shared_files_page_html(), 'text/html; charset=utf-8', no_store=True)

        elif path.startswith('/files/'):
            try:
                fname = _safe_shared_name(path[len('/files/'):])
            except ValueError as e:
                self.send_error(400, str(e))
                return
            fpath = os.path.join(SHARED_DIR, fname)
            ext = os.path.splitext(fname)[1].lower()
            ctype = _SHARED_MIME.get(ext, 'application/octet-stream')
            try:
                with open(fpath, 'rb') as f:
                    body = f.read()
                self.send_response(200)
                self.send_header('Content-Type', ctype)
                self.send_header('Content-Length', str(len(body)))
                self.send_header('Content-Disposition', f'attachment; filename="{quote(fname)}"')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(body)
            except FileNotFoundError:
                self.send_error(404, 'File not found')
            except (ConnectionAbortedError, ConnectionResetError, BrokenPipeError):
                pass

        elif path == '/api/files/list':
            self._serve_json({'files': list_shared_files()})

        elif path.startswith('/api/debug/'):
            code = path.split('/')[-1]
            if re.fullmatch(r'[A-Za-z0-9]{4,8}', code):
                clear_cache(code)
                data = get_stock_data(code)
                logs = _inv_debug.get(code, ['아직 조회 안 됨'])
                self._serve_json({
                    'code': code,
                    'ok': data.get('ok'),
                    'px': data.get('px'),
                    'prevClose': data.get('prevClose'),
                    'src': data.get('src'),
                    'isMarketOpen': data.get('isMarketOpen'),
                    'frgnRatio': data.get('frgnRatio'),
                    'investors': data.get('investors', {}),
                    'debug_log': logs,
                    'pykrx': PYKRX,
                    'pykrx_inv_dead': PYKRX_INV_DEAD,
                })
            else:
                self.send_error(400, 'Invalid code')
        else:
            self.send_error(404)

    def _serve_file(self, fpath, ctype, no_store=False):
        try:
            with open(fpath, 'rb') as f:
                body = f.read()
            self._send(200, body, ctype, no_store=no_store)
        except FileNotFoundError:
            self.send_error(404, f'File not found: {os.path.basename(fpath)}')
        except Exception:
            self.send_error(500, 'Internal server error')

    def _serve_json(self, obj):
        body = json.dumps(obj, ensure_ascii=False).encode('utf-8')
        self._send(200, body, 'application/json; charset=utf-8')

    def _send(self, code, body, ctype, no_store=False):
        try:
            self.send_response(code)
            self.send_header('Content-Type', ctype)
            self.send_header('Content-Length', str(len(body)))
            self.send_header('Access-Control-Allow-Origin', '*')
            self.send_header('Access-Control-Allow-Headers', 'X-API-Key, Content-Type')
            # HTML·SW는 절대 캐시 금지 (서비스 워커 구버전 방지)
            cc = 'no-store, no-cache, must-revalidate' if no_store else 'no-cache'
            self.send_header('Cache-Control', cc)
            self.end_headers()
            self.wfile.write(body)
        except (ConnectionAbortedError, ConnectionResetError, BrokenPipeError):
            pass  # 클라이언트가 먼저 끊은 경우 — 무시

    def do_POST(self):
        path = self.path.split('?')[0].rstrip('/')
        if not self._check_auth():
            self.send_error(401, 'Unauthorized')
            return

        if path == '/api/portfolio':
            try:
                length = int(self.headers.get('Content-Length', 0))
                body   = self.rfile.read(length)
                data   = json.loads(body.decode('utf-8'))
                # 필수 키 검증
                for key in ('stocks', 'bp', 'qty', 'dt'):
                    if key not in data:
                        data[key] = []  if key == 'stocks' else {}
                if 'memo' not in data:
                    data['memo'] = {}
                if 'sell' not in data:
                    data['sell'] = {}
                if 'watch' not in data:
                    data['watch'] = []
                save_portfolio(data)
                self._serve_json({'ok': True, 'saved': True})
            except Exception as e:
                _log(f'portfolio POST error: {e}')
                self.send_error(400, f'Bad request: {e}')

        elif path == '/api/files/upload':
            # 공유 파일함 업로드: ?name=파일명.xlsx 로 원본 파일명을 넘기고, 바디는 파일 바이트 그대로
            try:
                m = re.search(r'[?&]name=([^&]+)', self.path)
                raw_name = unquote(m.group(1)) if m else ''
                fname = _safe_shared_name(raw_name)
                length = int(self.headers.get('Content-Length', 0))
                if length <= 0 or length > 50 * 1024 * 1024:  # 50MB 제한
                    self.send_error(400, 'Invalid file size')
                    return
                body = self.rfile.read(length)
                with open(os.path.join(SHARED_DIR, fname), 'wb') as f:
                    f.write(body)
                _log(f'[공유파일] 업로드: {fname} ({length:,} bytes)')
                self._serve_json({'ok': True, 'name': fname, 'size': length})
            except ValueError as e:
                self._serve_json({'ok': False, 'error': str(e)})
            except Exception as e:
                _log(f'files upload error: {e}')
                self._serve_json({'ok': False, 'error': str(e)})

        elif path == '/api/recommend_dates':
            # 로컬 PC가 엑셀을 파싱한 결과를 클라우드에 push할 때 사용 (클라우드엔 엑셀이 없음)
            try:
                length = int(self.headers.get('Content-Length', 0))
                body   = self.rfile.read(length)
                data   = json.loads(body.decode('utf-8'))
                if not isinstance(data, dict):
                    self.send_error(400, 'Bad request: expected object')
                    return
                save_recommend_json(data)
                self._serve_json({'ok': True, 'saved': True, 'count': len(data)})
            except Exception as e:
                _log(f'recommend_dates POST error: {e}')
                self.send_error(400, f'Bad request: {e}')
        else:
            self.send_error(404)

    def do_OPTIONS(self):
        """CORS preflight (APK WebView 에서 크로스오리진 POST 지원)"""
        self.send_response(204)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'X-API-Key, Content-Type')
        self.send_header('Access-Control-Max-Age', '86400')
        self.end_headers()

    def log_message(self, fmt, *args):
        if len(args) > 1 and args[1] not in ('200', '304'):
            ts = datetime.now().strftime('%H:%M:%S')
            print(f"  [{ts}] {fmt % args}")


class ThreadedServer(ThreadingMixIn, HTTPServer):
    """동시 요청을 스레드로 처리"""
    daemon_threads = True


# ════════════════════════════════════════════════════════
# 메인
# ════════════════════════════════════════════════════════

def get_local_ip() -> str:
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return '알 수 없음'


def main():
    local_ip = get_local_ip()
    is_cloud = bool(os.environ.get('PORT'))

    print()
    print("=" * 55)
    print("   주식 수익률 트래커  ---  서버 시작 v1.1")
    print("=" * 55)
    if is_cloud:
        print(f"   환경:         클라우드 (PORT={PORT})")
        print(f"   API 인증:     {'활성' if API_KEY else '비활성 (API_KEY 미설정)'}")
    else:
        print(f"   PC 주소:      http://localhost:{PORT}")
        print(f"   스마트폰:     http://{local_ip}:{PORT}  (같은 Wi-Fi)")
        print(f"   API 인증:     {'활성' if API_KEY else '비활성 (로컬 개발 모드)'}")
    print(f"   데이터 소스:  {'pykrx + 네이버' if PYKRX else '네이버 fchart'}")
    print(f"   캐시 TTL:     장중 {CACHE_TTL}초 / 장외 {CACHE_TTL_OFF}초")
    print(f"   장중 여부:    {'장중' if is_market_open() else '장외'}")
    if not is_cloud:
        print(f"   자동 종료:    브라우저 탭을 닫고 {HEARTBEAT_GRACE//60}분 지나면 서버 자동 종료")
    print()
    print("   종료: Ctrl+C")
    print("=" * 55)
    print()

    if not is_cloud:
        global _last_heartbeat
        _last_heartbeat = time.time()   # 서버 막 시작한 시점부터 grace 시작
        threading.Thread(target=_heartbeat_watchdog, daemon=True).start()

    # 장전/장마감 자동 리포트는 항상 켜져있는 클라우드 서버에서만 실행
    # (로컬 PC는 꺼져있는 시간이 많아 중복 발송 방지 차원에서 제외)
    if is_cloud:
        threading.Thread(target=_daily_schedule_worker, daemon=True).start()

    srv = ThreadedServer(('', PORT), Handler)

    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print('\n  서버 종료됨')


if __name__ == '__main__':
    main()
