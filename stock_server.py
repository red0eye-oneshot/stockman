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
HEARTBEAT_GRACE = 90   # 이 시간(초) 동안 heartbeat가 없으면 종료
                       # (브라우저가 백그라운드 탭의 setInterval을 느리게 만드는 경우까지 감안한 여유)

def _heartbeat_watchdog():
    while True:
        time.sleep(10)
        idle = time.time() - _last_heartbeat
        if idle > HEARTBEAT_GRACE:
            print(f"\n  브라우저 탭이 닫힌 것으로 감지됨 ({idle:.0f}초간 응답 없음) → 서버 종료")
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
    # B) (제거) 네이버 모바일 JSON API
    #    /api/stock/{code}/investor 등은 2026년 모두 404 반환 → 제거.
    # ════════════════════════════════

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

    return {'frgnRatio': frgn_ratio, 'investors': normalized, 'invDaily': inv_daily, 'indivDaily': indiv_daily, 'name': stock_name}



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
        return float(str(s).replace(',', '').replace(' ', '').replace('%', ''))
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

def fetch_exchange_rate() -> dict:
    """달러/원 환율 (모바일 JSON API 우선 → 실패 시 HTML 페이지 백업)"""
    out = {'usdkrw': {'value': None, 'change': None, 'changePct': None, 'ok': False}}

    val, chg, pct = _fetch_mstock_json("https://m.stock.naver.com/api/marketindex/exchange/FX_USDKRW", 'fx')
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
    except Exception as e:
        _log(f"  [지수ERR] FX: {type(e).__name__}: {e}")
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
               'mstock': ['.IXIC', 'IXIC']},
    'sp500':  {'symbols': ['SPI@SPX'],                         'discover': 'S&amp;P500',
               'mstock': ['.INX', '.SPX', 'SPX']},
    'vix':    {'symbols': ['CBO@VIX', 'CBOE@VIX'],             'discover': 'VIX',
               'mstock': ['.VIX', 'VIX'],                      'stooq': '^vix',
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

    try:
        for fut in as_completed(futs, timeout=15):
            try:
                result.update(fut.result(timeout=1) or {})
            except Exception as e:
                _log(f"  [지수ERR] {futs.get(fut)}: {type(e).__name__}: {e}")
    except FutureTimeoutError:
        _log("  [지수ERR] market_indices 전체 15초 타임아웃 - 완료된 것만 반영")

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
            inv_data = {'frgnRatio': None, 'investors': {}, 'invDaily': [], 'indivDaily': [], 'name': ''}
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
                        '/manifest.json', '/sw.js', '/api/ping', '/api/heartbeat')

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
            # 관심종목+보유종목 일일 리포트를 생성해 텔레그램으로 발송 (예약 작업이 매일 호출)
            try:
                report = build_daily_report()
                sent = send_telegram_message(report)
                self._serve_json({'ok': sent, 'chars': len(report)})
            except Exception as e:
                _log(f"[리포트ERR] /api/send_report: {type(e).__name__}: {e}")
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
        print(f"   자동 종료:    브라우저 탭을 닫고 {HEARTBEAT_GRACE}초 지나면 서버 자동 종료")
    print()
    print("   종료: Ctrl+C")
    print("=" * 55)
    print()

    if not is_cloud:
        global _last_heartbeat
        _last_heartbeat = time.time()   # 서버 막 시작한 시점부터 grace 시작
        threading.Thread(target=_heartbeat_watchdog, daemon=True).start()

    srv = ThreadedServer(('', PORT), Handler)

    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print('\n  서버 종료됨')


if __name__ == '__main__':
    main()
