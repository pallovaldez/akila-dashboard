"""
╔══════════════════════════════════════════════════════════════╗
║         PORTFOLIO TRACKER — DASHBOARD WEB REAL-TIME         ║
║         Proyecto Omega  |  Miguel Camara                    ║
║                                                             ║
║  USO: Doble click en "Abrir Dashboard.command"              ║
║       Luego abre Chrome en: http://localhost:5555           ║
╚══════════════════════════════════════════════════════════════╝
"""

import json
import time
import threading
import os
import sys
from datetime import datetime, timezone
from http.server import HTTPServer, BaseHTTPRequestHandler

# ── Dependencias ─────────────────────────────────────────────
try:
    import yfinance as yf
    import openpyxl
except ImportError:
    os.system(f"{sys.executable} -m pip install yfinance openpyxl -q")
    try:
        import yfinance as yf
        import openpyxl
    except ImportError:
        os.system(f"{sys.executable} -m pip install --user yfinance openpyxl -q")
        import yfinance as yf
        import openpyxl

# ── Session anti rate-limit: curl_cffi emula fingerprint de Chrome ────────────
# yf.download() usa curl_cffi automáticamente si está instalado (yfinance >= 0.2.38).
# Sólo usamos la session manualmente en yf.Ticker() para info calls individuales.
yf.set_tz_cache_location(f"/tmp/yf_tz_cache_{os.getpid()}")
try:
    from curl_cffi import requests as _curl_requests
    _YF_SESSION = _curl_requests.Session(impersonate="chrome")
    print("  ✅ curl_cffi disponible — yfinance lo usa automáticamente en batch downloads", flush=True)
except ImportError:
    import requests as _requests
    _YF_SESSION = _requests.Session()
    _YF_SESSION.headers["User-Agent"] = (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"
    )
    print("  ⚠️  curl_cffi no disponible — usando requests estándar", flush=True)

# ── Configuración ─────────────────────────────────────────────
PORT              = int(os.environ.get("PORT", 5555))
# En Render los datos van en /data (disco persistente); local usa el directorio del script
_SCRIPT_DIR       = os.path.dirname(os.path.abspath(__file__))
BASE_DIR          = "/data" if os.path.isdir("/data") else _SCRIPT_DIR
EXCEL_PATH        = os.path.join(BASE_DIR, "Portfolio_Tracker_Omega.xlsx")
PORTFOLIOS_DIR    = os.path.join(BASE_DIR, "portfolios")
REFRESH_SEC       = int(os.environ.get("REFRESH_SEC", 15))   # segundos entre actualizaciones

# Portafolios disponibles
PORTFOLIO_NAMES   = ["miguel", "paulo"]

def get_portfolio_dir(portfolio="miguel"):
    """Devuelve el directorio de un portafolio, creándolo si no existe."""
    d = os.path.join(PORTFOLIOS_DIR, portfolio)
    os.makedirs(d, exist_ok=True)
    return d

def get_transactions_path(portfolio="miguel"):
    return os.path.join(get_portfolio_dir(portfolio), "transactions.json")

def get_analysis_path(portfolio="miguel"):
    return os.path.join(get_portfolio_dir(portfolio), "analysis.json")

def get_cash_path(portfolio="miguel"):
    return os.path.join(get_portfolio_dir(portfolio), "cash.json")

def load_cash(portfolio="miguel"):
    path = get_cash_path(portfolio)
    try:
        with open(path) as f:
            return json.load(f).get("balance", 0)
    except:
        return 0

def save_cash(balance, portfolio="miguel"):
    path = get_cash_path(portfolio)
    with open(path, "w") as f:
        json.dump({"balance": round(float(balance), 2)}, f)

WATCHLIST_PATH    = os.path.join(BASE_DIR, "watchlist.json")

# ── Watchlist helpers ───────────────────────────────────────────
watchlist_lock = threading.Lock()

def load_watchlist():
    try:
        with open(WATCHLIST_PATH) as f:
            return json.load(f)
    except:
        return []

def save_watchlist(items):
    with open(WATCHLIST_PATH, "w") as f:
        json.dump(items, f, indent=2, ensure_ascii=False)

def fetch_watchlist_data(ticker):
    """Obtiene métricas completas + estimates de analistas para un ticker vía yfinance."""
    try:
        tk = yf.Ticker(ticker)
        info = tk.info or {}

        # ── Precio y mercado ──
        price     = info.get("currentPrice") or info.get("regularMarketPrice")
        prev      = info.get("previousClose") or info.get("regularMarketPreviousClose")
        chg_pct   = ((price - prev) / prev) if price and prev and prev != 0 else None
        mkt_cap   = info.get("marketCap")
        sector    = info.get("sector", "")
        industry  = info.get("industry", "")
        name      = info.get("shortName") or info.get("longName") or ticker
        currency  = info.get("currency", "USD")

        # ── 52W ──
        hi52  = info.get("fiftyTwoWeekHigh")
        lo52  = info.get("fiftyTwoWeekLow")
        vs_hi = ((price - hi52) / hi52) if price and hi52 else None
        vs_lo = ((price - lo52) / lo52) if price and lo52 else None

        # ── Múltiplos ──
        pe        = info.get("trailingPE")
        fwd_pe    = info.get("forwardPE")
        ev_ebitda = info.get("enterpriseToEbitda")
        pb        = info.get("priceToBook")
        ps        = info.get("priceToSalesTrailing12Months")

        # ── Fundamentales ──
        revenue       = info.get("totalRevenue")
        rev_growth    = info.get("revenueGrowth")
        gross_margin  = info.get("grossMargins")
        op_margin     = info.get("operatingMargins")
        net_margin    = info.get("profitMargins")
        eps_ttm       = info.get("trailingEps")
        eps_fwd       = info.get("forwardEps")
        fcf           = info.get("freeCashflow")
        debt_equity   = info.get("debtToEquity")
        roe           = info.get("returnOnEquity")

        # ── Riesgo ──
        beta      = info.get("beta")
        short_pct = info.get("shortPercentOfFloat")

        # ── Analistas — price target y rating ──
        pt_mean    = info.get("targetMeanPrice")
        pt_high    = info.get("targetHighPrice")
        pt_low     = info.get("targetLowPrice")
        pt_median  = info.get("targetMedianPrice")
        n_analysts = info.get("numberOfAnalystOpinions")
        rating     = info.get("recommendationKey", "")   # "buy", "hold", "sell", etc.
        rating_mean= info.get("recommendationMean")       # 1=Strong Buy … 5=Sell
        upside     = ((pt_mean - price) / price) if pt_mean and price and price != 0 else None

        # ── Estimates FY0 y FY+1 ──
        estimates = {"fy0": {}, "fy1": {}}
        try:
            earnings = tk.get_earnings_estimate()   # DataFrame
            if earnings is not None and not earnings.empty:
                cols = list(earnings.columns)
                if len(cols) >= 1:
                    estimates["fy0"]["eps_est"]  = float(earnings[cols[0]].get("Avg") or 0) or None
                    estimates["fy0"]["eps_low"]  = float(earnings[cols[0]].get("Low") or 0) or None
                    estimates["fy0"]["eps_high"] = float(earnings[cols[0]].get("High") or 0) or None
                if len(cols) >= 2:
                    estimates["fy1"]["eps_est"]  = float(earnings[cols[1]].get("Avg") or 0) or None
                    estimates["fy1"]["eps_low"]  = float(earnings[cols[1]].get("Low") or 0) or None
                    estimates["fy1"]["eps_high"] = float(earnings[cols[1]].get("High") or 0) or None
        except:
            pass

        try:
            rev_est = tk.get_revenue_estimate()
            if rev_est is not None and not rev_est.empty:
                cols = list(rev_est.columns)
                if len(cols) >= 1:
                    estimates["fy0"]["rev_est"]  = float(rev_est[cols[0]].get("Avg") or 0) or None
                    estimates["fy0"]["rev_low"]  = float(rev_est[cols[0]].get("Low") or 0) or None
                    estimates["fy0"]["rev_high"] = float(rev_est[cols[0]].get("High") or 0) or None
                    estimates["fy0"]["year"]     = str(cols[0])
                if len(cols) >= 2:
                    estimates["fy1"]["rev_est"]  = float(rev_est[cols[1]].get("Avg") or 0) or None
                    estimates["fy1"]["rev_low"]  = float(rev_est[cols[1]].get("Low") or 0) or None
                    estimates["fy1"]["rev_high"] = float(rev_est[cols[1]].get("High") or 0) or None
                    estimates["fy1"]["year"]     = str(cols[1])
        except:
            pass

        # Descripción de la compañía
        description = info.get("longBusinessSummary", "")

        # Retornos históricos + price_history para gráfica
        r1m = r3m = r6m = r1y = ytd = None
        price_history = []   # [{t: epoch_ms, p: price}, ...]
        try:
            import datetime as dt
            hist = tk.history(period="1y")
            if hist is not None and len(hist) > 0:
                cur_p = hist["Close"].iloc[-1]
                def ret(days):
                    idx = max(0, len(hist) - days)
                    return (cur_p - hist["Close"].iloc[idx]) / hist["Close"].iloc[idx]
                r1m = ret(21)
                r3m = ret(63)
                r6m = ret(126)
                r1y = ret(252) if len(hist) >= 252 else (cur_p - hist["Close"].iloc[0]) / hist["Close"].iloc[0]
                # YTD — compatible con tz-aware index
                try:
                    jan1 = dt.date(dt.date.today().year, 1, 1)
                    idx_dates = hist.index.date if hasattr(hist.index, 'date') else hist.index.to_pydatetime()
                    ytd_mask = [d >= jan1 for d in (hist.index.date if hasattr(hist.index, 'date') else [x.date() for x in hist.index.to_pydatetime()])]
                    ytd_rows = hist[ytd_mask]
                    if len(ytd_rows) > 0:
                        ytd = (cur_p - ytd_rows["Close"].iloc[0]) / ytd_rows["Close"].iloc[0]
                except Exception as e_ytd:
                    print(f"  [WL] YTD calc error: {e_ytd}", flush=True)
                # Vol 30d
                try:
                    returns = hist["Close"].pct_change(fill_method=None).dropna()
                    vol_30d = float(returns.tail(30).std() * (252**0.5)) if len(returns) >= 10 else None
                except:
                    vol_30d = None
                # Max drawdown 1Y
                try:
                    roll_max = hist["Close"].cummax()
                    drawdown = (hist["Close"] - roll_max) / roll_max
                    max_dd   = float(drawdown.min())
                except:
                    max_dd = None
                # Serializar precio histórico (weekly para no enviar demasiado)
                try:
                    hist_weekly = hist["Close"].resample("W").last()
                    price_history = [
                        {"t": int(ts.timestamp() * 1000), "p": round(float(p), 2)}
                        for ts, p in hist_weekly.items()
                        if p == p and p is not None  # excluir NaN
                    ]
                    print(f"  [WL] {ticker} price_history: {len(price_history)} puntos", flush=True)
                except Exception as e_ph:
                    print(f"  [WL] price_history error: {e_ph}", flush=True)
                    price_history = []
            else:
                vol_30d = max_dd = None
        except Exception as e_hist:
            print(f"  [WL] hist error for {ticker}: {e_hist}", flush=True)
            vol_30d = max_dd = None

        return {
            "ticker":      ticker.upper(),
            "name":        name,
            "price":       price,
            "chg_pct":     chg_pct,
            "currency":    currency,
            "mkt_cap":     mkt_cap,
            "sector":      sector,
            "industry":    industry,
            "hi52":        hi52,
            "lo52":        lo52,
            "vs_52hi":     vs_hi,
            "vs_52lo":     vs_lo,
            "pe":          pe,
            "fwd_pe":      fwd_pe,
            "ev_ebitda":   ev_ebitda,
            "pb":          pb,
            "ps":          ps,
            "revenue":     revenue,
            "rev_growth":  rev_growth,
            "gross_margin":gross_margin,
            "op_margin":   op_margin,
            "net_margin":  net_margin,
            "eps_ttm":     eps_ttm,
            "eps_fwd":     eps_fwd,
            "fcf":         fcf,
            "debt_equity": debt_equity,
            "roe":         roe,
            "beta":        beta,
            "vol_30d":     vol_30d,
            "max_dd":      max_dd,
            "short_pct":   short_pct,
            "pt_mean":     pt_mean,
            "pt_high":     pt_high,
            "pt_low":      pt_low,
            "pt_median":   pt_median,
            "n_analysts":  n_analysts,
            "rating":      rating,
            "rating_mean": rating_mean,
            "upside":      upside,
            "r1m":         r1m,
            "r3m":         r3m,
            "r6m":         r6m,
            "r1y":         r1y,
            "ytd":         ytd,
            "estimates":     estimates,
            "price_history": price_history,
            "description":   description,
            "ok":            True,
        }
    except Exception as e:
        return {"ticker": ticker.upper(), "ok": False, "error": str(e)}

# Legacy paths (fallback para archivos que aún están en raíz)
POSITIONS_PATH    = os.path.join(BASE_DIR, "positions.json")
TRANSACTIONS_PATH = os.path.join(BASE_DIR, "transactions.json")

# ── Transacciones ──────────────────────────────────────────────
transactions_lock = threading.Lock()

def load_transactions(portfolio="miguel"):
    """Lee transacciones desde portfolios/{portfolio}/transactions.json."""
    path = get_transactions_path(portfolio)
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []

def save_transactions(txs, portfolio="miguel"):
    """Guarda transacciones en portfolios/{portfolio}/transactions.json."""
    path = get_transactions_path(portfolio)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(txs, f, ensure_ascii=False, indent=2)

def compute_positions_from_transactions(transactions):
    """
    Calcula posiciones actuales desde transacciones.
    Método: promedio ponderado (weighted average cost).
    Returns list of {ticker, units, avg_price}.
    """
    holdings = {}  # ticker -> {units, avg_price}
    for tx in sorted(transactions, key=lambda x: (x.get("date", ""), x.get("id", ""))):
        ticker = tx.get("ticker", "").upper().strip()
        qty    = float(tx.get("quantity", 0))
        price  = float(tx.get("price", 0))
        if not ticker or qty <= 0 or price < 0:
            continue
        if ticker not in holdings:
            holdings[ticker] = {"units": 0.0, "avg_price": 0.0}
        h = holdings[ticker]
        if tx["type"] == "buy":
            total_cost   = h["units"] * h["avg_price"] + qty * price
            h["units"]  += qty
            h["avg_price"] = total_cost / h["units"] if h["units"] > 0 else 0
        elif tx["type"] == "sell":
            h["units"] = max(0.0, h["units"] - qty)
            # avg_price sin cambio (weighted average method)
    return [
        {"ticker": t, "units": round(v["units"], 8), "avg_price": round(v["avg_price"], 6)}
        for t, v in holdings.items()
        if v["units"] > 0.0000001
    ]

# Inicializar archivos si no existen
for _pname in PORTFOLIO_NAMES:
    _txpath = get_transactions_path(_pname)
    if not os.path.exists(_txpath):
        with open(_txpath, "w") as _f:
            json.dump([], _f)
    _anpath = get_analysis_path(_pname)
    if not os.path.exists(_anpath):
        with open(_anpath, "w") as _f:
            json.dump({}, _f)

TRANSACTIONS = load_transactions("miguel")
transactions_lock = threading.Lock()

# ── Posiciones del portafolio ──────────────────────────────────
positions_lock = threading.Lock()

def load_positions(portfolio="miguel"):
    """Calcula posiciones actuales desde transacciones del portafolio."""
    return compute_positions_from_transactions(load_transactions(portfolio))

def save_positions(positions):
    """No-op: posiciones se derivan de transacciones."""
    pass

POSITIONS = load_positions("miguel")

# ── Activos Macro ─────────────────────────────────────────────
MACRO = {
    "equities": [
        {"ticker": "^GSPC",     "name": "S&P 500"},
        {"ticker": "^IXIC",     "name": "NASDAQ"},
        {"ticker": "^DJI",      "name": "Dow Jones"},
        {"ticker": "^STOXX50E", "name": "Euro Stoxx 50"},
        {"ticker": "^GDAXI",    "name": "DAX"},
        {"ticker": "^FTSE",     "name": "FTSE 100"},
        {"ticker": "^IBEX",     "name": "IBEX 35"},
        {"ticker": "^BVSP",     "name": "Bovespa"},
        {"ticker": "^N225",     "name": "Nikkei 225"},
        {"ticker": "^KS11",     "name": "KOSPI"},
        {"ticker": "^HSI",      "name": "Hang Seng"},
        {"ticker": "000001.SS", "name": "Shanghai"},
        {"ticker": "ACWI",      "name": "ACWI"},
        {"ticker": "^RUT",      "name": "Russell 2000"},
        {"ticker": "^VIX",      "name": "VIX"},
    ],
    "bonds": [
        {"ticker": "SPY",  "name": "S&P 500 ETF"},
        {"ticker": "TLT",  "name": "TLT 20yr"},
        {"ticker": "^TNX", "name": "US 10Y", "is_yield": True},
        {"ticker": "^IRX", "name": "US 2Y",  "is_yield": True},
        {"ticker": "^TYX", "name": "US 30Y", "is_yield": True},
        {"ticker": "HYG",  "name": "High Yield"},
    ],
    "commodities": [
        {"ticker": "GC=F",  "name": "Oro"},
        {"ticker": "SI=F",  "name": "Plata"},
        {"ticker": "HG=F",  "name": "Cobre"},
        {"ticker": "PL=F",  "name": "Platino"},
        {"ticker": "CL=F",  "name": "Petróleo WTI"},
        {"ticker": "BZ=F",  "name": "Petróleo Brent"},
        {"ticker": "NG=F",  "name": "Gas Natural"},
        {"ticker": "ZC=F",  "name": "Maíz"},
        {"ticker": "KC=F",  "name": "Café"},
        {"ticker": "PA=F",  "name": "Paladio"},
        {"ticker": "CC=F",  "name": "Cacao"},
        {"ticker": "SB=F",  "name": "Azúcar"},
    ],
    "fx": [
        {"ticker": "EURUSD=X", "name": "EUR/USD"},
        {"ticker": "JPYUSD=X", "name": "JPY/USD"},
        {"ticker": "GBPUSD=X", "name": "GBP/USD"},
        {"ticker": "MXN=X",    "name": "USD/MXN"},
        {"ticker": "DX-Y.NYB", "name": "DXY"},
    ],
    "crypto": [
        {"ticker": "BTC-USD",  "name": "Bitcoin"},
        {"ticker": "ETH-USD",  "name": "Ethereum"},
        {"ticker": "SOL-USD",  "name": "Solana"},
        {"ticker": "BNB-USD",  "name": "BNB"},
        {"ticker": "XRP-USD",  "name": "XRP"},
        {"ticker": "DOGE-USD", "name": "Dogecoin"},
        {"ticker": "ADA-USD",  "name": "Cardano"},
    ],
}

# ── Análisis personal (desde JSON) ───────────────────────────
ANALYSIS_PATH = os.path.join(BASE_DIR, "analysis.json")  # legacy fallback
analysis_lock = threading.Lock()

def load_analysis(portfolio="miguel"):
    path = get_analysis_path(portfolio)
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return {}

def save_analysis(data, portfolio="miguel"):
    path = get_analysis_path(portfolio)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

ANALYSIS = load_analysis("miguel")

# ── Cache de datos (separado por portafolio) ──────────────────
def _empty_cache():
    return {
        "portfolio": [],
        "macro": {},
        "metrics": {},
        "last_update": None,
        "market_open": False,
    }

cache = {
    "miguel": _empty_cache(),
    "paulo":  _empty_cache(),
    "akila":  _empty_cache(),
}
cache_lock = threading.Lock()

# ── Caché de tearsheet (por ticker, TTL 20 min) ──
tearsheet_cache      = {}   # { ticker: {"data": {...}, "ts": float} }
tearsheet_cache_lock = threading.Lock()
TEARSHEET_TTL        = 1200  # 20 minutos

risk_cache      = {}   # { portfolio: {"data": {...}, "ts": float} }
risk_cache_lock = threading.Lock()
RISK_TTL        = 3600  # 1 hora
_risk_computing = set()   # portfolios cuyo cálculo está en progreso

def get_tearsheet_cached(ticker, portfolio="miguel"):
    """Devuelve tearsheet desde caché si es reciente, si no lo fetcha y cachea.
    La clave incluye el portfolio para que pos_data sea correcto por portafolio."""
    import time
    cache_key = f"{ticker}::{portfolio}"
    with tearsheet_cache_lock:
        entry = tearsheet_cache.get(cache_key)
        if entry and (time.time() - entry["ts"]) < TEARSHEET_TTL:
            print(f"  📊 Tearsheet cache HIT: {cache_key}", flush=True)
            return entry["data"]
    print(f"  📊 Tearsheet cache MISS: {cache_key} — fetching…", flush=True)
    data = fetch_tearsheet(ticker, portfolio)
    with tearsheet_cache_lock:
        tearsheet_cache[cache_key] = {"data": data, "ts": time.time()}
    return data

# Alias para compatibilidad con código que usa cache["portfolio"], etc.
# Se resolverá accediendo siempre con get_cache(portfolio)
def get_cache(portfolio="miguel"):
    if portfolio not in cache:
        cache[portfolio] = _empty_cache()
    return cache[portfolio]

def is_market_open():
    """Verifica si el mercado US está abierto (9:30-16:00 ET, L-V)."""
    from datetime import timezone
    import datetime as dt
    now_utc = dt.datetime.now(timezone.utc)
    # ET = UTC-4 (EDT) o UTC-5 (EST)
    et_offset = -4  # EDT (mayo)
    now_et = now_utc.replace(tzinfo=None) + dt.timedelta(hours=et_offset)
    if now_et.weekday() >= 5:  # sábado/domingo
        return False
    market_open  = now_et.replace(hour=9, minute=30, second=0, microsecond=0)
    market_close = now_et.replace(hour=16, minute=0, second=0, microsecond=0)
    return market_open <= now_et <= market_close

def get_market_session():
    """Devuelve el estado actual del mercado US: OPEN, PRE, POST, CLOSED."""
    from datetime import timezone
    import datetime as dt
    now_utc = dt.datetime.now(timezone.utc)
    # Detectar offset ET automáticamente (EDT = UTC-4, EST = UTC-5)
    # DST en US: segundo domingo de marzo → primer domingo de noviembre
    year = now_utc.year
    # Segundo domingo de marzo
    march1 = dt.datetime(year, 3, 1, tzinfo=timezone.utc)
    dst_start = march1 + dt.timedelta(days=(6 - march1.weekday()) % 7 + 7)
    # Primer domingo de noviembre
    nov1 = dt.datetime(year, 11, 1, tzinfo=timezone.utc)
    dst_end = nov1 + dt.timedelta(days=(6 - nov1.weekday()) % 7)
    is_edt = dst_start <= now_utc < dst_end
    et_offset = -4 if is_edt else -5
    now_et = now_utc.replace(tzinfo=None) + dt.timedelta(hours=et_offset)

    if now_et.weekday() >= 5:
        return "CLOSED"

    pre_open   = now_et.replace(hour=4,  minute=0,  second=0, microsecond=0)
    mkt_open   = now_et.replace(hour=9,  minute=30, second=0, microsecond=0)
    mkt_close  = now_et.replace(hour=16, minute=0,  second=0, microsecond=0)
    post_close = now_et.replace(hour=20, minute=0,  second=0, microsecond=0)

    if mkt_open <= now_et < mkt_close:
        return "OPEN"
    elif pre_open <= now_et < mkt_open:
        return "PRE"
    elif mkt_close <= now_et < post_close:
        return "POST"
    else:
        return "CLOSED"

def fetch_price(ticker, is_yield=False):
    """Jala precio actual desde caché de precios batch; fallback a yf.Ticker si no hay cache."""
    import math
    entry = _price_cache.get(ticker)
    if entry:
        price, prev = entry["price"], entry["prev"]
        if is_yield and price and price > 20:
            price /= 10; prev = (prev or price) / 10
        chg_pct = (price - prev) / prev if prev and prev != 0 else 0
        return price, prev, chg_pct
    # Fallback individual si aún no hay batch
    try:
        t    = yf.Ticker(ticker, session=_YF_SESSION)
        info = t.info
        price = float(info.get("regularMarketPrice") or info.get("currentPrice") or info.get("regularMarketPreviousClose") or 0)
        prev  = float(info.get("regularMarketPreviousClose") or info.get("previousClose") or price)
        if not price: return None, None, None
        if is_yield and price > 20: price /= 10; prev /= 10
        chg_pct = (price - prev) / prev if prev else 0
        return price, prev, chg_pct
    except:
        return None, None, None

# ── Cache global de precios batch (actualizado en update_data) ──
_price_cache = {}   # { ticker: {"price": float, "prev": float} }
_info_cache  = {}   # { ticker: {"name":..., "sector":..., ...}, "ts": float }
_INFO_TTL    = 3600  # refrescar info (nombre, sector) cada 1 hora

def _yf_download_safe(tickers, timeout=25, **kwargs):
    """Wrapper de yf.download con timeout para evitar que el thread se quede colgado."""
    import concurrent.futures
    with concurrent.futures.ThreadPoolExecutor(max_workers=1) as ex:
        future = ex.submit(yf.download, tickers, **kwargs)
        try:
            return future.result(timeout=timeout)
        except concurrent.futures.TimeoutError:
            print(f"  ⚠️  yf.download timeout ({timeout}s) — saltando ciclo", flush=True)
            return None
        except Exception as e:
            print(f"  ⚠️  yf.download error: {e}", flush=True)
            return None

def _batch_update_prices(tickers):
    """Descarga precios de todos los tickers en una sola llamada yf.download."""
    import math
    if not tickers:
        return
    try:
        raw = _yf_download_safe(list(tickers), period="2d", interval="1d",
                               progress=False, auto_adjust=True, session=_YF_SESSION)
        if raw is None or raw.empty:
            return
        import pandas as pd
        # MultiIndex (field, ticker) con group_by default "column"
        if isinstance(raw.columns, pd.MultiIndex):
            lvl0 = raw.columns.get_level_values(0)
            lvl1 = raw.columns.get_level_values(1)
            if "Close" in lvl0:
                closes = raw["Close"]          # (field, ticker) → standard
            elif "Close" in lvl1:
                closes = raw.xs("Close", axis=1, level=1)  # (ticker, field) → fallback
            else:
                return
        else:
            # Single ticker — columnas planas
            closes = raw[["Close"]].rename(columns={"Close": list(tickers)[0]})
        if closes is None or closes.empty:
            return
        for tk in tickers:
            try:
                if tk not in closes.columns:
                    continue
                col = closes[tk].dropna()
                if len(col) < 1:
                    continue
                price = float(col.iloc[-1])
                prev  = float(col.iloc[-2]) if len(col) >= 2 else price
                if math.isnan(price):
                    continue
                _price_cache[tk] = {"price": price, "prev": prev}
            except:
                pass
    except Exception as e:
        print(f"  ⚠️  batch_update_prices error: {e}", flush=True)

def _get_info_cached(ticker):
    """Devuelve info de yfinance desde caché (TTL 1h). No bloquea si ya hay datos."""
    import time
    entry = _info_cache.get(ticker)
    if entry and (time.time() - entry.get("ts", 0)) < _INFO_TTL:
        return entry
    try:
        info = yf.Ticker(ticker, session=_YF_SESSION).info
        data = {
            "name":       info.get("longName") or info.get("shortName") or ticker,
            "sector":     info.get("sector", ""),
            "country":    info.get("country", ""),
            "quote_type": (info.get("quoteType") or "").upper(),
            "pt_consenso":info.get("targetMeanPrice"),
            "recomendacion": (info.get("recommendationKey") or "").upper(),
            "num_analistas": info.get("numberOfAnalystOpinions"),
            "ts": time.time(),
        }
        _info_cache[ticker] = data
        return data
    except:
        return _info_cache.get(ticker) or {"name": ticker, "sector": "", "country": "", "quote_type": "", "ts": 0}

def fetch_metrics(ticker):
    """Jala métricas completas para la pestaña de análisis."""
    try:
        t    = yf.Ticker(ticker)
        info = t.info

        # Rendimientos históricos
        hist_1y = t.history(period="1y",  interval="1d")
        hist_3m = t.history(period="3mo", interval="1d")
        hist_1m = t.history(period="1mo", interval="1d")

        def ret(hist):
            if hist.empty or len(hist) < 2:
                return None
            return (float(hist["Close"].iloc[-1]) - float(hist["Close"].iloc[0])) / float(hist["Close"].iloc[0])

        # YTD: desde 1 enero 2026
        import datetime as dt
        start_ytd = dt.datetime(dt.datetime.now().year, 1, 1, tzinfo=timezone.utc)
        hist_ytd  = t.history(start=start_ytd, interval="1d")
        ytd = ret(hist_ytd)
        r1m = ret(hist_1m)
        r3m = ret(hist_3m)
        r1y = ret(hist_1y)

        # Max Drawdown (52 semanas)
        max_dd = None
        if not hist_1y.empty:
            closes   = hist_1y["Close"]
            roll_max = closes.cummax()
            dd       = (closes - roll_max) / roll_max
            max_dd   = float(dd.min())

        # Volatilidad 30d (anualizada)
        vol_30d = None
        if not hist_1m.empty and len(hist_1m) > 5:
            returns  = hist_1m["Close"].pct_change(fill_method=None).dropna()
            vol_30d  = float(returns.std() * (252 ** 0.5))

        # 52W High/Low
        high_52w = info.get("fiftyTwoWeekHigh")
        low_52w  = info.get("fiftyTwoWeekLow")
        price    = info.get("currentPrice") or info.get("regularMarketPrice")
        vs_52w_h = (price - high_52w) / high_52w if (price and high_52w) else None

        return {
            # Rendimientos
            "ytd":     ytd,
            "r1m":     r1m,
            "r3m":     r3m,
            "r1y":     r1y,
            # Riesgo
            "beta":    info.get("beta"),
            "vol_30d": vol_30d,
            "max_dd":  max_dd,
            "vs_52wh": vs_52w_h,
            "high_52w":high_52w,
            "low_52w": low_52w,
            # Múltiplos
            "pe":      info.get("trailingPE"),
            "fwd_pe":  info.get("forwardPE"),
            "ev_ebitda":info.get("enterpriseToEbitda"),
            "pb":      info.get("priceToBook"),
            "ps":      info.get("priceToSalesTrailing12Months"),
            # Fundamentales
            "revenue":       info.get("totalRevenue"),
            "revenue_growth":info.get("revenueGrowth"),
            "gross_margin":  info.get("grossMargins"),
            "op_margin":     info.get("operatingMargins"),
            "net_margin":    info.get("profitMargins"),
            "eps_ttm":       info.get("trailingEps"),
            "eps_fwd":       info.get("forwardEps"),
            "fcf":           info.get("freeCashflow"),
            "mkt_cap":       info.get("marketCap"),
            "debt_equity":   info.get("debtToEquity"),
            "roe":           info.get("returnOnEquity"),
        }
    except Exception as e:
        print(f"  ⚠️  Metrics error {ticker}: {e}")
        return {}

def build_portfolio_data(portfolio="miguel"):
    """Construye los datos de portafolio para un portafolio específico."""
    ana_data = load_analysis(portfolio)
    current_positions = compute_positions_from_transactions(load_transactions(portfolio))
    total_cost = sum(p["units"] * p["avg_price"] for p in current_positions)
    total_mkt  = 0
    raw = []

    for pos in current_positions:
        price, prev, chg_pct = fetch_price(pos["ticker"])
        cost    = pos["units"] * pos["avg_price"]
        mkt_val = pos["units"] * price if price else cost
        pnl_usd = mkt_val - cost
        pnl_pct = pnl_usd / cost if cost else 0
        total_mkt += mkt_val
        ana = ana_data.get(pos["ticker"], {})
        cached        = _get_info_cached(pos["ticker"])
        name          = cached.get("name", pos["ticker"])
        sector        = cached.get("sector", "")
        country       = cached.get("country", "")
        quote_type    = cached.get("quote_type", "")
        pt_consenso   = cached.get("pt_consenso")
        recomendacion = cached.get("recomendacion")
        num_analistas = cached.get("num_analistas")

        raw.append({
            "ticker":        pos["ticker"],
            "name":          name,
            "sector":        sector,
            "country":       country,
            "is_etf":        quote_type == "ETF",
            "units":         pos["units"],
            "avg_price":     pos["avg_price"],
            "price":         price,
            "chg_pct":       chg_pct,
            "cost":          cost,
            "mkt_val":       mkt_val,
            "pnl_usd":       pnl_usd,
            "pnl_pct":       pnl_pct,
            "bull_pt":       ana.get("bull_pt"),
            "base_pt":       ana.get("base_pt"),
            "pt_consenso":   pt_consenso,
            "recomendacion": recomendacion,
            "num_analistas": num_analistas,
            "conviction":    ana.get("conviction", "—"),
            "next_cat":      ana.get("next_cat", "—"),
            "tesis":         ana.get("tesis", ""),
        })

    portfolio_data = []
    for item in raw:
        base_pt = item["base_pt"]
        p = item["price"]
        item["weight"] = item["mkt_val"] / total_mkt if total_mkt else 0
        item["upside"] = (base_pt - p) / p if (p and base_pt) else None
        portfolio_data.append(item)

    return portfolio_data, total_cost, total_mkt

def build_consolidated_data():
    """Construye el portafolio consolidado sumando Miguel + Paulo."""
    miguel_data, miguel_cost, miguel_mkt = build_portfolio_data("miguel")
    paulo_data,  paulo_cost,  paulo_mkt  = build_portfolio_data("paulo")

    total_cost = miguel_cost + paulo_cost
    total_mkt  = miguel_mkt  + paulo_mkt

    # Merge por ticker con weighted average cost
    merged = {}
    for item in miguel_data + paulo_data:
        tk = item["ticker"]
        item_cost = item["cost"]   # ya calculado correctamente en build_portfolio_data
        if tk not in merged:
            merged[tk] = dict(item)
            merged[tk]["_cost_basis"] = item_cost
        else:
            existing       = merged[tk]
            combined_cost  = existing["_cost_basis"] + item_cost
            combined_units = existing["units"] + item["units"]
            combined_mkt   = existing["mkt_val"] + item["mkt_val"]
            existing["units"]      = combined_units
            existing["cost"]       = combined_cost
            existing["_cost_basis"]= combined_cost
            existing["avg_price"]  = combined_cost / combined_units if combined_units else 0
            existing["mkt_val"]    = combined_mkt
            existing["pnl_usd"]    = combined_mkt - combined_cost
            existing["pnl_pct"]    = existing["pnl_usd"] / combined_cost if combined_cost else 0

    consolidated = []
    for item in merged.values():
        item.pop("_cost_basis", None)
        item["weight"] = item["mkt_val"] / total_mkt if total_mkt else 0
        p = item["price"]; bp = item["base_pt"]
        item["upside"] = (bp - p) / p if (p and bp) else None
        consolidated.append(item)

    return consolidated, total_cost, total_mkt

def update_metrics_background():
    """Actualiza métricas en ciclo separado, cada 5 minutos."""
    while True:
        try:
            print(f"  📐 Actualizando métricas... {datetime.now().strftime('%H:%M:%S')}", flush=True)
            # Unión de tickers de todos los portafolios
            all_tickers = set()
            for pname in PORTFOLIO_NAMES:
                for pos in load_positions(pname):
                    all_tickers.add(pos["ticker"])
            metrics_data = {}
            for ticker in all_tickers:
                metrics_data[ticker] = fetch_metrics(ticker)
            with cache_lock:
                for pname in PORTFOLIO_NAMES + ["akila"]:
                    get_cache(pname)["metrics"] = metrics_data
            print(f"  ✅ Métricas actualizadas ({len(all_tickers)} tickers).", flush=True)
        except Exception as e:
            print(f"  ⚠️  Error métricas: {e}", flush=True)
        time.sleep(300)

def update_data():
    """Actualiza todos los precios en background para todos los portafolios."""
    while True:
        try:
            print(f"  🔄 Actualizando precios... {datetime.now().strftime('%H:%M:%S')}", flush=True)

            import datetime as dt
            import math
            def hist_rows(df):
                rows = []
                if df is None or df.empty: return rows
                for date, row in df.iterrows():
                    try:
                        c = float(row["Close"])
                        if not math.isnan(c):
                            rows.append({"date": date.strftime("%Y-%m-%d"), "close": c})
                    except: pass
                return rows

            # ── Batch precios (1 request cubre macro + portafolio) ──
            all_macro_tickers = [a["ticker"] for assets in MACRO.values() for a in assets]
            all_port_tickers  = list({
                pos["ticker"]
                for pname in PORTFOLIO_NAMES
                for pos in compute_positions_from_transactions(load_transactions(pname))
            })
            _batch_update_prices(set(all_macro_tickers + all_port_tickers))

            # ── Macro (batch history download) ──
            start_ytd = dt.datetime(dt.datetime.now().year, 1, 1, tzinfo=timezone.utc)
            try:
                batch_ytd = _yf_download_safe(
                    all_macro_tickers, timeout=30, start=start_ytd,
                    progress=False, auto_adjust=True, session=_YF_SESSION
                )
            except Exception as e:
                print(f"  ⚠️  Macro batch download error: {e}", flush=True)
                batch_ytd = None

            def get_batch_series(df, ticker):
                """Extrae la serie Close de un batch download multi-ticker."""
                if df is None or df.empty:
                    return None
                try:
                    pd = __import__("pandas")
                    if isinstance(df.columns, pd.MultiIndex):
                        lvl0 = df.columns.get_level_values(0)
                        lvl1 = df.columns.get_level_values(1)
                        if "Close" in lvl0:
                            return df["Close"][ticker].dropna() if ticker in df["Close"].columns else None
                        elif "Close" in lvl1:
                            xs = df.xs("Close", axis=1, level=1)
                            return xs[ticker].dropna() if ticker in xs.columns else None
                        return None
                    else:
                        # Single ticker descargado
                        return df["Close"].dropna() if len(all_macro_tickers) == 1 else None
                except:
                    return None

            macro_data = {}
            for section, assets in MACRO.items():
                macro_data[section] = []
                for asset in assets:
                    is_yield = asset.get("is_yield", False)
                    price, prev, chg_pct = fetch_price(asset["ticker"], is_yield)
                    h5d, hytd = [], []
                    try:
                        series = get_batch_series(batch_ytd, asset["ticker"])
                        if series is not None and len(series) > 0:
                            hytd = [{"date": d.strftime("%Y-%m-%d"), "close": round(float(v), 4)}
                                    for d, v in series.items() if not math.isnan(float(v))]
                            h5d  = hytd[-5:] if len(hytd) >= 5 else hytd
                    except: pass
                    macro_data[section].append({
                        "ticker":   asset["ticker"],
                        "name":     asset["name"],
                        "price":    price,
                        "chg_pct":  chg_pct,
                        "is_yield": is_yield,
                        "hist_5d":  h5d,
                        "hist_ytd": hytd,
                    })

            now_str    = datetime.now().strftime("%d %b %Y  %H:%M:%S")
            mkt_open   = is_market_open()

            # Extraer USD/MXN del macro data
            mxn_px = 0
            for asset in macro_data.get("fx", []):
                if asset.get("ticker") == "MXN=X" and asset.get("price"):
                    mxn_px = float(asset["price"])
                    break

            # ── Portafolios individuales ───────────────────────
            for pname in PORTFOLIO_NAMES:
                try:
                    port_data, total_cost, total_mkt = build_portfolio_data(pname)
                    total_pnl = total_mkt - total_cost
                    with cache_lock:
                        c = get_cache(pname)
                        c["portfolio"]     = port_data
                        c["total_cost"]    = total_cost
                        c["total_mkt"]     = total_mkt
                        c["total_pnl"]     = total_pnl
                        c["total_pnl_pct"] = total_pnl / total_cost if total_cost else 0
                        c["macro"]         = macro_data
                        c["last_update"]   = now_str
                        c["market_open"]   = mkt_open
                        c["usd_mxn"]       = mxn_px if mxn_px else 0
                        c["cash"]          = load_cash(pname)
                    print(f"  ✅ [{pname}] P&L: ${total_pnl:+,.2f}", flush=True)
                except Exception as e:
                    print(f"  ⚠️  Error portafolio {pname}: {e}", flush=True)

            # ── Consolidado AKILA ──────────────────────────────
            try:
                cons_data, cons_cost, cons_mkt = build_consolidated_data()
                cons_pnl = cons_mkt - cons_cost
                with cache_lock:
                    c = get_cache("akila")
                    c["portfolio"]     = cons_data
                    c["total_cost"]    = cons_cost
                    c["total_mkt"]     = cons_mkt
                    c["total_pnl"]     = cons_pnl
                    c["total_pnl_pct"] = cons_pnl / cons_cost if cons_cost else 0
                    c["macro"]         = macro_data
                    c["last_update"]   = now_str
                    c["market_open"]   = mkt_open
                    c["usd_mxn"]       = mxn_px if mxn_px else 0
                    c["cash"]          = load_cash("miguel") + load_cash("paulo")
            except Exception as e:
                print(f"  ⚠️  Error consolidado: {e}", flush=True)

        except Exception as e:
            print(f"  ⚠️  Error en update: {e}", flush=True)

        time.sleep(REFRESH_SEC)

def fetch_analysis_live(ticker):
    """
    Jala datos de analistas en tiempo real desde yfinance:
    - Price targets (low/mean/median/high)
    - Distribución de recomendaciones (Strong Buy/Buy/Hold/Sell/Strong Sell)
    - Upgrades/Downgrades recientes
    - Noticias recientes
    """
    import math

    def safe(val):
        if val is None: return None
        if isinstance(val, float) and (math.isnan(val) or math.isinf(val)): return None
        return val

    try:
        t    = yf.Ticker(ticker)
        info = t.info

        # ── Price targets ─────────────────────────────────────
        pt_low    = safe(info.get("targetLowPrice"))
        pt_mean   = safe(info.get("targetMeanPrice"))
        pt_median = safe(info.get("targetMedianPrice"))
        pt_high   = safe(info.get("targetHighPrice"))
        price_now = safe(info.get("regularMarketPrice") or info.get("currentPrice"))
        num_analysts = info.get("numberOfAnalystOpinions")
        rec_key   = (info.get("recommendationKey") or "").upper()

        # ── Upside implícito ──────────────────────────────────
        upside_mean = (pt_mean - price_now) / price_now if (pt_mean and price_now) else None

        # ── Distribución recomendaciones ──────────────────────
        rec_dist = {}
        try:
            recs = t.recommendations
            if recs is not None and not recs.empty:
                # Puede tener columnas: strongBuy, buy, hold, sell, strongSell
                latest = recs.iloc[-1] if len(recs) > 0 else None
                if latest is not None:
                    for col in ["strongBuy","buy","hold","sell","strongSell"]:
                        val = latest.get(col)
                        if val is not None:
                            rec_dist[col] = int(val)
        except:
            pass

        # ── Upgrades / Downgrades (últimos 10) ────────────────
        upgrades_list = []
        try:
            upg = t.upgrades_downgrades
            if upg is not None and not upg.empty:
                # Ordenar por fecha desc, tomar 10
                upg_sorted = upg.sort_index(ascending=False).head(10)
                for date, row in upg_sorted.iterrows():
                    upgrades_list.append({
                        "date":   date.strftime("%Y-%m-%d") if hasattr(date, "strftime") else str(date),
                        "firm":   str(row.get("Firm", "—")),
                        "action": str(row.get("Action", "—")),      # up/down/main/init/reit
                        "from":   str(row.get("FromGrade", "—")),
                        "to":     str(row.get("ToGrade", "—")),
                    })
        except:
            pass

        # ── Noticias recientes (últimas 8) ────────────────────
        news_list = []
        try:
            import datetime as dt
            news = t.news
            if news:
                for item in news[:8]:
                    # yfinance >=0.2.x puede devolver estructura anidada
                    # Intentar extraer desde content.title / content.pubDate / etc.
                    content = item.get("content", {}) if isinstance(item, dict) else {}

                    # Title
                    title = (item.get("title")
                             or content.get("title")
                             or content.get("headline", ""))

                    # Publisher
                    pub_raw = (item.get("publisher")
                               or content.get("provider", {}).get("displayName", "")
                               or content.get("source", ""))

                    # URL
                    url = (item.get("link")
                           or item.get("url")
                           or content.get("canonicalUrl", {}).get("url", "")
                           or content.get("clickThroughUrl", {}).get("url", ""))

                    # Date
                    pub_date = ""
                    pub_ts = item.get("providerPublishTime") or item.get("pubDate")
                    if pub_ts:
                        try:
                            if isinstance(pub_ts, (int, float)):
                                pub_date = dt.datetime.utcfromtimestamp(pub_ts).strftime("%d %b")
                            else:
                                pub_date = str(pub_ts)[:10]
                        except:
                            pass
                    if not pub_date:
                        raw_date = content.get("pubDate", "") or content.get("displayTime", "")
                        if raw_date:
                            pub_date = str(raw_date)[:10]

                    if title:
                        news_list.append({
                            "title":     title,
                            "publisher": pub_raw,
                            "date":      pub_date,
                            "url":       url,
                        })
        except Exception as ne:
            print(f"  ⚠️  News parse error {ticker}: {ne}", flush=True)

        return {
            "ticker":       ticker,
            "price":        price_now,
            "rec_key":      rec_key,
            "num_analysts": num_analysts,
            "pt_low":       pt_low,
            "pt_mean":      pt_mean,
            "pt_median":    pt_median,
            "pt_high":      pt_high,
            "upside_mean":  safe(upside_mean),
            "rec_dist":     rec_dist,
            "upgrades":     upgrades_list,
            "news":         news_list,
        }

    except Exception as e:
        print(f"  ⚠️  Analysis live error {ticker}: {e}", flush=True)
        return {"error": str(e)}


def fetch_news():
    """
    Jala noticias financieras de múltiples RSS feeds gratuitos.
    Fuentes: Reuters, CNBC, MarketWatch, Yahoo Finance, FT, Investing.com, Bloomberg.
    Rankea por importancia usando keywords y frescura.
    """
    import urllib.request
    import xml.etree.ElementTree as ET
    import datetime as dt
    import html
    import re

    FEEDS = [
        {"name": "CNBC Markets",        "url": "https://search.cnbc.com/rs/search/combinedcms/view.xml?partnerId=wrss01&id=20910258", "weight": 9},
        {"name": "CNBC Economy",        "url": "https://search.cnbc.com/rs/search/combinedcms/view.xml?partnerId=wrss01&id=15839135", "weight": 9},
        {"name": "MarketWatch Top",     "url": "https://feeds.marketwatch.com/marketwatch/topstories/",   "weight": 8},
        {"name": "MarketWatch Markets", "url": "https://feeds.marketwatch.com/marketwatch/marketpulse/",  "weight": 8},
        {"name": "Yahoo Finance",       "url": "https://finance.yahoo.com/news/rssindex",                 "weight": 7},
        {"name": "FT Markets",          "url": "https://www.ft.com/markets?format=rss",                   "weight": 9},
        {"name": "Investing.com",       "url": "https://www.investing.com/rss/news.rss",                  "weight": 7},
        {"name": "Bloomberg Markets",   "url": "https://feeds.bloomberg.com/markets/news.rss",            "weight": 9},
        {"name": "WSJ Markets",         "url": "https://feeds.content.dowjones.io/public/rss/mw_topstories", "weight": 8},
        {"name": "Seeking Alpha",       "url": "https://seekingalpha.com/market_currents.xml",            "weight": 6},
        {"name": "Reuters Business",    "url": "https://feeds.reuters.com/reuters/businessNews",          "weight": 9},
        {"name": "Reuters Economy",     "url": "https://feeds.reuters.com/reuters/economicsNews",         "weight": 9},
        {"name": "Business Insider",    "url": "https://markets.businessinsider.com/rss/news",            "weight": 7},
        {"name": "Barron's",            "url": "https://www.barrons.com/xml/rss/3_7551.xml",              "weight": 8},
    ]

    # Keywords de alta importancia
    HIGH_IMPACT = [
        "fed", "federal reserve", "rate", "inflation", "cpi", "gdp", "recession",
        "powell", "fomc", "treasury", "yield", "tariff", "trade war", "sanction",
        "earnings", "beats", "misses", "guidance", "outlook", "job", "unemployment",
        "crisis", "crash", "rally", "surge", "plunge", "collapse", "war", "china",
        "nvidia", "tsm", "apple", "microsoft", "amazon", "meta", "google", "tesla",
        "s&p", "nasdaq", "dow", "market", "stock", "bond", "oil", "gold", "bitcoin",
        "ecb", "boj", "bank of england", "opec", "debt", "deficit", "gdp", "pmi",
        "semiconductor", "ai", "artificial intelligence", "chip", "upgrade", "downgrade",
    ]

    def parse_feed(feed):
        items = []
        try:
            req = urllib.request.Request(
                feed["url"],
                headers={"User-Agent": "Mozilla/5.0 (compatible; NewsBot/1.0)"}
            )
            with urllib.request.urlopen(req, timeout=8) as resp:
                raw = resp.read(500000).decode("utf-8", errors="replace")

            root = ET.fromstring(raw)
            ns = {"atom": "http://www.w3.org/2005/Atom"}

            # Detectar formato Atom o RSS
            entries = root.findall(".//item") or root.findall(".//atom:entry", ns)

            for entry in entries[:12]:
                def g(tag):
                    el = entry.find(tag) or entry.find(f"atom:{tag}", ns)
                    return el.text.strip() if el is not None and el.text else ""

                title   = html.unescape(g("title"))
                link    = g("link") or g("url")
                summary = html.unescape(g("description") or g("summary") or g("content"))
                pub     = g("pubDate") or g("published") or g("updated")

                # Limpiar HTML del summary
                summary = re.sub(r"<[^>]+>", "", summary).strip()[:300]

                if not title or len(title) < 10:
                    continue

                # Calcular score
                text_low = (title + " " + summary).lower()
                score = feed["weight"]
                for kw in HIGH_IMPACT:
                    if kw in text_low:
                        score += 3

                # Parsear fecha
                pub_ts = None
                try:
                    from email.utils import parsedate_to_datetime
                    pub_ts = parsedate_to_datetime(pub).timestamp()
                except:
                    try:
                        pub_ts = dt.datetime.fromisoformat(pub.replace("Z", "+00:00")).timestamp()
                    except:
                        pub_ts = 0

                # Penalizar noticias viejas (>24h = -5, >6h = -2)
                age_h = (time.time() - pub_ts) / 3600 if pub_ts else 99
                if age_h > 24: score -= 8
                elif age_h > 12: score -= 4
                elif age_h > 6:  score -= 2
                elif age_h < 1:  score += 3

                items.append({
                    "title":   title,
                    "link":    link,
                    "summary": summary,
                    "source":  feed["name"],
                    "pub_ts":  pub_ts or 0,
                    "age_h":   round(age_h, 1),
                    "score":   score,
                })
        except Exception as e:
            print(f"  ⚠️  Feed error [{feed['name']}]: {e}", flush=True)
        return items

    # Jalar todos los feeds en paralelo
    import concurrent.futures
    all_items = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=12) as ex:
        futures = [ex.submit(parse_feed, f) for f in FEEDS]
        for fut in concurrent.futures.as_completed(futures, timeout=20):
            try:
                result = fut.result(timeout=1)
                all_items.extend(result)
            except concurrent.futures.TimeoutError:
                pass
            except Exception:
                pass

    # Deduplicar por título similar
    seen = []
    unique = []
    for item in all_items:
        title_key = re.sub(r"[^a-z0-9]", "", item["title"].lower())[:60]
        if not any(title_key in s or s in title_key for s in seen):
            seen.append(title_key)
            unique.append(item)

    # Ordenar por score desc
    unique.sort(key=lambda x: x["score"], reverse=True)

    # Formatear timestamps a string legible
    for item in unique:
        try:
            dt_obj = dt.datetime.fromtimestamp(item["pub_ts"])
            item["pub_str"] = dt_obj.strftime("%-I:%M %p · %b %-d")
        except Exception:
            item["pub_str"] = "Reciente"

    print(f"  📰 fetch_news: {len(unique)} artículos de {len(FEEDS)} feeds.", flush=True)
    return unique[:40]


# Cache y thread de noticias
news_cache = []
news_lock  = threading.Lock()
news_last_update = 0

# ── FRED Macro Scorecard Auto-Update ─────────────────────────────────────────
FRED_API_KEY = "f0343fed2845bcfed0cd622bcea196d7"
FRED_BASE    = "https://api.stlouisfed.org/fred/series/observations"

FRED_SERIES = {
    "CPI (Inflación)":    {"id": "CPIAUCSL",  "fmt": lambda v, p: (f"{v:.1f}%",  f"{p:.1f}%",  "YoY"),      "transform": "pch"},
    "NFP (Empleos)":      {"id": "PAYEMS",    "fmt": lambda v, p: (f"+{int(v)}K", f"+{int(p)}K", ""),         "transform": "chg"},
    "PIB (QoQ)":          {"id": "A191RL1Q225SBEA","fmt": lambda v, p: (f"{v:+.1f}%", f"{p:+.1f}%", "QoQ"), "transform": None},
    "ISM Manufacturing":  {"id": "MANEMP",    "fmt": None},   # fallback manual
    "Fed Funds Rate":     {"id": "FEDFUNDS",  "fmt": lambda v, p: (f"{v:.2f}%", f"{p:.2f}%", ""),            "transform": None},
    "Desempleo":          {"id": "UNRATE",    "fmt": lambda v, p: (f"{v:.1f}%", f"{p:.1f}%", ""),            "transform": None},
}


def _fred_fetch(series_id, transform=None, limit=2):
    """Descarga las últimas N observaciones de una serie FRED."""
    try:
        import urllib.request
        params = f"series_id={series_id}&api_key={FRED_API_KEY}&file_type=json&sort_order=desc&limit={limit}"
        if transform:
            params += f"&units={transform}"
        url = f"{FRED_BASE}?{params}"
        with urllib.request.urlopen(url, timeout=10) as r:
            data = json.loads(r.read())
        obs = [o for o in data.get("observations", []) if o["value"] != "."]
        return obs
    except Exception as e:
        print(f"  ⚠️  FRED error ({series_id}): {e}", flush=True)
        return []

def _trend(actual, prev):
    if actual > prev + 0.001: return "up"
    if actual < prev - 0.001: return "down"
    return "flat"

def _signal_cpi(v, p):
    if v < 2.0: return "positive"
    if v < 3.0: return "neutral"
    return "negative"

def _signal_nfp(v):
    if v > 200: return "positive"
    if v > 100: return "neutral"
    return "negative"

def _signal_pib(v):
    if v > 1.0: return "positive"
    if v >= 0:  return "neutral"
    return "negative"

def _signal_ism(v):
    if v > 50: return "positive"
    if v > 48: return "neutral"
    return "negative"

def _signal_fed(v, p):
    if v < p:   return "positive"
    if v > p:   return "negative"
    return "neutral"

def _signal_unemp(v):
    if v < 4.0: return "positive"
    if v < 5.0: return "neutral"
    return "negative"

def update_macro_scorecard():
    """Actualiza macro_scorecard.json con datos frescos de FRED. Corre cada hora."""
    sc_path = os.path.join(BASE_DIR, "macro_scorecard.json")

    # Cargar scorecard existente como fallback
    try:
        with open(sc_path) as f:
            scorecard = json.load(f)
    except:
        scorecard = {"updated": "", "indicators": []}

    # Construir índice por nombre para actualizar in-place
    idx = {ind["name"]: ind for ind in scorecard.get("indicators", [])}

    def update_ind(name, actual_str, prev_str, trend, signal, note):
        if name not in idx:
            idx[name] = {"name": name}
        idx[name].update({"actual": actual_str, "prev": prev_str,
                          "trend": trend, "signal": signal, "note": note})

    now_label = datetime.now().strftime("%b %Y")

    # ── CPI YoY ──────────────────────────────────────────────
    obs = _fred_fetch("CPIAUCSL", transform="pc1", limit=2)
    if len(obs) >= 2:
        v, p = float(obs[0]["value"]), float(obs[1]["value"])
        date_label = obs[0]["date"][:7]
        update_ind("CPI (Inflación)", f"{v:.1f}%", f"{p:.1f}%",
                   _trend(v, p), _signal_cpi(v, p), f"{date_label} · YoY")

    # ── NFP (cambio mensual en miles) ────────────────────────
    obs = _fred_fetch("PAYEMS", transform="chg", limit=2)
    if len(obs) >= 2:
        v, p = float(obs[0]["value"]), float(obs[1]["value"])
        date_label = obs[0]["date"][:7]
        sign_v = "+" if v >= 0 else ""
        sign_p = "+" if p >= 0 else ""
        update_ind("NFP (Empleos)", f"{sign_v}{int(v)}K", f"{sign_p}{int(p)}K",
                   _trend(v, p), _signal_nfp(v), date_label)

    # ── PIB QoQ ──────────────────────────────────────────────
    obs = _fred_fetch("A191RL1Q225SBEA", limit=2)
    if len(obs) >= 2:
        v, p = float(obs[0]["value"]), float(obs[1]["value"])
        date_label = obs[0]["date"][:7]
        update_ind("PIB (QoQ)", f"{v:+.1f}%", f"{p:+.1f}%",
                   _trend(v, p), _signal_pib(v), f"{date_label} · QoQ")

    # ── ISM Manufacturing — dato propietario, no disponible en FRED
    # Se mantiene el valor manual del scorecard existente sin tocar

    # ── Fed Funds Rate ───────────────────────────────────────
    obs = _fred_fetch("FEDFUNDS", limit=2)
    if len(obs) >= 2:
        v, p = float(obs[0]["value"]), float(obs[1]["value"])
        date_label = obs[0]["date"][:7]
        update_ind("Fed Funds Rate", f"{v:.2f}%", f"{p:.2f}%",
                   _trend(v, p), _signal_fed(v, p), date_label)

    # ── Desempleo ────────────────────────────────────────────
    obs = _fred_fetch("UNRATE", limit=2)
    if len(obs) >= 2:
        v, p = float(obs[0]["value"]), float(obs[1]["value"])
        date_label = obs[0]["date"][:7]
        update_ind("Desempleo", f"{v:.1f}%", f"{p:.1f}%",
                   _trend(v, p), _signal_unemp(v), date_label)

    # Guardar
    scorecard["updated"]    = now_label
    scorecard["indicators"] = list(idx.values())
    try:
        with open(sc_path, "w") as f:
            json.dump(scorecard, f, indent=2, ensure_ascii=False)
        print(f"  📊 Macro scorecard actualizado ({now_label})", flush=True)
    except Exception as e:
        print(f"  ⚠️  Error guardando scorecard: {e}", flush=True)

def update_macro_scorecard_background():
    """Actualiza el macro scorecard al arrancar y luego cada hora."""
    time.sleep(10)   # esperar arranque del servidor
    while True:
        try:
            update_macro_scorecard()
        except Exception as e:
            print(f"  ⚠️  Macro scorecard thread error: {e}", flush=True)
        time.sleep(3600)  # cada hora

def update_news_background():
    """Actualiza noticias cada 10 minutos."""
    global news_last_update
    while True:
        try:
            print(f"  📰 Actualizando noticias... {datetime.now().strftime('%H:%M:%S')}", flush=True)
            items = fetch_news()
            with news_lock:
                news_cache.clear()
                news_cache.extend(items)
                news_last_update = time.time()
            print(f"  ✅ {len(items)} noticias cargadas.", flush=True)
        except Exception as e:
            print(f"  ⚠️  Error noticias: {e}", flush=True)
        time.sleep(600)  # cada 10 minutos


# ── Performance cache (por portafolio) ───────────────────────────────────────
perf_cache       = {"miguel": {"data": None, "ts": 0}, "paulo": {"data": None, "ts": 0}, "akila": {"data": None, "ts": 0}}
perf_short_cache = {"miguel": {"data": None, "ts": 0}, "paulo": {"data": None, "ts": 0}, "akila": {"data": None, "ts": 0}}
cal_cache           = {"data": None, "ts": 0}
CAL_TTL             = 3600   # 1 hora

perf_lock        = threading.Lock()
PERF_TTL         = 300   # 5 minutos
PERF_SHORT_TTL   = 60    # 1 minuto (intraday se refresca más seguido)


def fetch_performance(portfolio="miguel"):
    """
    Calcula la curva de rendimiento histórica usando las unidades actuales
    de cada posición multiplicadas por su precio histórico de cierre.
    """
    import pandas as pd
    import math

    if portfolio == "akila":
        # Consolidado: sumar unidades de ambos portafolios
        miguel_pos = load_positions("miguel")
        paulo_pos  = load_positions("paulo")
        merged = {}
        for p in miguel_pos + paulo_pos:
            tk = p["ticker"]
            merged[tk] = merged.get(tk, 0.0) + p["units"]
        positions = [{"ticker": tk, "units": u} for tk, u in merged.items()]
    else:
        positions = load_positions(portfolio)

    if not positions:
        return {"dates": [], "portfolio_values": [], "spy_prices": []}

    port_tickers = list({p["ticker"] for p in positions})
    dl_tickers   = list(set(port_tickers + ["SPY"]))
    units_map    = {p["ticker"]: p["units"] for p in positions}

    print(f"  📈 [{portfolio}] Descargando histórico de rendimiento: {dl_tickers}", flush=True)

    try:
        raw = _yf_download_safe(dl_tickers, timeout=30, period="400d", progress=False, auto_adjust=True, session=_YF_SESSION)
    except Exception as e:
        print(f"  ⚠️  yf.download performance error: {e}", flush=True)
        return {"dates": [], "portfolio_values": [], "spy_prices": []}

    if raw is None or raw.empty:
        return {"dates": [], "portfolio_values": [], "spy_prices": []}

    try:
        if isinstance(raw.columns, pd.MultiIndex):
            closes = raw["Close"].copy()
        else:
            closes = raw[["Close"]].copy()
            closes.columns = dl_tickers
    except Exception as e:
        print(f"  ⚠️  closes extraction error: {e}", flush=True)
        return {"dates": [], "portfolio_values": [], "spy_prices": []}

    closes = closes.ffill().dropna(how="all")

    def clean_val(v):
        if v is None: return None
        try:
            f = float(v)
            return None if (math.isnan(f) or math.isinf(f)) else round(f, 4)
        except:
            return None

    port_values = []
    for i in range(len(closes)):
        total = 0.0
        for tk, units in units_map.items():
            if tk in closes.columns:
                price = closes.iloc[i][tk]
                try:
                    p = float(price)
                    if not math.isnan(p):
                        total += units * p
                except:
                    pass
        port_values.append(total)

    port_series = pd.Series(port_values, index=closes.index)
    spy_series  = closes["SPY"] if "SPY" in closes.columns else pd.Series(dtype=float)

    dates     = [d.strftime("%Y-%m-%d") for d in closes.index]
    port_vals = [clean_val(v) for v in port_series.tolist()]
    spy_vals  = [clean_val(v) for v in spy_series.tolist()] if not spy_series.empty else []

    print(f"  ✅ Performance: {len(dates)} días de datos.", flush=True)
    # ── Capital neto invertido por fecha ──
    txns = load_transactions(portfolio)
    # Ordenar transacciones por fecha
    txns_sorted = sorted(txns, key=lambda t: t.get('date',''))
    # Construir serie de capital neto acumulado por fecha del calendario
    capital_neto = []
    running = 0.0
    tx_idx = 0
    for d in closes.index:
        d_str = d.strftime("%Y-%m-%d")
        # Acumular todas las tx hasta esta fecha
        while tx_idx < len(txns_sorted):
            tx_date = txns_sorted[tx_idx].get('date','')
            if tx_date <= d_str:
                qty   = float(txns_sorted[tx_idx].get('quantity', 0))
                price = float(txns_sorted[tx_idx].get('price', 0))
                tx_type = txns_sorted[tx_idx].get('type','buy').lower()
                if tx_type == 'buy':
                    running += qty * price
                else:
                    running -= qty * price
                tx_idx += 1
            else:
                break
        capital_neto.append(round(running, 2))

    # P&L acumulado = MV - capital neto
    pnl_acum = []
    for mv, cap in zip(port_values, capital_neto):
        try:
            pnl_acum.append(round(float(mv) - float(cap), 2))
        except:
            pnl_acum.append(None)

    return {
        "dates":            dates,
        "portfolio_values": port_vals,
        "spy_prices":       spy_vals,
        "capital_neto":     capital_neto,
        "pnl_acumulado":    pnl_acum,
    }


def fetch_performance_short(portfolio="miguel"):
    """
    Descarga datos intraday (5 días, intervalo 5min) para calcular
    la curva de rendimiento intradía del portafolio vs SPY.
    """
    import pandas as pd
    import math

    if portfolio == "akila":
        miguel_pos = load_positions("miguel")
        paulo_pos  = load_positions("paulo")
        merged = {}
        for p in miguel_pos + paulo_pos:
            tk = p["ticker"]
            merged[tk] = merged.get(tk, 0.0) + p["units"]
        positions = [{"ticker": tk, "units": u} for tk, u in merged.items()]
    else:
        positions = load_positions(portfolio)

    if not positions:
        return {"timestamps": [], "portfolio_values": [], "spy_prices": []}

    port_tickers = [p["ticker"] for p in positions]
    units_map    = {p["ticker"]: float(p.get("units", 0)) for p in positions}
    dl_tickers   = list(set(port_tickers + ["SPY"]))

    print(f"  📈 [{portfolio}] Descargando intraday: {dl_tickers}", flush=True)
    try:
        raw = _yf_download_safe(dl_tickers, timeout=25, period="5d", interval="5m", progress=False, auto_adjust=True)
    except Exception as e:
        print(f"  ⚠️  yf intraday error: {e}", flush=True)
        return {"timestamps": [], "portfolio_values": [], "spy_prices": []}

    if raw is None or raw.empty:
        return {"timestamps": [], "portfolio_values": [], "spy_prices": []}

    # Extraer Close — mismo patrón robusto que fetch_performance
    try:
        if isinstance(raw.columns, pd.MultiIndex):
            closes = raw["Close"].copy()
        else:
            closes = raw[["Close"]].copy()
            closes.columns = dl_tickers
    except Exception as e:
        print(f"  ⚠️  intraday closes error: {e}", flush=True)
        return {"timestamps": [], "portfolio_values": [], "spy_prices": []}

    closes = closes.ffill().dropna(how="all")
    print(f"  📊 Intraday closes shape: {closes.shape}, columns: {list(closes.columns)}", flush=True)

    def clean_val(v):
        if v is None: return None
        try:
            f = float(v)
            return None if (math.isnan(f) or math.isinf(f)) else round(f, 4)
        except:
            return None

    # Calcular valor del portafolio por timestamp usando pandas (más eficiente)
    port_values = pd.Series(0.0, index=closes.index)
    for tk in port_tickers:
        if tk in closes.columns:
            port_values = port_values.add(closes[tk].ffill() * units_map.get(tk, 0), fill_value=0)
        else:
            print(f"  ⚠️  Ticker {tk} no encontrado en closes intraday", flush=True)

    spy_series = closes["SPY"] if "SPY" in closes.columns else pd.Series(dtype=float)

    timestamps = []
    port_vals  = []
    spy_vals   = []

    for idx in closes.index:
        pv = clean_val(port_values.get(idx))
        if pv is None or pv == 0:
            continue
        try:
            # Convertir timestamp con timezone a ISO string limpio
            if hasattr(idx, 'tz_convert'):
                ts = idx.tz_convert('America/New_York').isoformat()
            elif hasattr(idx, 'isoformat'):
                ts = idx.isoformat()
            else:
                ts = str(idx)
            # Quitar microsegundos si los hay
            ts = ts[:19] + ts[19:].replace('.000000', '')
        except:
            ts = str(idx)
        sv = clean_val(spy_series.get(idx)) if not spy_series.empty else None
        timestamps.append(ts)
        port_vals.append(pv)
        spy_vals.append(sv)

    print(f"  ✅ Performance intraday: {len(timestamps)} puntos.", flush=True)
    return {
        "timestamps":       timestamps,
        "portfolio_values": port_vals,
        "spy_prices":       spy_vals,
    }


def _frontier_montecarlo(returns_df, tickers, risk_free=0.0525, trading_days=252, n_sims=80000):
    """
    Monte Carlo para la frontera eficiente (numpy puro).
    Usa 80k simulaciones Dirichlet y aplica cummax para curva suave y monótona.
    """
    import numpy as np
    n = len(tickers)
    if n < 2:
        return None

    mu  = returns_df[tickers].mean().values * trading_days
    cov = returns_df[tickers].cov().values  * trading_days

    # Generar pesos aleatorios (Dirichlet: suman 1, todos >= 0)
    W = np.random.dirichlet(np.ones(n), size=n_sims)
    rets   = W @ mu
    vars_  = np.einsum('ij,jk,ik->i', W, cov, W)
    vols   = np.sqrt(np.maximum(vars_, 0))
    sharps = np.where(vols > 1e-9, (rets - risk_free) / vols, -np.inf)

    idx_ms = int(np.argmax(sharps))
    idx_mv = int(np.argmin(vols))

    # Frontera: ordenar por vol → cummax → suavizado gaussiano → monotonía
    order = np.argsort(vols)
    sv = vols[order]
    sr = rets[order]
    cummax_r = np.maximum.accumulate(sr)

    # Suavizado gaussiano (kernel puro numpy, sin scipy)
    sigma = max(len(sv) // 25, 10)
    x_k = np.arange(-3 * sigma, 3 * sigma + 1)
    kernel = np.exp(-x_k**2 / (2 * sigma**2))
    kernel /= kernel.sum()
    padded = np.pad(cummax_r, (3 * sigma, 3 * sigma), mode='edge')
    smooth_all = np.convolve(padded, kernel, mode='valid')[:len(sv)]

    # cummax final para garantizar que la curva solo sube
    smooth_all = np.maximum.accumulate(smooth_all)

    # Muestrear 60 puntos uniformes (hasta pct 88 para no incluir la cola ruidosa)
    vol_max_eff = float(np.percentile(sv, 88))
    smooth_v = np.linspace(float(sv[0]), vol_max_eff, 60)
    smooth_r = np.interp(smooth_v, sv, smooth_all)

    frontier = []
    for v, r in zip(smooth_v, smooth_r):
        s = (float(r) - risk_free) / float(v) if float(v) > 1e-9 else None
        frontier.append({"vol":    round(float(v) * 100, 2),
                         "ret":    round(float(r) * 100, 2),
                         "sharpe": round(s, 3) if s else None})

    def fmt_w(idx):
        return {tk: round(float(W[idx, j]) * 100, 1) for j, tk in enumerate(tickers)}

    return {
        "frontier":   frontier,
        "method":     "montecarlo",
        "n_sims":     n_sims,
        "assets":     [{"ticker": tk,
                        "vol": round(float(np.sqrt(cov[i, i])) * 100, 2),
                        "ret": round(float(mu[i]) * 100, 2)}
                       for i, tk in enumerate(tickers)],
        "max_sharpe": {"vol":     round(float(vols[idx_ms]) * 100, 2),
                       "ret":     round(float(rets[idx_ms]) * 100, 2),
                       "sharpe":  round(float(sharps[idx_ms]), 3),
                       "weights": fmt_w(idx_ms)},
        "min_vol":    {"vol":     round(float(vols[idx_mv]) * 100, 2),
                       "ret":     round(float(rets[idx_mv]) * 100, 2),
                       "sharpe":  round(float(sharps[idx_mv]), 3)
                                  if vols[idx_mv] > 1e-9 else None,
                       "weights": fmt_w(idx_mv)},
    }


def compute_efficient_frontier(returns_df, tickers, risk_free=0.0525, trading_days=252, n_points=60):
    """
    Híbrido: Monte Carlo para pesos (Max Sharpe / Min Vol distribuidos) +
             scipy SLSQP para la CURVA (60 puntos exactos, matemáticamente suave).
    Si scipy no está disponible, cae a Monte Carlo puro.
    """
    import numpy as np
    n = len(tickers)
    if n < 2:
        return None

    # ── Paso 1: Monte Carlo → pesos distribuidos de Max Sharpe y Min Vol ─────
    mc = _frontier_montecarlo(returns_df, tickers, risk_free, trading_days)
    print(f"  [frontier] Monte Carlo OK — pesos obtenidos", flush=True)

    # ── Paso 2: scipy → curva exacta (solo geometría, no pesos) ──────────────
    mu  = returns_df[tickers].mean().values * trading_days
    cov = returns_df[tickers].cov().values  * trading_days

    try:
        from scipy.optimize import minimize

        def port_ret(w): return float(w @ mu)
        def port_vol(w): return float(np.sqrt(w @ cov @ w))

        cons = [{'type': 'eq', 'fun': lambda w: np.sum(w) - 1}]
        bounds = [(0.0, 1.0)] * n
        w0 = np.ones(n) / n

        # Min Vol → ancla izquierda de la curva
        res_mv = minimize(lambda w: float(w @ cov @ w), w0, method='SLSQP',
                          bounds=bounds, constraints=cons,
                          options={'ftol': 1e-12, 'maxiter': 1000})
        w_mv = res_mv.x if res_mv.success else w0

        ret_lo = port_ret(w_mv)
        ret_hi = float(mu.max()) * 0.98
        if ret_hi <= ret_lo:
            ret_hi = ret_lo * 1.5

        # Trazar la curva: para cada target-return, minimizar varianza
        frontier = []
        for t in np.linspace(ret_lo, ret_hi, n_points):
            c_t = cons + [{'type': 'eq', 'fun': lambda w, _t=t: port_ret(w) - _t}]
            r = minimize(lambda w: float(w @ cov @ w), w0, method='SLSQP',
                         bounds=bounds, constraints=c_t,
                         options={'ftol': 1e-12, 'maxiter': 1000})
            if r.success:
                pv = port_vol(r.x)
                frontier.append({
                    "vol":    round(pv * 100, 2),
                    "ret":    round(port_ret(r.x) * 100, 2),
                    "sharpe": round((port_ret(r.x) - risk_free) / pv, 3) if pv > 1e-9 else None
                })

        # Fusionar: curva scipy + pesos MC
        if mc:
            mc['frontier'] = frontier
            mc['method']   = 'hybrid'
        else:
            mc = {"frontier": frontier, "method": "scipy",
                  "assets": [{"ticker": tk,
                               "vol": round(float(np.sqrt(cov[i,i]))*100,2),
                               "ret": round(float(mu[i])*100,2)}
                              for i, tk in enumerate(tickers)]}

        print(f"  [frontier] híbrido OK — {len(frontier)} puntos curva + pesos MC", flush=True)
        return mc

    except ImportError:
        print("  [frontier] scipy no disponible — usando Monte Carlo puro", flush=True)
    except Exception as e:
        print(f"  [frontier] scipy error ({e}) — usando Monte Carlo puro", flush=True)

    # Fallback: Monte Carlo puro
    return mc


def compute_risk_metrics(portfolio="miguel"):
    """Calcula métricas de riesgo del portafolio usando historial de precios 2Y."""
    import math
    import numpy as np

    BENCHMARK      = "SPY"
    RISK_FREE_RATE = 0.0525
    TRADING_DAYS   = 252
    MIN_HISTORY    = 120   # días mínimos para incluir un ticker

    # ── Obtener posiciones y pesos ─────────────────────────────────────────────
    if portfolio == "akila":
        miguel_pos = compute_positions_from_transactions(load_transactions("miguel"))
        paulo_pos  = compute_positions_from_transactions(load_transactions("paulo"))
        # Merge por ticker sumando units
        merged = {}
        for p in miguel_pos + paulo_pos:
            tk = p["ticker"]
            if tk not in merged:
                merged[tk] = dict(p)
            else:
                merged[tk]["units"] += p["units"]
        positions = list(merged.values())
    else:
        positions = compute_positions_from_transactions(load_transactions(portfolio))

    tickers = [p["ticker"] for p in positions]
    if not tickers:
        return {"error": "No hay posiciones"}

    # ── Fetch precios ──────────────────────────────────────────────────────────
    print(f"  [risk] Fetching precios para {portfolio}: {tickers}", flush=True)
    all_tickers = tickers + [BENCHMARK]

    # Intento 1: descarga en bloque (robusto a distintas versiones de yfinance)
    import pandas as pd
    raw = None
    try:
        _dl = _yf_download_safe(all_tickers, timeout=40, period="2y", auto_adjust=True, progress=False, session=_YF_SESSION)
        if _dl is not None and not _dl.empty:
            if isinstance(_dl.columns, pd.MultiIndex):
                lvl0 = _dl.columns.get_level_values(0)
                lvl1 = _dl.columns.get_level_values(1)
                if "Close" in lvl0:
                    raw = _dl["Close"]
                elif "Close" in lvl1:
                    raw = _dl.xs("Close", axis=1, level=1)
            else:
                # Single ticker — columnas planas
                raw = _dl[["Close"]].rename(columns={"Close": all_tickers[0]})
        if raw is None or raw.empty:
            raw = None
    except Exception as e:
        print(f"  [risk] Bulk download falló ({e}), intentando ticker a ticker…", flush=True)

    # Intento 2 (fallback): uno a uno
    if raw is None:
        import time as _time
        frames = {}
        for tk in all_tickers:
            try:
                t = yf.Ticker(tk)
                h = t.history(period="2y", auto_adjust=True)
                if not h.empty:
                    frames[tk] = h["Close"]
                _time.sleep(0.3)
            except Exception as ex:
                print(f"  [risk] No se pudo obtener {tk}: {ex}", flush=True)
        if not frames:
            return {"error": "No se pudo obtener datos de precios (401 / sin datos)"}
        raw = pd.DataFrame(frames)

    if raw is None or raw.empty:
        return {"error": "No se pudo obtener datos de precios"}

    # ── Filtrar tickers con suficiente historia y alinear ─────────────────────
    valid_tickers = [tk for tk in tickers
                     if tk in raw.columns and raw[tk].dropna().shape[0] >= MIN_HISTORY]
    if not valid_tickers:
        return {"error": "Ningún ticker tiene suficiente historia"}

    # ffill cubre holidays/calendarios distintos (ASML, EWY, EWZ, etc.); dropna solo
    # elimina fechas previas a cualquier dato
    bench_cols = [BENCHMARK] if BENCHMARK in raw.columns else []
    aligned = raw[valid_tickers + bench_cols].ffill().dropna()
    if len(aligned) < 60:
        return {"error": "Datos insuficientes para calcular métricas"}

    # ── Pesos actuales (renormalizados a tickers válidos) ─────────────────────
    current_prices = aligned.iloc[-1]
    pos_map = {p["ticker"]: p["units"] for p in positions}
    mkt_vals = {tk: pos_map.get(tk, 0) * current_prices.get(tk, 0) for tk in valid_tickers}
    total_mkt = sum(mkt_vals.values())
    weights = {tk: v / total_mkt for tk, v in mkt_vals.items()} if total_mkt > 0 else {}

    # ── Retornos diarios ───────────────────────────────────────────────────────
    returns     = aligned[valid_tickers].pct_change().dropna()
    spy_ret     = aligned[BENCHMARK].pct_change().dropna()
    w_arr       = np.array([weights.get(tk, 0) for tk in valid_tickers])
    port_ret    = returns[valid_tickers].values @ w_arr

    rf_daily    = RISK_FREE_RATE / TRADING_DAYS

    # Retorno geométrico verdadero: (valor_final/valor_inicial)^(252/N) - 1
    # Más preciso que (1+mean_diario)^252 para portafolios de alta volatilidad
    n_days      = len(port_ret)
    cum_ret     = float(np.prod(1 + port_ret))
    ann_return  = float(cum_ret ** (TRADING_DAYS / n_days) - 1) if n_days > 0 else 0.0
    ann_vol     = float(port_ret.std() * np.sqrt(TRADING_DAYS))

    # Sharpe
    sharpe = float((ann_return - RISK_FREE_RATE) / ann_vol) if ann_vol > 0 else None

    # Sortino — downside deviation vs rf diario como MAR
    downside_ret = port_ret[port_ret < rf_daily]
    downside_vol = float(downside_ret.std() * np.sqrt(TRADING_DAYS)) if len(downside_ret) > 0 else None
    sortino = float((ann_return - RISK_FREE_RATE) / downside_vol) if downside_vol else None

    # Beta y Alpha
    cov_mat  = np.cov(np.column_stack([port_ret, spy_ret]).T)
    beta     = float(cov_mat[0, 1] / cov_mat[1, 1]) if cov_mat[1, 1] > 0 else None
    n_spy    = len(spy_ret)
    spy_ann  = float(np.prod(1 + spy_ret.values) ** (TRADING_DAYS / n_spy) - 1) if n_spy > 0 else 0.0
    alpha    = float(ann_return - (RISK_FREE_RATE + beta * (spy_ann - RISK_FREE_RATE))) if beta else None

    # Treynor y Calmar
    treynor  = float((ann_return - RISK_FREE_RATE) / beta) if beta else None

    # Max Drawdown
    cum          = (1 + port_ret).cumprod()
    rolling_max  = np.maximum.accumulate(cum)
    drawdowns    = (cum - rolling_max) / rolling_max
    max_dd       = float(drawdowns.min())
    calmar       = float(ann_return / abs(max_dd)) if max_dd != 0 else None

    # VaR y CVaR histórico
    var_95   = float(np.percentile(port_ret, 5))
    cvar_95  = float(port_ret[port_ret <= var_95].mean())

    # Tracking Error e Information Ratio
    tracking_error = float((port_ret - spy_ret).std() * np.sqrt(TRADING_DAYS))
    info_ratio     = float(alpha / tracking_error) if (alpha and tracking_error > 0) else None

    # Correlación entre posiciones
    corr = returns[valid_tickers].corr().round(3)
    corr_matrix = {tk: {tk2: round(float(corr.loc[tk, tk2]), 3)
                        for tk2 in valid_tickers} for tk in valid_tickers}

    # Underwater chart (drawdown diario con fechas)
    underwater_dates  = [d.strftime("%Y-%m-%d") for d in aligned.index[1:]]
    underwater_values = [round(float(v) * 100, 4) for v in drawdowns.tolist()]

    # Retornos históricos del portafolio (para histograma)
    daily_returns_list = [round(float(v) * 100, 4) for v in port_ret.tolist()]
    return_dates       = [d.strftime("%Y-%m-%d") for d in returns.index]

    # Beta individual por ticker
    betas_individual = {}
    for tk in valid_tickers:
        try:
            cov_tk = np.cov(np.column_stack([returns[tk].values, spy_ret.values]).T)
            betas_individual[tk] = round(float(cov_tk[0, 1] / cov_tk[1, 1]), 3)
        except:
            betas_individual[tk] = None

    def safe_round(v, d=3):
        if v is None: return None
        if isinstance(v, float) and (math.isnan(v) or math.isinf(v)): return None
        return round(v, d)

    # ── Efficient Frontier ────────────────────────────────────────────────────
    print(f"  [risk] Calculando frontera eficiente para {portfolio}…", flush=True)
    frontier_data = compute_efficient_frontier(returns, valid_tickers, RISK_FREE_RATE, TRADING_DAYS)

    return {
        "portfolio":        portfolio,
        "tickers":          valid_tickers,
        "weights":          {tk: round(weights.get(tk, 0) * 100, 2) for tk in valid_tickers},
        "n_days":           len(returns),
        "date_range":       [str(aligned.index[0].date()), str(aligned.index[-1].date())],
        # Retorno y riesgo
        "ann_return":       safe_round(ann_return * 100, 2),
        "ann_vol":          safe_round(ann_vol * 100, 2),
        "spy_ann_return":   safe_round(spy_ann * 100, 2),
        # Ratios
        "sharpe":           safe_round(sharpe),
        "sortino":          safe_round(sortino),
        "calmar":           safe_round(calmar),
        "treynor":          safe_round(treynor, 4),
        "beta":             safe_round(beta),
        "alpha":            safe_round(alpha * 100, 2) if alpha else None,
        "info_ratio":       safe_round(info_ratio),
        "tracking_error":   safe_round(tracking_error * 100, 2),
        # Drawdown
        "max_drawdown":     safe_round(max_dd * 100, 2),
        # VaR
        "var_95":           safe_round(var_95 * 100, 2),
        "cvar_95":          safe_round(cvar_95 * 100, 2),
        # Series para charts
        "underwater_dates":  underwater_dates,
        "underwater_values": underwater_values,
        "return_dates":      return_dates,
        "daily_returns":     daily_returns_list,
        # Matrices
        "corr_matrix":      corr_matrix,
        "betas_individual": betas_individual,
        # Efficient Frontier
        "frontier":         frontier_data,
    }


def _compute_risk_background(portfolio):
    """Corre compute_risk_metrics en background y guarda en caché al terminar."""
    try:
        data = compute_risk_metrics(portfolio)
        with risk_cache_lock:
            risk_cache[portfolio] = {"data": data, "ts": time.time()}
        print(f"  ✅ Risk calculado en background: {portfolio}", flush=True)
    except Exception as e:
        print(f"  ⚠️  Risk background error [{portfolio}]: {e}", flush=True)
    finally:
        _risk_computing.discard(portfolio)

def get_risk_cached(portfolio="miguel"):
    """
    Devuelve métricas de riesgo de forma no-bloqueante.
    - Si hay caché fresco: devuelve inmediatamente.
    - Si caché existe pero está stale: devuelve stale + dispara recompute en background.
    - Si no hay caché: dispara compute en background y responde {"computing": true}.
    """
    import time
    with risk_cache_lock:
        entry = risk_cache.get(portfolio)
        has_fresh = entry and (time.time() - entry["ts"]) < RISK_TTL
        has_stale = entry and not has_fresh

    if has_fresh:
        print(f"  📊 Risk cache HIT: {portfolio}", flush=True)
        return entry["data"]

    # Disparar cálculo en background si no está ya corriendo
    if portfolio not in _risk_computing:
        _risk_computing.add(portfolio)
        t = threading.Thread(target=_compute_risk_background, args=(portfolio,), daemon=True)
        t.start()
        print(f"  📊 Risk cache MISS: {portfolio} — calculando en background…", flush=True)

    if has_stale:
        # Devolver datos viejos mientras se recalcula
        stale_data = dict(entry["data"])
        stale_data["_stale"] = True
        return stale_data

    # Sin caché — informar al cliente que espere
    return {"computing": True, "portfolio": portfolio}


def fetch_tearsheet(ticker, portfolio="miguel"):
    """Jala todos los datos necesarios para el tearsheet de una empresa."""
    import math
    import datetime as dt

    def safe(val):
        if val is None: return None
        if isinstance(val, float) and (math.isnan(val) or math.isinf(val)): return None
        return val

    try:
        t = yf.Ticker(ticker)

        # ── Info de empresa (fallback a dict vacío si falla) ──
        try:
            info = t.info or {}
        except Exception as e_info:
            print(f"  [tearsheet] info error {ticker}: {e_info}", flush=True)
            info = {}

        # ── Histórico de precios 2y ───────────────────────────
        hist_rows = []
        try:
            hist_2y = t.history(period="2y", interval="1d")
            if not hist_2y.empty:
                for date, row in hist_2y.iterrows():
                    try:
                        hist_rows.append({
                            "date":   date.strftime("%Y-%m-%d"),
                            "open":   safe(float(row["Open"])),
                            "high":   safe(float(row["High"])),
                            "low":    safe(float(row["Low"])),
                            "close":  safe(float(row["Close"])),
                            "volume": safe(int(row["Volume"])),
                        })
                    except:
                        pass
        except Exception as e_hist:
            print(f"  [tearsheet] history error {ticker}: {e_hist}", flush=True)

        # ── Precio y cambio ───────────────────────────────────
        price    = safe(info.get("regularMarketPrice") or info.get("currentPrice"))
        prev     = safe(info.get("regularMarketPreviousClose") or info.get("previousClose"))
        # Si info no tiene precio, intentar desde el último row del historial
        if price is None and hist_rows:
            price = hist_rows[-1]["close"]
        if prev is None and len(hist_rows) >= 2:
            prev = hist_rows[-2]["close"]
        chg_pct  = (price - prev) / prev if (price and prev) else None

        # ── YTD return ────────────────────────────────────────
        ytd_ret = None
        try:
            start_ytd = dt.datetime(dt.datetime.now().year, 1, 1, tzinfo=timezone.utc)
            hist_ytd  = t.history(start=start_ytd, interval="1d")
            if not hist_ytd.empty and len(hist_ytd) >= 2:
                ytd_ret = (float(hist_ytd["Close"].iloc[-1]) - float(hist_ytd["Close"].iloc[0])) / float(hist_ytd["Close"].iloc[0])
        except:
            pass

        # ── Posición del portafolio ───────────────────────────
        # Para AKILA consolidado, buscar en ambos portafolios
        if portfolio == "akila":
            positions_to_search = (compute_positions_from_transactions(load_transactions("miguel")) +
                                   compute_positions_from_transactions(load_transactions("paulo")))
        else:
            positions_to_search = compute_positions_from_transactions(load_transactions(portfolio))
        analysis_data = load_analysis(portfolio if portfolio != "akila" else "miguel")

        pos_data = None
        for pos in positions_to_search:
            if pos["ticker"] == ticker:
                cost    = pos["units"] * pos["avg_price"]
                mkt_val = pos["units"] * price if price else cost
                pnl_usd = mkt_val - cost
                pnl_pct = pnl_usd / cost if cost else 0
                ana     = analysis_data.get(ticker, {})
                pos_data = {
                    "units":      pos["units"],
                    "avg_price":  pos["avg_price"],
                    "cost":       cost,
                    "mkt_val":    mkt_val,
                    "pnl_usd":    pnl_usd,
                    "pnl_pct":    pnl_pct,
                    "bull_pt":    ana.get("bull_pt"),
                    "base_pt":    ana.get("base_pt"),
                    "conviction": ana.get("conviction"),
                    "next_cat":   ana.get("next_cat"),
                    "tesis":      ana.get("tesis"),
                    "pt_consenso":safe(info.get("targetMeanPrice")),
                    "recomendacion": (info.get("recommendationKey") or "").upper(),
                    "num_analistas": info.get("numberOfAnalystOpinions"),
                }
                break

        # ── EPS trimestral ────────────────────────────────────
        eps_quarterly = []
        try:
            fi = t.quarterly_income_stmt
            if fi is not None and not fi.empty and "Diluted EPS" in fi.index:
                eps_row = fi.loc["Diluted EPS"]
                for col in sorted(eps_row.index):
                    val = eps_row[col]
                    if val is not None and not (isinstance(val, float) and (math.isnan(val) or math.isinf(val))):
                        eps_quarterly.append({"date": str(col.date()), "eps": float(val)})
        except Exception as e_eps:
            print(f"  [tearsheet] EPS quarterly error: {e_eps}", flush=True)

        result = {
            "ticker":        ticker,
            "name":          info.get("longName") or info.get("shortName", ticker),
            "sector":        info.get("sector", "—"),
            "industry":      info.get("industry", "—"),
            "country":       info.get("country", "—"),
            "description":   info.get("longBusinessSummary", ""),
            "website":       info.get("website", ""),
            "employees":     info.get("fullTimeEmployees"),
            "exchange":      info.get("exchange", ""),
            # Precio
            "price":         price,
            "prev_close":    prev,
            "chg_pct":       safe(chg_pct),
            "ytd_ret":       safe(ytd_ret),
            "open":          safe(info.get("regularMarketOpen") or info.get("open")),
            "day_high":      safe(info.get("dayHigh") or info.get("regularMarketDayHigh")),
            "day_low":       safe(info.get("dayLow") or info.get("regularMarketDayLow")),
            "high_52w":      safe(info.get("fiftyTwoWeekHigh")),
            "low_52w":       safe(info.get("fiftyTwoWeekLow")),
            "volume":        info.get("regularMarketVolume") or info.get("volume"),
            "avg_volume":    info.get("averageVolume"),
            # Valuación
            "mkt_cap":       safe(info.get("marketCap")),
            "ev":            safe(info.get("enterpriseValue")),
            "pe":            safe(info.get("trailingPE")),
            "fwd_pe":        safe(info.get("forwardPE")),
            "pb":            safe(info.get("priceToBook")),
            "ps":            safe(info.get("priceToSalesTrailing12Months")),
            "ev_ebitda":     safe(info.get("enterpriseToEbitda")),
            "ev_revenue":    safe(info.get("enterpriseToRevenue")),
            # Fundamentales
            "revenue":       safe(info.get("totalRevenue")),
            "revenue_growth":safe(info.get("revenueGrowth")),
            "gross_margin":  safe(info.get("grossMargins")),
            "op_margin":     safe(info.get("operatingMargins")),
            "net_margin":    safe(info.get("profitMargins")),
            "eps_ttm":       safe(info.get("trailingEps")),
            "eps_fwd":       safe(info.get("forwardEps")),
            "eps_growth":    safe(info.get("earningsGrowth")),
            "fcf":           safe(info.get("freeCashflow")),
            "debt_equity":   safe(info.get("debtToEquity")),
            "roe":           safe(info.get("returnOnEquity")),
            "roa":           safe(info.get("returnOnAssets")),
            "beta":          safe(info.get("beta")),
            "dividend_yield":safe(info.get("dividendYield")),
            # Posición
            "position":      pos_data,
            # Histórico
            "history":       hist_rows,
            # EPS trimestral histórico
            "eps_quarterly": eps_quarterly,
        }
        return result

    except Exception as e:
        print(f"  ⚠️  Tearsheet error {ticker}: {e}", flush=True)
        return {"error": str(e)}


# ══════════════════════════════════════════════════════════════
#  GENERADOR DE PDF — AKILA CAPITAL PARTNERS
# ══════════════════════════════════════════════════════════════

def generate_portfolio_pdf():
    """
    Equity-research + wealth-management client report style.
    Una sola hoja A4. Resumen full-width arriba, dos columnas 50/50.
    Incluye: performance QTD/YTD/ITD vs SPY, holdings, market commentary,
    activity summary, allocation donut, métricas clave.
    """
    import io, math
    from datetime import datetime, date
    from collections import defaultdict
    import yfinance as yf
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import matplotlib.dates as mdates
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY, TA_RIGHT
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
        Table, TableStyle, Image, HRFlowable)

    # ── Medidas ───────────────────────────────────────────────────
    W, H   = A4
    MX     = 1.4 * cm
    MY     = 1.1 * cm
    HDR    = 22
    FTR    = 14
    CW     = W - 2 * MX
    GAP    = 7
    COL    = (CW - GAP) / 2

    # ── Colores ───────────────────────────────────────────────────
    NAVY  = colors.HexColor('#0D2B4E')
    NAVY2 = colors.HexColor('#1A4A7A')
    GOLD  = colors.HexColor('#C9A248')
    WHITE = colors.white
    LGRAY = colors.HexColor('#F4F7FB')
    MGRAY = colors.HexColor('#B0BEC5')
    DGRAY = colors.HexColor('#5A5A5A')
    GREEN = colors.HexColor('#1A6B3A')
    RED   = colors.HexColor('#B52B27')
    BG    = colors.HexColor('#EEF3F9')
    LGRN  = colors.HexColor('#E8F8EE')
    LRED  = colors.HexColor('#FDECEA')
    LGOLD = colors.HexColor('#FDF8EC')

    # ── Estilos ───────────────────────────────────────────────────
    def S(n, **kw): return ParagraphStyle(n, **kw)

    body   = S('bd', fontSize=7.4, textColor=colors.HexColor('#1A1F24'),
               fontName='Helvetica', leading=11, spaceAfter=2, alignment=TA_JUSTIFY)
    note   = S('nt', fontSize=6.2, textColor=DGRAY, fontName='Helvetica',
               leading=8.5, spaceAfter=2, alignment=TA_JUSTIFY)
    blt    = S('bl', fontSize=7, textColor=colors.HexColor('#1A1F24'),
               fontName='Helvetica', leading=10, leftIndent=8,
               firstLineIndent=-6, spaceAfter=1)
    th     = S('th', fontSize=6.5, textColor=WHITE, fontName='Helvetica-Bold',
               alignment=TA_CENTER, leading=8)
    tc     = S('tc', fontSize=6.5, textColor=colors.black, fontName='Helvetica',
               alignment=TA_CENTER, leading=8)
    tl     = S('tl', fontSize=6.5, textColor=colors.black, fontName='Helvetica',
               alignment=TA_LEFT, leading=8)
    tb     = S('tb', fontSize=6.5, textColor=NAVY, fontName='Helvetica-Bold',
               alignment=TA_CENTER, leading=8)
    tpos   = S('tp', fontSize=6.5, textColor=GREEN, fontName='Helvetica-Bold',
               alignment=TA_RIGHT, leading=8)
    tneg   = S('tn', fontSize=6.5, textColor=RED, fontName='Helvetica-Bold',
               alignment=TA_RIGHT, leading=8)
    tc_r   = S('tr', fontSize=6.5, textColor=colors.black, fontName='Helvetica',
               alignment=TA_RIGHT, leading=8)
    card_l = S('cl', fontSize=5.8, textColor=DGRAY, fontName='Helvetica',
               alignment=TA_CENTER, leading=7.5)
    card_v = S('cv', fontSize=10, textColor=NAVY, fontName='Helvetica-Bold',
               alignment=TA_CENTER, leading=12)
    card_s = S('cs', fontSize=5.8, textColor=DGRAY, fontName='Helvetica',
               alignment=TA_CENTER, leading=7.5)
    sec_t  = S('st', fontSize=6.8, textColor=WHITE, fontName='Helvetica-Bold',
               leading=9)
    mini_t = S('mt', fontSize=6.5, textColor=WHITE, fontName='Helvetica-Bold',
               leading=8)

    def sp(h=3): return Spacer(1, h)

    def sec_bar(title, w=CW):
        t = Table([[Paragraph(title, sec_t)]], colWidths=[w])
        t.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,-1),NAVY),
            ('LEFTPADDING',(0,0),(-1,-1),6),
            ('TOPPADDING',(0,0),(-1,-1),3),
            ('BOTTOMPADDING',(0,0),(-1,-1),3),
        ]))
        return t

    def mini_sec(title, w=COL):
        t = Table([[Paragraph(title, mini_t)]], colWidths=[w])
        t.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,-1),NAVY2),
            ('LEFTPADDING',(0,0),(-1,-1),5),
            ('TOPPADDING',(0,0),(-1,-1),2.5),
            ('BOTTOMPADDING',(0,0),(-1,-1),2.5),
        ]))
        return t

    def make_tbl(headers, rows, cw, tiny=False, bold_last=False):
        sz = 5.8 if tiny else 6.5
        th2 = S('th2', fontSize=sz, textColor=WHITE, fontName='Helvetica-Bold',
                alignment=TA_CENTER, leading=sz+1.5)
        data = [[Paragraph(h, th2) for h in headers]] + rows
        t = Table(data, colWidths=cw, repeatRows=1)
        style = [
            ('BACKGROUND',(0,0),(-1,0),NAVY),
            ('ROWBACKGROUNDS',(0,1),(-1,-1),[WHITE,LGRAY]),
            ('GRID',(0,0),(-1,-1),0.25,MGRAY),
            ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
            ('TOPPADDING',(0,0),(-1,-1),1.5),
            ('BOTTOMPADDING',(0,0),(-1,-1),1.5),
            ('LEFTPADDING',(0,0),(-1,-1),3),
            ('RIGHTPADDING',(0,0),(-1,-1),3),
        ]
        if bold_last:
            style += [
                ('BACKGROUND',(0,-1),(-1,-1),BG),
                ('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),
            ]
        t.setStyle(TableStyle(style))
        return t

    def callout(text, w=COL, bg=BG, border=GOLD):
        t = Table([[Paragraph(text, S('cb', fontSize=6.8, textColor=NAVY,
                   fontName='Helvetica-Bold', leading=10))]], colWidths=[w])
        t.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,-1),bg),
            ('LEFTPADDING',(0,0),(-1,-1),6),
            ('RIGHTPADDING',(0,0),(-1,-1),6),
            ('TOPPADDING',(0,0),(-1,-1),4),
            ('BOTTOMPADDING',(0,0),(-1,-1),4),
            ('LINEBEFORE',(0,0),(-1,-1),2.5,border),
        ]))
        return t

    def hr(w=CW, c=GOLD, th=0.8):
        return HRFlowable(width=w, thickness=th, color=c, spaceAfter=3, spaceBefore=1)

    def pct_cell(v, style_pos=tpos, style_neg=tneg):
        if v is None: return Paragraph('—', tc)
        s = ('+' if v >= 0 else '') + f'{v*100:.2f}%'
        return Paragraph(s, style_pos if v >= 0 else style_neg)

    # ── Datos ─────────────────────────────────────────────────────
    positions = load_positions()
    txns      = load_transactions()
    tickers   = [p['ticker'] for p in positions]
    macro_tk  = ['SPY','^VIX','^TNX','MXN=X','GC=F','CL=F']
    all_tk    = list(set(tickers + macro_tk))

    prices_cur = {}; prices_prev = {}
    try:
        raw   = _yf_download_safe(all_tk, timeout=25, period='5d', auto_adjust=True, progress=False)
        close = raw['Close'] if 'Close' in raw.columns else raw
        for tk in all_tk:
            if tk in close.columns:
                s = close[tk].dropna()
                if len(s) >= 2:
                    prices_cur[tk]  = float(s.iloc[-1])
                    prices_prev[tk] = float(s.iloc[-2])
                elif len(s) == 1:
                    prices_cur[tk] = prices_prev[tk] = float(s.iloc[-1])
    except Exception as e:
        print(f'  ⚠ Prices: {e}', flush=True)

    def day_chg(tk):
        c = prices_cur.get(tk); p = prices_prev.get(tk)
        return (c - p) / p if c and p and p else None

    # Enriquecer posiciones
    total_cost = total_value = 0.0
    enriched = []
    for pos in positions:
        tk     = pos['ticker']
        units  = pos.get('units', 0) or 0
        avg_px = pos.get('avg_price', 0) or 0
        cur_px = prices_cur.get(tk)
        cost   = units * avg_px
        value  = units * cur_px if cur_px else cost
        pnl    = value - cost
        pnl_pct= pnl / cost if cost else 0
        total_cost  += cost
        total_value += value
        enriched.append(dict(ticker=tk, name=pos.get('name',tk)[:14],
            sector=pos.get('sector','—'), units=units, avg_price=avg_px,
            cur_price=cur_px, cost=cost, value=value,
            pnl=pnl, pnl_pct=pnl_pct, day_chg=day_chg(tk)))
    enriched.sort(key=lambda x: -x['value'])

    total_pnl     = total_value - total_cost
    total_pnl_pct = total_pnl / total_cost if total_cost else 0
    sign_t        = '+' if total_pnl >= 0 else ''
    pos_green = len([p for p in enriched if p['pnl'] >= 0])
    pos_red   = len([p for p in enriched if p['pnl'] < 0])
    best  = max(enriched, key=lambda x: x['pnl_pct']) if enriched else None
    worst = min(enriched, key=lambda x: x['pnl_pct']) if enriched else None
    movers  = [p for p in enriched if p['day_chg'] is not None]
    top_day = max(movers, key=lambda x: x['day_chg']) if movers else None
    bot_day = min(movers, key=lambda x: x['day_chg']) if movers else None

    now = datetime.now()
    today_str = now.strftime('%d de %B de %Y')
    month_str = now.strftime('%B %Y')
    time_str  = now.strftime('%H:%M')
    pnl_dir   = 'ganancia' if total_pnl >= 0 else 'pérdida'

    spy_px  = prices_cur.get('SPY', 0)
    spy_chg = day_chg('SPY') or 0
    vix_px  = prices_cur.get('^VIX', 0)
    tnx_px  = prices_cur.get('^TNX', 0)
    mxn_px  = prices_cur.get('MXN=X', 0)
    gold_px = prices_cur.get('GC=F', 0)
    oil_px  = prices_cur.get('CL=F', 0)

    def fmt_chg(v, decimals=2):
        if v is None: return '—'
        return ('+' if v >= 0 else '') + f'{v*100:.{decimals}f}%'

    # ── Performance periods desde histórico ───────────────────────
    perf_data = fetch_performance()
    port_ret  = {}   # {'ytd':..., 'qtd':..., 'itd':...}
    spy_ret   = {}
    import pandas as pd

    try:
        dates  = pd.to_datetime(perf_data['dates'])
        port_v = perf_data['portfolio_values']
        spy_v  = perf_data.get('spy_prices', [])
        df = pd.DataFrame({'port': port_v, 'spy': spy_v}, index=dates)

        def period_ret(df, col, cutoff_date):
            sub = df[df.index >= pd.Timestamp(cutoff_date)]
            if len(sub) < 2: return None
            return (sub[col].iloc[-1] / sub[col].iloc[0]) - 1

        yr = now.year
        ytd_start = f'{yr}-01-01'
        qtd_month = ((now.month - 1) // 3) * 3 + 1
        qtd_start = f'{yr}-{qtd_month:02d}-01'

        port_ret['ytd'] = period_ret(df, 'port', ytd_start)
        port_ret['qtd'] = period_ret(df, 'port', qtd_start)
        port_ret['itd'] = total_pnl_pct

        spy_ret['ytd']  = period_ret(df, 'spy', ytd_start)
        spy_ret['qtd']  = period_ret(df, 'spy', qtd_start)
        spy_ret['itd']  = (df['spy'].iloc[-1] / df['spy'].iloc[0]) - 1 if len(df) else None
    except Exception as e:
        print(f'  ⚠ Period calc: {e}', flush=True)

    def alpha(p, b):
        if p is None or b is None: return None
        return p - b

    # ── Activity summary desde transacciones ──────────────────────
    ytd_buys = ytd_sells = 0.0; ytd_txn_count = 0
    for tx in txns:
        try:
            tx_date = datetime.strptime(tx['date'], '%Y-%m-%d')
            if tx_date.year == now.year:
                total_tx = tx['price'] * tx['quantity']
                if tx['type'] == 'buy':  ytd_buys  += total_tx
                else:                    ytd_sells += total_tx
                ytd_txn_count += 1
        except: pass

    # ── Gráfico rendimiento ───────────────────────────────────────
    chart_perf_buf = None
    CHART_W = COL
    CHART_H = COL * 0.46
    try:
        p0 = perf_data['portfolio_values'][0] or 1
        s0 = perf_data['spy_prices'][0] or 1 if perf_data.get('spy_prices') else 1
        pn = [v/p0*100 for v in perf_data['portfolio_values']]
        sn = [v/s0*100 for v in perf_data['spy_prices']] if perf_data.get('spy_prices') else None
        dts = pd.to_datetime(perf_data['dates'])

        dpi = 160
        fw  = CHART_W / 72 * (dpi/72) * 1.45
        fh  = CHART_H / 72 * (dpi/72) * 1.45
        fig, ax = plt.subplots(figsize=(fw, fh))
        fig.patch.set_facecolor('white')
        ax.set_facecolor('white')
        c_port = '#0D2B4E'
        c_spy  = '#C9A248'
        ax.plot(dts, pn, color=c_port, linewidth=1.5, label='Portafolio', zorder=3)
        if sn:
            ax.plot(dts, sn, color=c_spy, linewidth=0.9, linestyle='--',
                    label='SPY (benchmark)', alpha=0.9, zorder=2)
        ax.fill_between(dts, pn, 100, alpha=0.07, color=c_port, zorder=1)
        ax.axhline(100, color='#B0BEC5', linewidth=0.4, linestyle=':')
        ax.set_ylabel('Base 100', fontsize=5.5, color='#5A5A5A')
        ax.tick_params(colors='#5A5A5A', labelsize=5.5)
        for spine in ax.spines.values():
            spine.set_edgecolor('#D0D7DE'); spine.set_linewidth(0.4)
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%b %y'))
        ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
        ax.legend(fontsize=5.5, framealpha=0.9, edgecolor='#D0D7DE',
                  loc='best', handlelength=1.2)
        ax.grid(True, color='#E8ECF0', linewidth=0.3, alpha=0.8)
        plt.tight_layout(pad=0.4)
        chart_perf_buf = io.BytesIO()
        plt.savefig(chart_perf_buf, format='png', dpi=dpi,
                    facecolor='white', bbox_inches='tight')
        plt.close(fig)
        chart_perf_buf.seek(0)
    except Exception as e:
        print(f'  ⚠ Perf chart: {e}', flush=True)

    # ── Gráfico donut ─────────────────────────────────────────────
    chart_alloc_buf = None
    DONUT_W = COL
    DONUT_H = COL * 0.72
    try:
        vals   = [p['value'] for p in enriched]
        labels = [p['ticker'] for p in enriched]
        tv     = sum(vals) or 1
        pal    = ['#0D2B4E','#1A4A7A','#C9A248','#2E6DA4','#3D8EBF',
                  '#5AADE0','#7DC4ED','#A4D4F0','#B0BEC5','#CFD8DC',
                  '#E8A838','#F0C060']
        col    = [pal[i % len(pal)] for i in range(len(labels))]

        dpi = 160
        fw  = DONUT_W / 72 * (dpi/72) * 1.45
        fh  = DONUT_H / 72 * (dpi/72) * 1.45
        fig, ax = plt.subplots(figsize=(fw, fh))
        fig.patch.set_facecolor('white')
        ax.set_facecolor('white')
        wedges, _ = ax.pie(vals, colors=col, startangle=90,
                           wedgeprops=dict(width=0.50, edgecolor='white', linewidth=0.6))
        ax.text(0, 0.08, f'${total_value:,.0f}', ha='center', va='center',
                fontsize=6.5, fontweight='bold', color='#0D2B4E')
        ax.text(0, -0.2, 'USD total', ha='center', va='center',
                fontsize=5, color='#5A5A5A')
        leg_lbl = [f'{lb}  {v/tv*100:.1f}%' for lb,v in zip(labels, vals)]
        ax.legend(wedges, leg_lbl, loc='lower center',
                  bbox_to_anchor=(0.5, -0.46), ncol=3,
                  fontsize=5, framealpha=0.9, edgecolor='#D0D7DE',
                  handlelength=0.9, handleheight=0.7, columnspacing=0.7)
        plt.tight_layout(pad=0.2)
        chart_alloc_buf = io.BytesIO()
        plt.savefig(chart_alloc_buf, format='png', dpi=dpi,
                    facecolor='white', bbox_inches='tight')
        plt.close(fig)
        chart_alloc_buf.seek(0)
    except Exception as e:
        print(f'  ⚠ Alloc chart: {e}', flush=True)

    # ── Chrome ────────────────────────────────────────────────────
    def _chrome(canv, doc):
        canv.saveState()
        canv.setFillColor(NAVY)
        canv.rect(0, H-HDR, W, HDR, fill=1, stroke=0)
        canv.setFillColor(GOLD)
        canv.rect(0, H-HDR-1.5, W, 1.5, fill=1, stroke=0)
        canv.setFillColor(WHITE)
        canv.setFont('Helvetica-Bold', 8)
        canv.drawString(MX, H-14, 'AKILA CAPITAL PARTNERS')
        canv.setFont('Helvetica', 6.5)
        canv.drawCentredString(W/2, H-14, f'REPORTE DE PORTAFOLIO  ·  {month_str.upper()}')
        canv.drawRightString(W-MX, H-14, f'CONFIDENCIAL  ·  {time_str} hrs')
        canv.setFillColor(NAVY)
        canv.rect(0, 0, W, FTR, fill=1, stroke=0)
        canv.setFillColor(GOLD)
        canv.rect(0, FTR, W, 1.2, fill=1, stroke=0)
        canv.setFillColor(WHITE)
        canv.setFont('Helvetica', 5.8)
        canv.drawString(MX, 4.5,
            'Para uso interno exclusivo · No constituye asesoría de inversión · '
            'Rendimientos pasados no garantizan resultados futuros · '
            'Precios: Yahoo Finance')
        canv.drawRightString(W-MX, 4.5, today_str)
        canv.restoreState()

    # ── Story ─────────────────────────────────────────────────────
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
        leftMargin=MX, rightMargin=MX,
        topMargin=MY+HDR, bottomMargin=MY+FTR,
        title=f'Akila Capital Partners — Portfolio Report {month_str}',
        author='Akila Capital Partners')
    story = []

    # ══════════════════════════════════════════════
    # FULL WIDTH — RESUMEN EJECUTIVO
    # ══════════════════════════════════════════════
    story.append(sec_bar('RESUMEN EJECUTIVO  ·  MIGUEL CAMARA  ·  AKILA CAPITAL PARTNERS'))
    story.append(sp(3))

    # Párrafo 1: estado del portafolio
    p1 = (
        f'Al {today_str} ({time_str} hrs), el portafolio <b>Akila Capital Partners</b> '
        f'registra un valor de mercado de <b>${total_value:,.2f} USD</b> sobre una base de '
        f'costo de ${total_cost:,.2f} USD, acumulando una {pnl_dir} no realizada de '
        f'<b>{sign_t}${abs(total_pnl):,.2f} ({sign_t}{total_pnl_pct*100:.2f}%)</b> desde '
        f'inicio. La cartera mantiene <b>{len(enriched)} posiciones activas</b> en '
        f'{len(set(p["sector"] for p in enriched))} sectores, con '
        f'{pos_green} en terreno positivo y {pos_red} en negativo.'
    )
    story.append(Paragraph(p1, body))

    # Párrafo 2: market commentary breve
    vix_desc  = 'elevada — entorno de cautela recomendado' if vix_px > 25 else 'moderada — condiciones de mercado estables'
    best_s    = f"{best['ticker']} ({sign_t}{best['pnl_pct']*100:.1f}%)" if best else '—'
    worst_s   = f"{worst['ticker']} ({worst['pnl_pct']*100:.1f}%)" if worst else '—'
    spy_dir   = 'avanza' if spy_chg >= 0 else 'retrocede'
    p2 = (
        f'El S&P 500 (SPY) {spy_dir} <b>{fmt_chg(spy_chg)}</b> en la sesión, '
        f'cotizando a ${spy_px:,.2f}. El VIX se ubica en <b>{vix_px:.1f} pts</b>, '
        f'señalando volatilidad implícita {vix_desc}. '
        f'El US 10Y en {tnx_px:.2f}% y el USD/MXN en ${mxn_px:.2f} completan '
        f'el cuadro macro. Mayor retorno del portafolio: <b>{best_s}</b> · '
        f'Mayor retroceso: <b>{worst_s}</b>.'
    )
    story.append(Paragraph(p2, body))
    story.append(sp(4))

    # ── Fila 4 tarjetas ──────────────────────────────────────────
    pnl_cv = S('pcv', fontSize=10, textColor=GREEN if total_pnl>=0 else RED,
               fontName='Helvetica-Bold', alignment=TA_CENTER, leading=12)
    spy_cv = S('scv', fontSize=10, textColor=GREEN if spy_chg>=0 else RED,
               fontName='Helvetica-Bold', alignment=TA_CENTER, leading=12)
    c4 = Table([
        [Paragraph('Valor de Mercado', card_l), Paragraph('P&L Acumulado', card_l),
         Paragraph('SPY (hoy)', card_l),         Paragraph('VIX', card_l)],
        [Paragraph(f'${total_value:,.2f}', card_v),
         Paragraph(f'{sign_t}${abs(total_pnl):,.2f}', pnl_cv),
         Paragraph(fmt_chg(spy_chg), spy_cv),
         Paragraph(f'{vix_px:.1f}', card_v)],
        [Paragraph('USD Total', card_s),
         Paragraph(f'{sign_t}{total_pnl_pct*100:.2f}% s/costo', card_s),
         Paragraph(f'${spy_px:,.2f}', card_s),
         Paragraph('Volatilidad implícita', card_s)],
    ], colWidths=[CW/4]*4, rowHeights=[10,16,9])
    c4.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,-1),LGRAY),
        ('BOX',(0,0),(-1,-1),0.5,MGRAY),
        ('LINEBELOW',(0,0),(-1,0),1.2,GOLD),
        ('INNERGRID',(0,0),(-1,-1),0.3,WHITE),
        ('TOPPADDING',(0,0),(-1,-1),2),
        ('BOTTOMPADDING',(0,0),(-1,-1),2),
    ]))
    story.append(c4)
    story.append(sp(4))

    # ── Tabla performance QTD/YTD/ITD vs SPY ─────────────────────
    def fmt_period(v):
        if v is None: return Paragraph('N/D', tc)
        s = ('+' if v >= 0 else '') + f'{v*100:.2f}%'
        return Paragraph(s, tpos if v >= 0 else tneg)

    def fmt_alpha(v):
        if v is None: return Paragraph('—', tc)
        s = ('+' if v >= 0 else '') + f'{v*100:.2f}%'
        sty = S('fa', fontSize=6.5, textColor=GREEN if v>=0 else RED,
                fontName='Helvetica-Bold', alignment=TA_CENTER, leading=8)
        return Paragraph(s, sty)

    perf_rows = [
        [Paragraph('Portafolio Akila', tl),
         fmt_period(port_ret.get('qtd')), fmt_period(port_ret.get('ytd')),
         fmt_period(port_ret.get('itd'))],
        [Paragraph('Benchmark (SPY)', tl),
         fmt_period(spy_ret.get('qtd')), fmt_period(spy_ret.get('ytd')),
         fmt_period(spy_ret.get('itd'))],
        [Paragraph('Alpha vs Benchmark', S('ta',fontSize=6.5,textColor=NAVY,
                   fontName='Helvetica-Bold',alignment=TA_LEFT,leading=8)),
         fmt_alpha(alpha(port_ret.get('qtd'), spy_ret.get('qtd'))),
         fmt_alpha(alpha(port_ret.get('ytd'), spy_ret.get('ytd'))),
         fmt_alpha(alpha(port_ret.get('itd'), spy_ret.get('itd')))],
    ]
    perf_cw = [CW*0.34, CW*0.22, CW*0.22, CW*0.22]
    pt = make_tbl(['', 'QTD', 'YTD', 'Desde Inicio'], perf_rows, perf_cw, bold_last=True)
    pt.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,0),NAVY),
        ('ROWBACKGROUNDS',(0,1),(-1,-1),[WHITE,LGRAY]),
        ('BACKGROUND',(0,-1),(-1,-1),LGOLD),
        ('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),
        ('GRID',(0,0),(-1,-1),0.25,MGRAY),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('TOPPADDING',(0,0),(-1,-1),2.5),
        ('BOTTOMPADDING',(0,0),(-1,-1),2.5),
        ('LEFTPADDING',(0,0),(-1,-1),4),
        ('RIGHTPADDING',(0,0),(-1,-1),4),
        ('LINEABOVE',(0,-1),(-1,-1),0.8,GOLD),
    ]))
    story.append(pt)
    story.append(sp(2))
    story.append(Paragraph(
        '* Rendimientos calculados sobre precio promedio ponderado de adquisición. '
        'Benchmark: SPDR S&P 500 ETF (SPY). QTD = Quarter-to-date · YTD = Year-to-date · '
        'ITD = Inception-to-date. Fuente: Yahoo Finance.',
        note))
    story.append(sp(4))
    story.append(HRFlowable(width=CW, thickness=0.8, color=GOLD, spaceAfter=4, spaceBefore=0))

    # ══════════════════════════════════════════════
    # DOS COLUMNAS 50/50
    # ══════════════════════════════════════════════
    L = []   # Izquierda
    R = []   # Derecha

    # ── IZQUIERDA ─────────────────────────────────────────────────

    # 1. Rendimiento histórico + driver del día
    L.append(mini_sec('RENDIMIENTO HISTÓRICO VS. BENCHMARK', w=COL))
    L.append(sp(2))
    L.append(Paragraph(
        f'Evolución normalizada (Base 100) del portafolio vs. SPY '
        f'en los últimos 400 días de mercado. '
        f'Retorno acumulado desde inicio: <b>{sign_t}{total_pnl_pct*100:.2f}%</b>.',
        note))
    L.append(sp(2))
    if chart_perf_buf:
        L.append(Image(chart_perf_buf, width=COL, height=CHART_H))
    L.append(sp(2))
    L.append(Paragraph(
        '* Base 100 desde el primer registro histórico disponible. '
        'Línea punteada = SPY benchmark.',
        note))
    L.append(sp(4))

    # 2. Driver del día
    L.append(mini_sec('DRIVER DEL DÍA  ·  MACRO SNAPSHOT', w=COL))
    L.append(sp(2))
    if top_day and bot_day:
        driver = (
            f'Sesión de hoy: <b>{top_day["ticker"]}</b> lidera ganancias en cartera '
            f'(<b>{fmt_chg(top_day["day_chg"])}</b>); <b>{bot_day["ticker"]}</b> '
            f'registra mayor presión ({fmt_chg(bot_day["day_chg"])}). '
            f'SPY {fmt_chg(spy_chg)} · VIX {vix_px:.1f} pts · '
            f'US 10Y {tnx_px:.2f}% · USD/MXN ${mxn_px:.2f} · '
            f'Oro ${gold_px:,.0f} · WTI ${oil_px:.1f}.'
        )
    else:
        driver = f'SPY {fmt_chg(spy_chg)} · VIX {vix_px:.1f} · US 10Y {tnx_px:.2f}% · USD/MXN ${mxn_px:.2f}.'
    L.append(Paragraph(driver, note))
    L.append(sp(2))

    macro_rows = []
    for nm, tk, sfx in [
        ('S&P 500 (SPY)','SPY',''),('VIX','^VIX',''),
        ('US 10Y','^TNX','%'),('USD/MXN','MXN=X',''),
        ('Oro (XAU)','GC=F',''),('WTI Crude','CL=F',''),
    ]:
        px  = prices_cur.get(tk, 0)
        chg = day_chg(tk)
        c2  = tpos if (chg or 0) >= 0 else tneg
        macro_rows.append([
            Paragraph(nm, tl),
            Paragraph(f'${px:,.2f}{sfx}' if px else '—', tc),
            Paragraph(fmt_chg(chg), c2),
        ])
    L.append(make_tbl(['Indicador','Precio','Var. Día'], macro_rows,
                       [COL*0.50, COL*0.28, COL*0.22], tiny=True))
    L.append(sp(4))

    # 3. Holdings
    L.append(mini_sec('POSICIONES ACTIVAS  ·  HOLDINGS DETAIL', w=COL))
    L.append(sp(2))
    L.append(Paragraph(
        f'{len(enriched)} posiciones ordenadas por valor de mercado. '
        f'P&L calculado neto de comisiones sobre precio promedio ponderado.',
        note))
    L.append(sp(2))
    cw_p = [COL*0.14, COL*0.15, COL*0.17, COL*0.18, COL*0.18, COL*0.18]
    pos_rows = []
    for p in enriched:
        alloc = p['value'] / total_value * 100 if total_value else 0
        sg = '+' if p['pnl'] >= 0 else ''
        pc = tpos if p['pnl'] >= 0 else tneg
        dc = tpos if (p['day_chg'] or 0) >= 0 else tneg
        pos_rows.append([
            Paragraph(p['ticker'], tl),
            Paragraph(f"${p['cur_price']:,.1f}" if p['cur_price'] else '—', tc),
            Paragraph(fmt_chg(p['day_chg']), dc),
            Paragraph(f"{sg}${abs(p['pnl']):,.1f}", pc),
            Paragraph(f"{sg}{p['pnl_pct']*100:.1f}%", pc),
            Paragraph(f"{alloc:.1f}%", tc),
        ])
    L.append(make_tbl(['Ticker','Precio','Hoy','P&L $','P&L %','Alloc.'],
                       pos_rows, cw_p, tiny=True))

    # ── DERECHA ──────────────────────────────────────────────────

    # 1. Distribución
    R.append(mini_sec('DISTRIBUCIÓN DE CARTERA  ·  ASSET ALLOCATION', w=COL))
    R.append(sp(2))
    conc3 = sum(p['value'] for p in enriched[:3]) / total_value * 100 if total_value else 0
    R.append(Paragraph(
        f'Cartera de <b>{len(enriched)} posiciones</b> distribuida en '
        f'{len(set(p["sector"] for p in enriched))} sectores. '
        f'Las tres posiciones más grandes ({enriched[0]["ticker"]}, '
        f'{enriched[1]["ticker"] if len(enriched)>1 else "—"}, '
        f'{enriched[2]["ticker"] if len(enriched)>2 else "—"}) '
        f'concentran el <b>{conc3:.1f}%</b> del portafolio.',
        note))
    R.append(sp(2))
    if chart_alloc_buf:
        R.append(Image(chart_alloc_buf, width=COL, height=DONUT_H))
    R.append(sp(3))

    # 2. Performance & métricas
    R.append(mini_sec('MÉTRICAS CLAVE DEL PORTAFOLIO', w=COL))
    R.append(sp(2))
    top3 = sorted(enriched, key=lambda x: -x['pnl'])[:3]
    top3_str = ' · '.join(f"{p['ticker']} ({'+' if p['pnl']>=0 else ''}${abs(p['pnl']):,.0f})"
                           for p in top3)
    R.append(Paragraph(
        f'Top contribuidores al P&L: <b>{top3_str}</b>. '
        f'VIX en {vix_px:.1f} — '
        f'{"volatilidad elevada, revisar coberturas." if vix_px>25 else "condiciones estables de mercado."}',
        note))
    R.append(sp(2))

    met_rows = [
        [Paragraph('Valor Total (MKT)', tl), Paragraph(f'${total_value:,.2f}', tc_r)],
        [Paragraph('Costo Base Total',  tl), Paragraph(f'${total_cost:,.2f}', tc_r)],
        [Paragraph('P&L No Realizado',  tl), Paragraph(f'{sign_t}${abs(total_pnl):,.2f}',
                                             tpos if total_pnl>=0 else tneg)],
        [Paragraph('Retorno Total ITD', tl), Paragraph(fmt_chg(total_pnl_pct),
                                             tpos if total_pnl>=0 else tneg)],
        [Paragraph('Retorno YTD',       tl), Paragraph(fmt_chg(port_ret.get('ytd')),
                                             tpos if (port_ret.get('ytd') or 0)>=0 else tneg)],
        [Paragraph('Alpha YTD vs SPY',  tl), Paragraph(
                                             fmt_chg(alpha(port_ret.get('ytd'),spy_ret.get('ytd'))),
                                             tpos if (alpha(port_ret.get('ytd'),spy_ret.get('ytd')) or 0)>=0 else tneg)],
        [Paragraph('Posiciones Verdes', tl), Paragraph(f'{pos_green} de {len(enriched)}', tc)],
        [Paragraph('Concentración Top 3', tl), Paragraph(f'{conc3:.1f}%', tc)],
    ]
    cw_mt = [COL*0.60, COL*0.40]
    R.append(make_tbl(['Métrica','Valor'], met_rows, cw_mt, tiny=True))
    R.append(sp(3))

    # 3. Activity summary
    R.append(mini_sec('ACTIVIDAD YTD  ·  ACTIVITY SUMMARY', w=COL))
    R.append(sp(2))
    R.append(Paragraph(
        f'En lo que va del año <b>{now.year}</b>, el portafolio registra '
        f'<b>{ytd_txn_count} operaciones</b>: compras por '
        f'<b>${ytd_buys:,.2f}</b> y ventas por <b>${ytd_sells:,.2f}</b>. '
        f'Capital neto desplegado en el período: '
        f'<b>${ytd_buys - ytd_sells:,.2f}</b>.',
        note))
    R.append(sp(2))

    act_rows = [
        [Paragraph('Operaciones YTD', tl), Paragraph(str(ytd_txn_count), tc)],
        [Paragraph('Compras YTD', tl), Paragraph(f'${ytd_buys:,.2f}', tc)],
        [Paragraph('Ventas YTD', tl), Paragraph(f'${ytd_sells:,.2f}', tc)],
        [Paragraph('Capital Neto Desplegado', tl), Paragraph(f'${ytd_buys-ytd_sells:,.2f}', tc)],
        [Paragraph('Total Operaciones Históricas', tl), Paragraph(str(len(txns)), tc)],
    ]
    R.append(make_tbl(['Concepto','Valor'], act_rows, cw_mt, tiny=True))
    R.append(sp(3))

    # 4. Callout riesgo / recomendación
    risk_txt = (
        f'Concentración en top 3: {conc3:.1f}% · '
        f'VIX {vix_px:.1f} pts — '
        f'{"Volatilidad elevada: revisar stop-loss y coberturas." if vix_px>25 else "Mercado estable: mantener posicionamiento actual."} '
        f'Próxima revisión recomendada: fin de trimestre.'
    )
    R.append(callout(risk_txt, w=COL,
                     bg=LGRN if total_pnl >= 0 else LRED,
                     border=GREEN if total_pnl >= 0 else RED))

    # ── Ensamblar columnas ────────────────────────────────────────
    def col_wrap(items, w):
        t = Table([[item] for item in items], colWidths=[w])
        t.setStyle(TableStyle([
            ('TOPPADDING',(0,0),(-1,-1),0),('BOTTOMPADDING',(0,0),(-1,-1),0),
            ('LEFTPADDING',(0,0),(-1,-1),0),('RIGHTPADDING',(0,0),(-1,-1),0),
        ]))
        return t

    two_col = Table(
        [[col_wrap(L, COL), Spacer(GAP, 1), col_wrap(R, COL)]],
        colWidths=[COL, GAP, COL])
    two_col.setStyle(TableStyle([
        ('VALIGN',(0,0),(-1,-1),'TOP'),
        ('TOPPADDING',(0,0),(-1,-1),0),('BOTTOMPADDING',(0,0),(-1,-1),0),
        ('LEFTPADDING',(0,0),(-1,-1),0),('RIGHTPADDING',(0,0),(-1,-1),0),
    ]))
    story.append(two_col)

    doc.build(story, onFirstPage=_chrome, onLaterPages=_chrome)
    buf.seek(0)
    return buf.read()



# ── Servidor HTTP ─────────────────────────────────────────────
class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass  # silenciar logs HTTP

    def handle_error(self, request, client_address):
        pass  # ignorar errores de conexión (BrokenPipe, etc.)

    def handle_one_request(self):
        try:
            super().handle_one_request()
        except (BrokenPipeError, ConnectionResetError):
            pass  # cliente cerró la conexión antes de recibir respuesta

    def _get_portfolio_param(self):
        """Extrae el parámetro ?portfolio= del path, default 'miguel'."""
        from urllib.parse import urlparse, parse_qs
        params = parse_qs(urlparse(self.path).query)
        p = params.get("portfolio", ["miguel"])[0].lower().strip()
        if p not in PORTFOLIO_NAMES + ["akila"]:
            p = "miguel"
        return p

    def do_GET(self):
        from urllib.parse import urlparse
        base_path = urlparse(self.path).path

        if base_path == "/api/portfolios":
            # Lista los portafolios disponibles
            payload = json.dumps({
                "portfolios": [
                    {"id": "miguel", "name": "Miguel",  "label": "M"},
                    {"id": "paulo",  "name": "Paulo",   "label": "P"},
                    {"id": "akila",  "name": "AKILA",   "label": "A"},
                ]
            })
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(payload.encode())
            return

        if base_path == "/api/data":
            portfolio = self._get_portfolio_param()
            import math
            def clean(obj):
                if isinstance(obj, float):
                    return None if (math.isnan(obj) or math.isinf(obj)) else obj
                if isinstance(obj, dict):
                    return {k: clean(v) for k, v in obj.items()}
                if isinstance(obj, list):
                    return [clean(v) for v in obj]
                return obj
            with cache_lock:
                data = json.dumps(clean(get_cache(portfolio)), default=str)
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(data.encode())

        elif base_path.startswith("/api/tearsheet") or self.path.startswith("/api/tearsheet"):
            from urllib.parse import urlparse, parse_qs
            params = parse_qs(urlparse(self.path).query)
            ticker    = params.get("ticker", [None])[0]
            portfolio = params.get("portfolio", ["miguel"])[0]
            if not ticker:
                self.send_response(400)
                self.end_headers()
                return
            data = get_tearsheet_cached(ticker, portfolio)
            import math
            def clean(obj):
                if isinstance(obj, float):
                    return None if (math.isnan(obj) or math.isinf(obj)) else obj
                if isinstance(obj, dict):
                    return {k: clean(v) for k, v in obj.items()}
                if isinstance(obj, list):
                    return [clean(v) for v in obj]
                return obj
            payload = json.dumps(clean(data), default=str)
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(payload.encode())

        elif base_path == "/api/performance":
            portfolio = self._get_portfolio_param()
            with perf_lock:
                pc = perf_cache.setdefault(portfolio, {"data": None, "ts": 0})
                if pc["data"] is not None and time.time() - pc["ts"] < PERF_TTL:
                    data = pc["data"]
                else:
                    data = fetch_performance(portfolio)
                    pc["data"] = data
                    pc["ts"]   = time.time()
            payload = json.dumps(data, ensure_ascii=False)
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(payload.encode())

        elif base_path == "/api/performance/short":
            portfolio = self._get_portfolio_param()
            with perf_lock:
                psc = perf_short_cache.setdefault(portfolio, {"data": None, "ts": 0})
                if psc["data"] is not None and time.time() - psc["ts"] < PERF_SHORT_TTL:
                    data = psc["data"]
                else:
                    data = fetch_performance_short(portfolio)
                    psc["data"] = data
                    psc["ts"]   = time.time()
            payload = json.dumps(data, ensure_ascii=False)
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(payload.encode())

        elif base_path == "/api/cash":
            portfolio = self._get_portfolio_param()
            if portfolio == "akila":
                balance = load_cash("miguel") + load_cash("paulo")
            else:
                balance = load_cash(portfolio)
            payload = json.dumps({"balance": balance}, ensure_ascii=False)
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(payload.encode())

        elif base_path == "/api/risk":
            portfolio = self._get_portfolio_param()
            data = get_risk_cached(portfolio)
            payload = json.dumps(data, ensure_ascii=False, default=str)
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(payload.encode())

        elif base_path == "/api/watchlist":
            # GET /api/watchlist — devuelve lista con metadatos (sin datos de mercado)
            with watchlist_lock:
                items = load_watchlist()
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(json.dumps(items, ensure_ascii=False).encode())

        elif base_path.startswith("/api/watchlist/"):
            # GET /api/watchlist/{TICKER} — devuelve datos de mercado en tiempo real
            ticker = base_path.split("/api/watchlist/")[-1].strip().upper()
            data   = fetch_watchlist_data(ticker)
            # Merge con metadatos guardados (notas, estado)
            with watchlist_lock:
                items = load_watchlist()
            meta = next((i for i in items if i["ticker"] == ticker), {})
            data["notes"]  = meta.get("notes", "")
            data["status"] = meta.get("status", "watching")
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(json.dumps(data, ensure_ascii=False).encode())

        elif base_path == "/api/transactions":
            portfolio = self._get_portfolio_param()
            with transactions_lock:
                if portfolio == "akila":
                    # Consolidado: combinar transacciones de Miguel y Paulo con etiqueta de origen
                    miguel_txs = [dict(tx, _owner="Miguel") for tx in load_transactions("miguel")]
                    paulo_txs  = [dict(tx, _owner="Paulo")  for tx in load_transactions("paulo")]
                    txs = sorted(miguel_txs + paulo_txs, key=lambda t: t.get("date",""), reverse=True)
                else:
                    txs = load_transactions(portfolio)
            payload = json.dumps(txs, ensure_ascii=False)
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(payload.encode())

        elif base_path == "/api/positions":
            portfolio = self._get_portfolio_param()
            with transactions_lock:
                computed = compute_positions_from_transactions(load_transactions(portfolio))
            payload = json.dumps(computed, ensure_ascii=False)
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(payload.encode())

        elif base_path == "/api/analysis":
            portfolio = self._get_portfolio_param()
            with analysis_lock:
                payload = json.dumps(load_analysis(portfolio), ensure_ascii=False)
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(payload.encode())

        elif base_path.startswith("/api/analysis/live") or self.path.startswith("/api/analysis/live"):
            from urllib.parse import urlparse, parse_qs
            params = parse_qs(urlparse(self.path).query)
            ticker = params.get("ticker", [None])[0]
            if not ticker:
                self.send_response(400)
                self.send_header("Content-Type", "application/json")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(b'{"error":"Missing ticker"}')
                return
            print(f"  📈 Analysis live: {ticker}", flush=True)
            data = fetch_analysis_live(ticker)
            import math
            def clean_analysis(obj):
                if isinstance(obj, float):
                    return None if (math.isnan(obj) or math.isinf(obj)) else obj
                if isinstance(obj, dict):
                    return {k: clean_analysis(v) for k, v in obj.items()}
                if isinstance(obj, list):
                    return [clean_analysis(v) for v in obj]
                return obj
            payload = json.dumps(clean_analysis(data), ensure_ascii=False, default=str)
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(payload.encode())


        elif base_path == "/api/macro-scorecard":
            try:
                sc_path = os.path.join(BASE_DIR, "macro_scorecard.json")
                with open(sc_path, "r", encoding="utf-8") as f:
                    sc_data = f.read()
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(sc_data.encode())
            except Exception as e:
                self.send_response(500)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(json.dumps({"error": str(e)}).encode())
            return

        elif base_path == "/api/export-pdf":
            print("  📄 Generando PDF...", flush=True)
            try:
                pdf_bytes = generate_portfolio_pdf()
                from datetime import datetime
                fname = f"Akila_Portfolio_Report_{datetime.now().strftime('%Y%m')}.pdf"
                self.send_response(200)
                self.send_header("Content-Type", "application/pdf")
                self.send_header("Content-Disposition", f'attachment; filename="{fname}"')
                self.send_header("Content-Length", str(len(pdf_bytes)))
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(pdf_bytes)
                print(f"  ✅ PDF generado: {fname} ({len(pdf_bytes)//1024}KB)", flush=True)
            except Exception as e:
                import traceback
                print(f"  ❌ Error PDF: {e}", flush=True)
                traceback.print_exc()
                self.send_response(500)
                self.send_header("Content-Type", "application/json")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"error": str(e)}).encode())

        elif self.path == "/" or self.path == "/index.html":
            html_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                     "dashboard.html")
            if os.path.exists(html_path):
                with open(html_path, "rb") as f:
                    content = f.read()
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.end_headers()
                self.wfile.write(content)
            else:
                self.send_response(404)
                self.end_headers()
        else:
            self.send_response(404)
            self.end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, DELETE, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_POST(self):
        from urllib.parse import urlparse
        global TRANSACTIONS, POSITIONS
        base_path = urlparse(self.path).path

        if base_path == "/api/watchlist":
            # POST /api/watchlist — añade un ticker o actualiza notas/estado
            length = int(self.headers.get("Content-Length", 0))
            body   = self.rfile.read(length)
            try:
                data   = json.loads(body)
                ticker = data.get("ticker", "").strip().upper()
                if not ticker:
                    raise ValueError("ticker requerido")
                with watchlist_lock:
                    items = load_watchlist()
                    existing = next((i for i in items if i["ticker"] == ticker), None)
                    if existing:
                        # Actualizar campos editables
                        if "notes"  in data: existing["notes"]  = data["notes"]
                        if "status" in data: existing["status"] = data["status"]
                    else:
                        items.append({
                            "ticker":   ticker,
                            "notes":    data.get("notes", ""),
                            "status":   data.get("status", "watching"),
                            "added_at": datetime.now().strftime("%Y-%m-%d"),
                        })
                    save_watchlist(items)
                print(f"  👁 Watchlist: {ticker} {'actualizado' if existing else 'añadido'}", flush=True)
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"ok": True, "ticker": ticker}).encode())
            except Exception as e:
                self.send_response(400)
                self.send_header("Content-Type", "application/json")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"ok": False, "error": str(e)}).encode())
            return

        if base_path == "/api/cash":
            portfolio = self._get_portfolio_param()
            length = int(self.headers.get("Content-Length", 0))
            body   = self.rfile.read(length)
            try:
                data    = json.loads(body)
                balance = float(data.get("balance", 0))
                if portfolio == "akila":
                    # No editar AKILA directamente — es suma de los otros
                    self.send_response(400)
                    self.send_header("Content-Type", "application/json")
                    self.send_header("Access-Control-Allow-Origin", "*")
                    self.end_headers()
                    self.wfile.write(json.dumps({"ok": False, "error": "AKILA es consolidado, edita Miguel o Paulo directamente"}).encode())
                    return
                save_cash(balance, portfolio)
                print(f"  💵 Cash actualizado: ${balance:,.2f} ({portfolio})", flush=True)
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"ok": True, "balance": balance}).encode())
            except Exception as e:
                self.send_response(500)
                self.send_header("Content-Type", "application/json")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"ok": False, "error": str(e)}).encode())
            return

        if base_path == "/api/transactions":
            # Registra una nueva transacción (compra o venta)
            import uuid
            length = int(self.headers.get("Content-Length", 0))
            body   = self.rfile.read(length)
            try:
                tx = json.loads(body)
                # Validar campos mínimos
                required = ["ticker", "type", "date", "price", "quantity"]
                if any(tx.get(f) is None or tx.get(f) == "" for f in required):
                    self.send_response(400)
                    self.send_header("Content-Type", "application/json")
                    self.send_header("Access-Control-Allow-Origin", "*")
                    self.end_headers()
                    self.wfile.write(b'{"error":"Faltan campos requeridos"}')
                    return
                if tx["type"] not in ("buy", "sell"):
                    raise ValueError("type debe ser 'buy' o 'sell'")

                tx["id"]       = str(uuid.uuid4())[:8]
                tx["ticker"]   = tx["ticker"].strip().upper()
                tx["price"]    = float(tx["price"])
                tx["quantity"] = float(tx["quantity"])
                tx["notes"]    = tx.get("notes", "").strip()

                portfolio = self._get_portfolio_param()
                with transactions_lock:
                    txs = load_transactions(portfolio)
                    # Validar que venta no exceda posición actual
                    if tx["type"] == "sell":
                        current_pos = compute_positions_from_transactions(txs)
                        held = next((p["units"] for p in current_pos if p["ticker"] == tx["ticker"]), 0.0)
                        if held <= 0:
                            self.send_response(400)
                            self.send_header("Content-Type", "application/json")
                            self.send_header("Access-Control-Allow-Origin", "*")
                            self.end_headers()
                            self.wfile.write(json.dumps({"ok": False, "error": f"No tienes posición abierta en {tx['ticker']}"}).encode())
                            return
                        if tx["quantity"] > held + 1e-9:
                            self.send_response(400)
                            self.send_header("Content-Type", "application/json")
                            self.send_header("Access-Control-Allow-Origin", "*")
                            self.end_headers()
                            self.wfile.write(json.dumps({"ok": False, "error": f"Máximo vendible: {held:.4f} unidades de {tx['ticker']}"}).encode())
                            return
                    txs.append(tx)
                    save_transactions(txs, portfolio)
                    TRANSACTIONS = txs

                    # Ajustar efectivo: +venta / -compra
                    trade_amount = float(tx["price"]) * float(tx["quantity"])
                    current_cash = load_cash(portfolio)
                    if tx["type"] == "sell":
                        new_cash = current_cash + trade_amount
                        print(f"  💵 Cash +${trade_amount:,.2f} → total ${new_cash:,.2f} ({portfolio})", flush=True)
                    else:  # buy
                        new_cash = max(0, current_cash - trade_amount)
                        print(f"  💵 Cash -${trade_amount:,.2f} → total ${new_cash:,.2f} ({portfolio})", flush=True)
                    save_cash(new_cash, portfolio)

                print(f"  📝 Transacción: {tx['type'].upper()} {tx['quantity']} {tx['ticker']} @ ${tx['price']} ({tx['date']})", flush=True)
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"ok": True, "id": tx["id"]}).encode())
            except Exception as e:
                self.send_response(500)
                self.send_header("Content-Type", "application/json")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"error": str(e)}).encode())

        elif base_path == "/api/positions":
            # Agrega o edita una posición
            length = int(self.headers.get("Content-Length", 0))
            body   = self.rfile.read(length)
            try:
                pos = json.loads(body)
                # Validar campos mínimos
                if not pos.get("ticker") or pos.get("units") is None or pos.get("avg_price") is None:
                    self.send_response(400)
                    self.end_headers()
                    self.wfile.write(b'{"error":"Faltan campos requeridos"}')
                    return
                pos["ticker"]    = pos["ticker"].strip().upper()
                pos["units"]     = float(pos["units"])
                pos["avg_price"] = float(pos["avg_price"])
                pos["name"]      = pos.get("name", pos["ticker"])
                pos["sector"]    = pos.get("sector", "—")
                pos["country"]   = pos.get("country", "—")

                with positions_lock:
                    # Si ya existe, reemplazar; si no, agregar
                    idx = next((i for i,p in enumerate(POSITIONS) if p["ticker"] == pos["ticker"]), None)
                    if idx is not None:
                        POSITIONS[idx] = pos
                    else:
                        POSITIONS.append(pos)
                    save_positions(POSITIONS)

                print(f"  ✏️  Posición guardada: {pos['ticker']} x{pos['units']} @ ${pos['avg_price']}", flush=True)
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(b'{"ok":true}')
            except Exception as e:
                self.send_response(500)
                self.end_headers()
                self.wfile.write(json.dumps({"error": str(e)}).encode())
        elif base_path == "/api/analysis":
            # Guarda/actualiza análisis personal de un ticker (tesis, conviction, next_cat)
            length = int(self.headers.get("Content-Length", 0))
            body   = self.rfile.read(length)
            try:
                update = json.loads(body)
                ticker = update.get("ticker", "").strip().upper()
                if not ticker:
                    self.send_response(400)
                    self.send_header("Content-Type", "application/json")
                    self.send_header("Access-Control-Allow-Origin", "*")
                    self.end_headers()
                    self.wfile.write(b'{"error":"Falta ticker"}')
                    return
                portfolio = self._get_portfolio_param()
                with analysis_lock:
                    ana = load_analysis(portfolio)
                    if ticker not in ana:
                        ana[ticker] = {}
                    for field in ["tesis", "conviction", "next_cat"]:
                        if field in update:
                            ana[ticker][field] = update[field]
                    save_analysis(ana, portfolio)
                    global ANALYSIS
                    ANALYSIS = ana
                print(f"  📝 Análisis guardado: {ticker}", flush=True)
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(b'{"ok":true}')
            except Exception as e:
                self.send_response(500)
                self.send_header("Content-Type", "application/json")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"error": str(e)}).encode())
        elif base_path == "/api/restore":
            # ── Importación segura de datos desde local → Railway ──────────
            restore_token = os.environ.get("RESTORE_TOKEN", "")
            if not restore_token:
                self.send_response(403)
                self.send_header("Content-Type", "application/json")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"ok": False, "error": "RESTORE_TOKEN no configurado en Railway"}).encode())
                return
            req_token = self.headers.get("X-Restore-Token", "")
            if req_token != restore_token:
                self.send_response(401)
                self.send_header("Content-Type", "application/json")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"ok": False, "error": "Token inválido"}).encode())
                return
            length = int(self.headers.get("Content-Length", 0))
            body   = self.rfile.read(length)
            try:
                payload = json.loads(body)
                imported = {}
                if "miguel_txs" in payload:
                    save_transactions(payload["miguel_txs"], "miguel")
                    imported["miguel_txs"] = len(payload["miguel_txs"])
                if "paulo_txs" in payload:
                    save_transactions(payload["paulo_txs"], "paulo")
                    imported["paulo_txs"] = len(payload["paulo_txs"])
                if "miguel_analysis" in payload:
                    save_analysis(payload["miguel_analysis"], "miguel")
                    imported["miguel_analysis"] = True
                if "paulo_analysis" in payload:
                    save_analysis(payload["paulo_analysis"], "paulo")
                    imported["paulo_analysis"] = True
                if "miguel_cash" in payload:
                    save_cash(float(payload["miguel_cash"]), "miguel")
                    imported["miguel_cash"] = payload["miguel_cash"]
                if "paulo_cash" in payload:
                    save_cash(float(payload["paulo_cash"]), "paulo")
                    imported["paulo_cash"] = payload["paulo_cash"]
                if "watchlist" in payload:
                    with watchlist_lock:
                        save_watchlist(payload["watchlist"])
                    imported["watchlist"] = len(payload["watchlist"])
                print(f"  📥 Datos restaurados desde local: {imported}", flush=True)
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"ok": True, "imported": imported}).encode())
            except Exception as e:
                print(f"  ⚠️  Restore error: {e}", flush=True)
                self.send_response(500)
                self.send_header("Content-Type", "application/json")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"ok": False, "error": str(e)}).encode())
        else:
            self.send_response(404)
            self.end_headers()

    def do_DELETE(self):
        from urllib.parse import urlparse
        global TRANSACTIONS, POSITIONS
        base_path = urlparse(self.path).path

        if base_path.startswith("/api/watchlist/"):
            ticker = base_path.split("/api/watchlist/")[-1].strip().upper()
            with watchlist_lock:
                items  = load_watchlist()
                before = len(items)
                items  = [i for i in items if i["ticker"] != ticker]
                save_watchlist(items)
            print(f"  🗑️  Watchlist: {ticker} eliminado", flush=True)
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(json.dumps({"ok": True, "removed": before - len(items)}).encode())
            return

        if base_path.startswith("/api/transactions/"):
            portfolio = self._get_portfolio_param()
            tx_id = base_path.split("/api/transactions/")[-1].strip()
            with transactions_lock:
                txs    = load_transactions(portfolio)
                before = len(txs)
                txs    = [t for t in txs if t.get("id") != tx_id]
                save_transactions(txs, portfolio)
                TRANSACTIONS = txs
                removed = before - len(txs)
            print(f"  🗑️  Transacción eliminada: {tx_id}", flush=True)
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(json.dumps({"ok": True, "removed": removed}).encode())

        elif self.path.startswith("/api/positions/"):
            ticker = self.path.split("/api/positions/")[-1].strip().upper()
            with positions_lock:
                before = len(POSITIONS)
                POSITIONS = [p for p in POSITIONS if p["ticker"] != ticker]
                save_positions(POSITIONS)
                removed = before - len(POSITIONS)
            print(f"  🗑️  Posición eliminada: {ticker}", flush=True)
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(json.dumps({"ok": True, "removed": removed}).encode())
        else:
            self.send_response(404)
            self.end_headers()

def main():
    print("=" * 60)
    print("  PORTFOLIO TRACKER — Dashboard Web Real-Time")
    print(f"  {datetime.now().strftime('%d/%m/%Y  %H:%M:%S')}")
    print("=" * 60)
    print(f"\n  Iniciando servidor en http://localhost:{PORT}")
    print(f"  Actualizando precios cada {REFRESH_SEC} segundos")
    print(f"\n  ➡️  Abre Chrome en: http://localhost:{PORT}")
    print(f"\n  Presiona Ctrl+C para detener.\n")

    # Hilo de precios (cada 15s)
    updater = threading.Thread(target=update_data, daemon=True)
    updater.start()

    # Hilo de métricas (cada 5 min, arranca con 20s de delay para no bloquear)
    def metrics_delayed():
        time.sleep(20)
        update_metrics_background()
    metrics_thread = threading.Thread(target=metrics_delayed, daemon=True)
    metrics_thread.start()

    # Hilo de macro scorecard FRED (al arrancar + cada hora)
    macro_sc_thread = threading.Thread(target=update_macro_scorecard_background, daemon=True)
    macro_sc_thread.start()

    # Esperar primera carga
    time.sleep(2)

    # Servidor HTTP
    server = HTTPServer(("0.0.0.0", PORT), Handler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n  👋 Servidor detenido.")

if __name__ == "__main__":
    main()
