#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FOF 组合投后管理 · 周更工具
=========================================================
用途：读取持仓 Excel（基金.xls / .xlsx）→ 拉取基金最新净值与业绩数据
      → 重算全价市值与浮动盈亏 → 输出更新后的 Excel、分析报告 HTML、高清截图
      → 记录历史快照，自动识别本期新增 / 增持 / 减持 / 清仓

用法：
    python fof_update.py                      # 全自动：拉数据 → 更新 → 出报告 + 截图
    python fof_update.py --no-shot            # 跳过截图
    python fof_update.py --holdings "路径"     # 指定持仓表
    python fof_update.py --out "目录"          # 指定输出目录
    python fof_update.py --date 20260916      # 指定估值日标签
    python fof_update.py --nav-json nav.json  # 用外部净值数据覆盖（如同花顺 iFinD 导出）
    python fof_update.py --check              # 只做体检，不写任何文件

数据源：东方财富公开接口（免密钥，可离线定时运行）
        净值：api.fund.eastmoney.com/f10/lsjz   业绩/风险/规模：fund.eastmoney.com/pingzhongdata
"""

import argparse
import json
import math
import os
import re
import shutil
import subprocess
import sys
import time
import urllib.request
import urllib.error
from datetime import datetime, timedelta

# ------------------------------------------------------------------ 依赖
try:
    import xlrd
except ImportError:
    xlrd = None
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    openpyxl = None
try:
    from PIL import Image
except ImportError:
    Image = None

# ------------------------------------------------------------------ 默认配置
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_HOLDINGS = r"D:\BaiduSyncdisk\桌面工作文件\FOF研究\FOF组合投后管理\基金.xls"
DEFAULT_OUTDIR = os.environ.get("FOF_OUTDIR", r"D:\BaiduSyncdisk\桌面工作文件\FOF研究\FOF组合投后管理")
HIST_SUBDIR = "投后管理历史"
TEMPLATE_PATH = os.path.join(BASE_DIR, "report_template.html")
ECHARTS_PATH = os.path.join(BASE_DIR, "echarts.min.js")

CHROME_CANDIDATES = [
    # Linux (GitHub Actions / 云端)
    "/usr/bin/google-chrome", "/usr/bin/google-chrome-stable",
    "/usr/bin/chromium", "/usr/bin/chromium-browser",
    # Windows
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
]

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")
RF = 1.5          # 无风险利率(%)，用于 Sharpe
TRADING_DAYS = 252

RED, GREEN = "#d93f3f", "#1f9c62"


# ================================================================== 小工具
LOG_FILE = None


def log(msg):
    line = "[%s] %s" % (datetime.now().strftime("%H:%M:%S"), msg)
    print(line, flush=True)
    if LOG_FILE:
        try:
            os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)
            with open(LOG_FILE, "a", encoding="utf-8") as fh:
                fh.write(line + "\n")
        except Exception:                   # noqa: BLE001
            pass


def http_get(url, referer=None, retries=3, timeout=25):
    """带重试的 GET，返回文本。"""
    last = None
    for i in range(retries):
        try:
            req = urllib.request.Request(url, headers={
                "User-Agent": UA,
                "Referer": referer or "https://fund.eastmoney.com/",
                "Accept": "*/*",
            })
            with urllib.request.urlopen(req, timeout=timeout) as r:
                raw = r.read()
            for enc in ("utf-8-sig", "utf-8", "gbk", "latin-1"):
                try:
                    return raw.decode(enc)
                except UnicodeDecodeError:
                    continue
            return raw.decode("utf-8", "replace")
        except Exception as e:          # noqa: BLE001
            last = e
            time.sleep(1.2 * (i + 1))
    raise RuntimeError("请求失败 %s：%s" % (url, last))


def num(x, default=0.0):
    """把 '1,234.5' / '12%' / '' / None 统一转 float。"""
    if x is None:
        return default
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip().replace(",", "").replace("%", "").replace("元", "")
    if s in ("", "-", "--", "—", "None", "nan"):
        return default
    try:
        return float(s)
    except ValueError:
        return default


def fmt(v, d=2, plus=False):
    if v is None:
        return "—"
    s = ("{:,.%df}" % d).format(v)
    return ("+" + s) if (plus and v >= 0) else s


def pct(v, d=2, plus=True):
    if v is None:
        return "—"
    return fmt(v, d, plus) + "%"


def js(obj):
    """安全序列化为 JS 字面量。"""
    return json.dumps(obj, ensure_ascii=False, default=str)


def find_header_row(rows, key="基金代码"):
    for i, r in enumerate(rows[:15]):
        for c in r:
            if isinstance(c, str) and key in c:
                return i
    return 0


# ================================================================== 1. 读持仓
def read_holdings(path):
    """读取持仓 Excel，返回 (列名列表, 记录列表)。支持 .xls / .xlsx。"""
    if not os.path.exists(path):
        raise FileNotFoundError("找不到持仓表：%s" % path)
    ext = os.path.splitext(path)[1].lower()

    rows = []
    if ext in (".xlsx", ".xlsm"):
        if openpyxl is None:
            raise RuntimeError("需要 openpyxl 才能读取 .xlsx")
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.worksheets[0]
        for r in ws.iter_rows(values_only=True):
            rows.append(list(r))
    else:
        if xlrd is None:
            raise RuntimeError("需要 xlrd 才能读取 .xls")
        wb = xlrd.open_workbook(path)
        sh = wb.sheet_by_index(0)
        for i in range(sh.nrows):
            rows.append([sh.cell_value(i, c) for c in range(sh.ncols)])

    hi = find_header_row(rows)
    header = [str(c).strip() for c in rows[hi]]
    data = []
    for r in rows[hi + 1:]:
        rec = {}
        blank = True
        for k, v in zip(header, r):
            if not k:
                continue
            rec[k] = v
            if isinstance(v, str) and v.strip():
                blank = False
            elif isinstance(v, (int, float)) and v:
                blank = False
        if not blank and str(rec.get("基金代码", "")).strip():
            data.append(rec)
    return header, data


def pick(rec, *names, default=None):
    """按候选列名取值（模糊包含匹配）。"""
    for n in names:
        if n in rec:
            return rec[n]
    for k, v in rec.items():
        for n in names:
            if n in k:
                return v
    return default


# ================================================================== 2. 取数
_JS_CACHE = {}


def _js_var(text, name):
    """从 pingzhongdata JS 中抽取 var name = <json>;"""
    m = re.search(r"var\s+" + re.escape(name) + r"\s*=\s*", text)
    if not m:
        return None
    s = text[m.end():]
    s = s.lstrip()
    if s.startswith('"'):
        end = s.find('"', 1)
        return s[1:end]
    depth, j, instr = 0, 0, False
    while j < len(s):
        ch = s[j]
        if instr:
            if ch == "\\":
                j += 2
                continue
            if ch == '"':
                instr = False
        else:
            if ch == '"':
                instr = True
            elif ch in "[{":
                depth += 1
            elif ch in "]}":
                depth -= 1
                if depth == 0:
                    j += 1
                    break
            elif depth == 0 and ch == ";":
                break
        j += 1
    raw = s[:j]
    try:
        return json.loads(raw)
    except Exception:                       # noqa: BLE001
        return None


def fetch_fund(code, verbose=True):
    """抓取单只基金的净值与业绩数据。"""
    if code in _JS_CACHE:
        return _JS_CACHE[code]

    out = {"code": code, "ok": False}

    # ---- 2.1 最新净值 + 申赎状态
    try:
        u = ("https://api.fund.eastmoney.com/f10/lsjz?fundCode=%s&pageIndex=1&pageSize=1" % code)
        j = json.loads(http_get(u, referer="https://fundf10.eastmoney.com/"))
        it = (j.get("Data") or {}).get("LSJZList") or []
        if it:
            it = it[0]
            out["navDate"] = it.get("FSRQ", "")
            out["nav"] = num(it.get("DWJZ"))
            out["accNav"] = num(it.get("LJJZ"))
            out["dayChg"] = num(it.get("JZZZL"))
            out["sgzt"] = (it.get("SGZT") or "").strip()
            out["shzt"] = (it.get("SHZT") or "").strip()
    except Exception as e:                  # noqa: BLE001
        if verbose:
            log("  ! %s 净值接口异常：%s" % (code, e))

    # ---- 2.2 业绩 / 风险 / 规模 / 持仓结构
    try:
        txt = http_get("https://fund.eastmoney.com/pingzhongdata/%s.js" % code,
                       referer="https://fund.eastmoney.com/%s.html" % code)
        out["name"] = _js_var(txt, "fS_name") or ""
        out["fee"] = _js_var(txt, "fund_Rate") or ""
        out["srcFee"] = _js_var(txt, "fund_sourceRate") or ""
        out["sy_1y"] = num(_js_var(txt, "syl_1y"))       # 近1月
        out["sy_3y"] = num(_js_var(txt, "syl_3y"))       # 近3月
        out["sy_6y"] = num(_js_var(txt, "syl_6y"))       # 近6月
        out["sy_1n"] = num(_js_var(txt, "syl_1n"))       # 近1年

        trend = _js_var(txt, "Data_netWorthTrend") or []
        out["_trend"] = trend
        acc = _js_var(txt, "Data_ACWorthTrend") or []
        out["_acc"] = acc

        aa = _js_var(txt, "Data_assetAllocation") or {}
        out["asset"] = {}
        if isinstance(aa, dict):
            out["assetCats"] = aa.get("categories") or []
            for s in aa.get("series", []):
                out["asset"][s.get("name", "")] = s.get("data") or []

        fs = _js_var(txt, "Data_fluctuationScale") or {}
        if isinstance(fs, dict) and fs.get("series"):
            out["scaleNum"] = num(fs["series"][-1].get("y"))
            out["scaleMom"] = fs["series"][-1].get("mom", "")
            out["_scaleCats"] = fs.get("categories") or []
            out["_scaleVals"] = [num(s.get("y")) for s in fs["series"]]

        hs = _js_var(txt, "Data_holderStructure") or {}
        if isinstance(hs, dict):
            for s in hs.get("series", []):
                if "机构" in s.get("name", "") and s.get("data"):
                    out["instHold"] = num(s["data"][-1])
                    out["_hsCats"] = hs.get("categories") or []
                    out["_hsInst"] = [num(v) for v in s["data"]]
                if "个人" in s.get("name", "") and s.get("data"):
                    out["retailHold"] = num(s["data"][-1])

        rk = _js_var(txt, "Data_rateInSimilarPersent")
        if isinstance(rk, list) and rk:
            out["rank"] = num(rk[-1][1])

        mgr = _js_var(txt, "Data_currentFundManager") or []
        if isinstance(mgr, list) and mgr:
            out["mgrPerson"] = " / ".join([m.get("name", "") for m in mgr if m.get("name")])
            out["mgrStar"] = mgr[0].get("star")
            out["mgrTenure"] = mgr[0].get("workTime", "")

        out["ok"] = True
    except Exception as e:                  # noqa: BLE001
        if verbose:
            log("  ! %s 业绩接口异常：%s" % (code, e))

    _JS_CACHE[code] = out
    return out


def derive_metrics(f):
    """基于官方日增长率构造复权净值指数，计算区间收益与风险指标。"""
    trend = f.get("_trend") or []
    if not trend:
        return

    pts = []          # (date_str, 复权指数)
    idx = 1.0
    for p in trend:
        try:
            d = datetime.fromtimestamp(p["x"] / 1000).strftime("%Y-%m-%d")
        except Exception:                   # noqa: BLE001
            continue
        r = num(p.get("equityReturn"))
        idx *= (1 + r / 100.0)
        pts.append((d, idx))
    if not pts:
        return

    f["_rebased"] = pts
    last_d, last_v = pts[-1]

    def val_at(days_ago=None, ymd=None):
        if ymd:
            cands = [x for x in pts if x[0] <= ymd]
        else:
            tgt = (datetime.strptime(last_d, "%Y-%m-%d") - timedelta(days=days_ago)).strftime("%Y-%m-%d")
            cands = [x for x in pts if x[0] <= tgt]
        return cands[-1][1] if cands else None

    # 近1周 / 今年以来（官方接口未直接给出）
    w = val_at(days_ago=7)
    if w:
        f["sy_1w"] = (last_v / w - 1) * 100
    prev_year_end = "%d-12-31" % (int(last_d[:4]) - 1)
    y0 = val_at(ymd=prev_year_end)
    if y0:
        f["ytd"] = (last_v / y0 - 1) * 100

    # 近1年窗口风险指标
    start = "%d-%s" % (int(last_d[:4]) - 1, last_d[5:])
    win = [x for x in pts if x[0] >= start]
    if len(win) < 20:
        return
    f["winStart"] = win[0][0]

    peak, mdd = win[0][1], 0.0
    rets = []
    for i in range(1, len(win)):
        peak = max(peak, win[i][1])
        mdd = max(mdd, 1 - win[i][1] / peak)
        rets.append(win[i][1] / win[i - 1][1] - 1)
    f["mdd"] = mdd * 100

    if len(rets) > 5:
        mu = sum(rets) / len(rets)
        var = sum((x - mu) ** 2 for x in rets) / (len(rets) - 1)
        f["vol"] = math.sqrt(var) * math.sqrt(TRADING_DAYS) * 100
    else:
        f["vol"] = 0.0

    y1 = f.get("sy_1n", 0.0)
    f["y1"] = y1
    f["sharpe"] = (y1 - RF) / f["vol"] if f["vol"] > 1e-9 else 0.0
    f["calmar"] = y1 / f["mdd"] if f["mdd"] > 1e-9 else 0.0

    f["_win"] = win


def nav_series_for_chart(f, months=12, step_days=7):
    """近一年复权净值走势，按周抽样。"""
    pts = f.get("_rebased")
    if not pts:
        return []
    end = datetime.strptime(f["navDate"] or pts[-1][0], "%Y-%m-%d")
    start = end - timedelta(days=int(months * 30.5))
    sel = [x for x in pts if x[0] >= start.strftime("%Y-%m-%d")]
    if not sel:
        sel = pts[-40:]
    out, last_t = [], None
    for d, v in sel:
        t = datetime.strptime(d, "%Y-%m-%d")
        if last_t is None or (t - last_t).days >= step_days:
            out.append([d, round(v, 6)])
            last_t = t
    if out and out[-1][0] != sel[-1][0]:
        out.append([sel[-1][0], round(sel[-1][1], 6)])
    return out


# ================================================================== 3. 组装
def classify(rec, f):
    """判定产品分类、市场、流动性分组等。"""
    cat = str(pick(rec, "产品分类", "分类", default="") or "").strip()
    mkt = str(pick(rec, "基金市场", "市场", default="") or "").strip()
    atype = str(pick(rec, "资产类型", default="") or "").strip()
    code = str(pick(rec, "基金代码", default="") or "").strip()

    if not cat:
        if "ETF" in atype or mkt in ("上交所", "深交所"):
            cat = "债券ETF" if "债" in (f.get("name") or "") else "ETF"
        else:
            cat = "债券型"
    if not atype:
        atype = "ETF基金" if mkt in ("上交所", "深交所") else "开放式基金"
    return code, cat, mkt, atype


def liq_of(sgzt, shzt, name):
    """流动性分组。"""
    s = (sgzt or "") + "|" + (shzt or "")
    if "暂停申购" in sgzt and "暂停赎回" in shzt:
        return "lock", "封闭期不可申赎", "#d93f3f", "封闭期"
    if "暂停" in sgzt or "限制" in sgzt or "限大额" in sgzt:
        if "暂停赎回" in shzt:
            return "lock", "封闭期不可申赎", "#d93f3f", "封闭期"
        return "limit", "限大额申购·可赎回", "#d98a1f", "限大额申购"
    if "开放" in sgzt and "开放" in shzt:
        return "open", "每日开放申赎", "#1f9c62", "每日开放"
    if "场内" in s:
        return "open", "每日开放申赎", "#1f9c62", "场内可交易"
    return "other", (sgzt or "—") + "｜" + (shzt or "—"), "#8794a8", (sgzt or "—")


def build(nav_override, holdings_path, verbose=True):
    """核心：读表 + 取数 + 计算。返回 (header, holdings_df, funds, totals, meta)"""
    header, recs = read_holdings(holdings_path)
    log("读取持仓：%s，共 %d 条记录" % (os.path.basename(holdings_path), len(recs)))

    funds, closed, warnings = [], [], []

    for rec in recs:
        code = str(pick(rec, "基金代码", "代码", default="") or "").strip()
        if not code:
            continue
        name_excel = str(pick(rec, "名称", "基金名称", default="") or "").strip()
        shares_raw = num(pick(rec, "持仓份额", default=0))
        cost = num(pick(rec, "持仓成本(万元)", "持仓成本", default=0))
        realized = num(pick(rec, "已实现损益(万元)", "已实现损益", default=0))
        mv_excel = num(pick(rec, "全价市值(万元)", "市值", default=0))

        if shares_raw <= 0:
            closed.append({
                "code": code, "name": name_excel or code,
                "realized": realized, "cost": cost,
            })
            continue

        ov = (nav_override or {}).get(code) or {}
        f = fetch_fund(code, verbose=verbose) if not ov.get("nav") else {"code": code, "ok": True}
        for k, v in ov.items():
            f[k] = v
        derive_metrics(f)

        c, cat, mkt, atype = classify(rec, f)
        # 名称优先沿用持仓表中的正式全称，缺失时回落到接口名称
        name = name_excel or f.get("name") or code
        nav = num(f.get("nav"))
        if nav <= 0:
            warnings.append("%s(%s) 未取到净值，沿用表内市值/盈亏" % (name, code))
            nav = (mv_excel * 10000.0 / shares_raw) if shares_raw else 0.0
        acc = num(f.get("accNav"), nav)

        mv = shares_raw * nav / 10000.0
        pnl = mv - cost if cost else 0.0
        ret = (pnl / cost * 100.0) if cost else 0.0
        unit_cost = num(pick(rec, "单位持仓成本(元)", "单位成本", default=0)) or (cost * 10000.0 / shares_raw if shares_raw else 0)

        sgzt, shzt = f.get("sgzt", ""), f.get("shzt", "")
        if not sgzt and not shzt:
            sgzt = str(pick(rec, "申购状态", default="") or "")
            shzt = str(pick(rec, "赎回状态", default="") or "")
        lk, ltxt, lcolor, lshort = liq_of(sgzt, shzt, name)

        asset = f.get("asset") or {}
        stock = _lastv(asset.get("股票占净比"))
        bond = _lastv(asset.get("债券占净比"))
        cash = _lastv(asset.get("现金占净比"))
        other = max(0.0, 100.0 - stock - bond - cash) if (stock or bond or cash) else 0.0

        funds.append({
            "code": code, "name": name,
            "short": _short_name(name, code),
            "cls": cat, "mk": mkt, "acct": str(pick(rec, "内部证券账户", default="FOF投资账户") or "FOF投资账户"),
            "atype": atype,
            "sharesRaw": shares_raw, "shares": round(shares_raw / 10000.0, 6),
            "cost": cost, "navcost": unit_cost, "mv": mv, "pnl": pnl, "ret": ret,
            "realized": realized,
            "nav": nav, "acc": acc, "navDate": f.get("navDate", ""), "dayChg": num(f.get("dayChg")),
            "mvExcel": mv_excel,
            "w1": num(f.get("sy_1w")), "m1": num(f.get("sy_1y")), "m3": num(f.get("sy_3y")),
            "m6": num(f.get("sy_6y")), "ytd": num(f.get("ytd")), "y1": num(f.get("y1")),
            "vol": num(f.get("vol")), "mdd": num(f.get("mdd")),
            "sharpe": num(f.get("sharpe")), "calmar": num(f.get("calmar")),
            "rank": num(f.get("rank")),
            "scale": (fmt(f.get("scaleNum"), 2) + "亿") if f.get("scaleNum") else "—",
            "scaleNum": num(f.get("scaleNum")),
            "scaleMom": f.get("scaleMom", ""),
            "fee": (f.get("fee") or "—") + "%" if f.get("fee") not in (None, "", "—") else "—",
            "status": (sgzt or "—") + "｜" + (shzt or "—"),
            "sgzt": sgzt, "shzt": shzt,
            "liqGroup": lk, "liqText": ltxt, "liqColor": lcolor, "liqShort": lshort,
            "stock": stock, "bond": bond, "deposit": cash, "other": other,
            "instHold": num(f.get("instHold")), "retailHold": num(f.get("retailHold")),
            "mgrPerson": f.get("mgrPerson") or "—",
            "mgrStar": f.get("mgrStar"),
            "inception": str(pick(rec, "成立日期", default="") or "—"),
            "mgr": _mgr_of(name, code),
            "top5": None, "top5name": "—",
            "_rebased": f.get("_rebased"),
        })
        # 规模变动趋势 + 机构持有比例（按报告期对齐）
        sc_c = f.get("_scaleCats") or []
        sc_v = f.get("_scaleVals") or []
        inst_map = dict(zip(f.get("_hsCats") or [], f.get("_hsInst") or []))
        inst = [inst_map.get(c) for c in sc_c]
        funds[-1]["scaleTrend"] = {
            "cats": sc_c, "vals": sc_v,
            "inst": [None if v is None else round(v, 2) for v in inst],
        } if sc_c else {"cats": [], "vals": [], "inst": []}

    total_cost = sum(x["cost"] for x in funds)
    total_mv = sum(x["mv"] for x in funds)
    total_pnl = sum(x["pnl"] for x in funds)
    total_real = sum(x["realized"] for x in funds) + sum(x["realized"] for x in closed)

    # 统一数值精度，避免报告里出现长小数
    _R = {"cost": 2, "mv": 2, "pnl": 4, "realized": 4, "ret": 4, "weight": 4,
          "nav": 4, "acc": 4, "dayChg": 4, "mvExcel": 2,
          "w1": 4, "m1": 4, "m3": 4, "m6": 4, "ytd": 4, "y1": 4,
          "vol": 4, "mdd": 4, "sharpe": 4, "calmar": 4, "rank": 2,
          "stock": 2, "bond": 2, "deposit": 2, "other": 2,
          "instHold": 2, "retailHold": 2, "scaleNum": 2}
    for x in funds:
        for k, nd in _R.items():
            if k in x:
                x[k] = round(num(x[k]), nd)
        x["navcost"] = round(x["navcost"], 4)
        x["shares"] = round(x["shares"], 6)
    for x in funds:
        x["weight"] = (x["mv"] / total_mv * 100.0) if total_mv else 0.0

    totals = {
        "cost": total_cost, "mv": total_mv, "pnl": total_pnl,
        "ret": (total_pnl / total_cost * 100.0) if total_cost else 0.0,
        "realized": total_real,
        "count": len(funds),
        "profitN": len([x for x in funds if x["pnl"] > 0]),
        "lossN": len([x for x in funds if x["pnl"] < 0]),
    }
    if warnings:
        for w in warnings:
            log("  ! " + w)

    vals = [x["navDate"] for x in funds if x["navDate"]]
    meta = {
        "valDate": max(vals) if vals else datetime.now().strftime("%Y-%m-%d"),
        "navDate": max(vals) if vals else "—",
        "genTime": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "accounts": len(set(x["acct"] for x in funds)),
        "warnings": warnings,
        "closed": closed,
    }
    return header, recs, funds, totals, meta


def _lastv(arr):
    if isinstance(arr, list) and arr:
        v = arr[-1]
        if isinstance(v, (int, float)):
            return float(v)
    return 0.0


def _short_name(name, code):
    """把长基金名压缩成图例用短名（保留 ETF 等后缀）。"""
    s = re.sub(r"（.*?）|\(.*?\)", "", name)
    s = re.sub(r"(证券投资基金|发起式|发起)", "", s).strip()
    if len(s) > 16:
        s = s[:9] + "…" + s[-6:]
    return s


_MGR_MAP = {}
_MGR_FIX = {"渤海汇金": "渤海汇金资管"}


def _mgr_of(name, code):
    if code in _MGR_MAP:
        return _MGR_MAP[code]
    m = re.search(r"^(易方达|海富通|中金|渤海汇金|华夏|南方|广发|招商|博时|富国|工银|嘉实|汇添富|鹏华|银华|国泰|华安|万家|中银|兴全|宝盈|长城|东方|华宝|融通|国投|天弘|永赢|平安|鑫元|景顺长城|交银|建信|农银|民生|浦银|信达澳亚|中欧|华商|诺安|长盛|大成|国海富兰克林|汇安|中加|德邦|恒生前海|惠升|太平|国寿安保|大家|泰康)", name)
    if not m:
        return "—"
    base = m.group(1)
    return _MGR_FIX.get(base, base + "基金")


# ================================================================== 4. 历史快照 & 变动
def hist_dir(outdir):
    d = os.path.join(outdir, HIST_SUBDIR)
    os.makedirs(d, exist_ok=True)
    return d


def load_snapshots(outdir):
    d = hist_dir(outdir)
    out = []
    for fn in sorted(os.listdir(d)):
        if fn.startswith("snapshot_") and fn.endswith(".json"):
            try:
                with open(os.path.join(d, fn), encoding="utf-8") as fh:
                    out.append(json.load(fh))
            except Exception:               # noqa: BLE001
                pass
    return out


def save_snapshot(outdir, funds, totals, meta):
    d = hist_dir(outdir)
    snap = {
        "date": meta["valDate"],
        "generated": meta["genTime"],
        "totals": totals,
        "holdings": {x["code"]: {
            "name": x["name"], "shares": x["sharesRaw"], "cost": x["cost"],
            "nav": x["nav"], "mv": x["mv"], "pnl": x["pnl"], "ret": x["ret"],
        } for x in funds},
    }
    p = os.path.join(d, "snapshot_%s_%s.json" % (
        meta["valDate"].replace("-", ""), datetime.now().strftime("%H%M%S")))
    with open(p, "w", encoding="utf-8") as fh:
        json.dump(snap, fh, ensure_ascii=False, indent=2)
    return p, snap


def diff_snapshots(prev, funds, closed, totals):
    """与上一期快照对比，生成持仓变动。"""
    if not prev:
        return None
    prev_h = prev.get("holdings") or {}
    cur_h = {x["code"]: x for x in funds}
    items, rows = [], []
    allc = list(dict.fromkeys(list(prev_h.keys()) + list(cur_h.keys())))

    for c in allc:
        p, u = prev_h.get(c), cur_h.get(c)
        nm = (u or p).get("name", c)
        short = _short_name(nm, c)
        if p and not u:
            cl = next((x for x in closed if x["code"] == c), None)
            items.append({"code": c, "name": nm, "short": short, "type": "清仓",
                          "prevShares": p["shares"] / 10000.0, "curShares": 0.0,
                          "dShares": -p["shares"] / 10000.0,
                          "dMv": -(p.get("mv") or 0),
                          "note": ("已实现损益 %s 万元" % fmt(cl["realized"])) if cl else "本期已清仓"})
            rows.append(items[-1])
        elif u and not p:
            items.append({"code": c, "name": nm, "short": short, "type": "新增",
                          "prevShares": 0.0, "curShares": u["shares"],
                          "dShares": u["shares"], "dMv": u["mv"],
                          "note": "本期新增建仓"})
            rows.append(items[-1])
        else:
            ds = u["shares"] - p["shares"] / 10000.0
            dmv = u["mv"] - (p.get("mv") or 0)
            if abs(u["shares"] - p["shares"] / 10000.0) < 0.005:
                items.append({"code": c, "name": nm, "short": short, "type": "持有未变",
                              "prevShares": p["shares"] / 10000.0, "curShares": u["shares"],
                              "dShares": 0.0, "dMv": dmv, "note": "份额未变，市值随净值波动"})
                rows.append(items[-1])
            else:
                t = "增持" if ds > 0 else "减持"
                items.append({"code": c, "name": nm, "short": short, "type": t,
                              "prevShares": p["shares"] / 10000.0, "curShares": u["shares"],
                              "dShares": ds, "dMv": dmv,
                              "note": "份额 %s %s 万份" % ("增加" if ds > 0 else "减少", fmt(abs(ds), 2))})
                rows.append(items[-1])

    changed = [x for x in items if x["type"] != "持有未变"]
    return {
        "prevDate": prev.get("date", "—"),
        "curDate": None,
        "items": items,
        "rows": rows,
        "changed": changed,
        "dPnl": totals["pnl"] - (prev.get("totals") or {}).get("pnl", 0),
        "dMv": totals["mv"] - (prev.get("totals") or {}).get("mv", 0),
        "newN": len([x for x in changed if x["type"] == "新增"]),
        "addN": len([x for x in changed if x["type"] == "增持"]),
        "cutN": len([x for x in changed if x["type"] == "减持"]),
        "outN": len([x for x in changed if x["type"] == "清仓"]),
    }


# ================================================================== 5. 报告 HTML
def _kpi_html(totals, changes):
    def k(k_, v_, u_, cls=""):
        return ('<div class="kpi"><div class="k">%s</div>'
                '<div class="v %s">%s<span class="u">%s</span></div></div>' % (k_, cls, v_, u_))

    cls = "up" if totals["pnl"] >= 0 else "down"
    h = k("组合总市值", fmt(totals["mv"]), "万元")
    h += k("持仓总成本", fmt(totals["cost"]), "万元")
    h += k("累计浮动盈亏", fmt(totals["pnl"], 2, True), "万元", cls)
    h += k("组合账面收益率", fmt(totals["ret"], 2, True), "%", cls)
    h += k("盈利 / 亏损基金", "%d / %d" % (totals["profitN"], totals["lossN"]), "只")
    return h


def _changes_section(changes, totals):
    if not changes or not changes["changed"]:
        return ""
    c = changes
    tc = {"新增": "#tag-b", "增持": "#tag-r", "减持": "#tag-g", "清仓": "#tag-x"}
    rows = ""
    for it in c["changed"]:
        cls = {"新增": "tag-b", "增持": "tag-r", "减持": "tag-g", "清仓": "tag-x"}[it["type"]]
        dsc = (RED if it["dShares"] >= 0 else GREEN)
        rows += ('<tr><td class="l"><b>%s</b><div class="code">%s</div></td>'
                 '<td class="l"><span class="tag %s">%s</span></td>'
                 '<td>%s</td><td>%s</td>'
                 '<td style="color:%s;font-weight:700">%s</td>'
                 '<td>%s</td><td class="l" style="font-family:inherit;font-size:12px">%s</td></tr>'
                 % (it["name"], it["code"], cls, it["type"],
                    fmt(it["prevShares"], 2), fmt(it["curShares"], 2),
                    dsc, fmt(it["dShares"], 2, True), fmt(it["dMv"], 2, True), it["note"]))

    tags = []
    if c["newN"]:
        tags.append("新增 %d 只" % c["newN"])
    if c["addN"]:
        tags.append("增持 %d 只" % c["addN"])
    if c["cutN"]:
        tags.append("减持 %d 只" % c["cutN"])
    if c["outN"]:
        tags.append("清仓 %d 只" % c["outN"])
    tagstr = "、".join(tags) if tags else "仅有净值波动"

    return """
  <section>
    <div class="sec-head"><div class="sec-no" style="background:#c93636">Δ</div><h2>本期持仓变动</h2>
      <span class="hint">与上期快照 %s 对比 · %s</span></div>
    <div class="grid2-13">
      <div class="card">
        <h3>份额与市值变动</h3>
        <div class="desc">柱=份额变动（万份，左轴）｜折线=市值变动（万元，右轴）</div>
        <div class="chart" id="c_change"></div>
      </div>
      <div class="card">
        <h3>变动明细</h3>
        <div class="desc">新增 / 增持 / 减持 / 清仓 逐笔对照</div>
        <table>
          <thead><tr><th class="l">基金</th><th class="l">变动类型</th><th>上期份额</th><th>本期份额</th>
          <th>份额变动</th><th>市值变动</th><th class="l">备注</th></tr></thead>
          <tbody>%s</tbody>
        </table>
      </div>
    </div>
    <div class="insight warn">
      <b>本期变动：</b>%s。组合总市值较上期变动 <b>%s 万元</b>，累计浮动盈亏变动 <b>%s 万元</b>。
      持仓调整后需重新评估集中度与流动性约束（详见第 2 节）。
    </div>
  </section>
""" % (c["prevDate"], tagstr, rows, tagstr, fmt(c["dMv"], 2, True), fmt(c["dPnl"], 2, True))


def _tbl_hold(funds, totals):
    r = ""
    for f in sorted(funds, key=lambda x: -x["mv"]):
        c = RED if f["pnl"] >= 0 else GREEN
        r += ('<tr><td class="l"><b>%s</b><div class="code">%s · %s</div></td>'
              '<td>%s</td><td>%s</td><td>%s</td>'
              '<td style="color:%s;font-weight:700">%s</td>'
              '<td style="color:%s;font-weight:700">%s</td><td><b>%.2f%%</b></td></tr>'
              % (f["name"], f["code"], f["cls"], fmt(f["shares"], 2), fmt(f["cost"], 2),
                 fmt(f["mv"], 2), c, fmt(f["pnl"], 4, True), c, pct(f["ret"], 4), f["weight"]))
    c = RED if totals["pnl"] >= 0 else GREEN
    r += ('<tr class="total"><td class="l">合计（%d 只）</td><td>%s</td><td>%s</td><td>%s</td>'
          '<td style="color:%s">%s</td><td style="color:%s">%s</td><td>100.00%%</td></tr>'
          % (len(funds), fmt(sum(x["shares"] for x in funds), 2), fmt(totals["cost"]),
             fmt(totals["mv"]), c, fmt(totals["pnl"], 4, True), c, pct(totals["ret"], 4)))
    return r


def _insight_scale(funds, totals, meta):
    d = sorted(funds, key=lambda x: -x["mv"])
    top = d[0]
    w2 = top["weight"] + d[1]["weight"] if len(d) > 1 else top["weight"]
    cats = sorted({x["cls"] for x in funds})
    return (
        "<b>规模特征：</b>组合合计市值 <b>%s 万元</b>，共持有 <b>%d 只</b>基金，覆盖 %s 品类。"
        "规模分布%s：单只 <b>%s（%s）</b> 一只即占 <b>%.2f%%</b>%s。"
        % (fmt(totals["mv"]), len(funds), "、".join(cats),
           "极不均衡" if top["weight"] > 35 else "相对均衡",
           top["name"], top["code"], top["weight"],
           "，前两大合计占 %.2f%%" % w2 if len(d) > 1 else ""))


def _insight_struct(funds, totals):
    def grp(key):
        m = {}
        for f in funds:
            m[f[key]] = m.get(f[key], 0) + f["mv"]
        return sorted(m.items(), key=lambda x: -x[1])

    byacc = grp("acct") if len(set(x["acct"] for x in funds)) > 1 else grp("mk")
    parts = []
    parts.append('<div class="insight" style="margin-top:0"><b>① 品种结构：</b>%s。%s</div>' % (
        "、".join("%s %.2f%%" % (k, v / totals["mv"] * 100) for k, v in grp("cls")),
        "无股票型 / 混合型工具，权益敞口仅来自含权债基。" if all(x["stock"] < 3 for x in funds)
        else "存在含权品种，权益敞口见下。"))

    top = max(funds, key=lambda x: x["weight"])
    parts.append('<div class="insight %s"><b>② 集中度：</b>第一大持仓 %s 占 <b>%.2f%%</b>%s。</div>' % (
        "risk" if top["weight"] > 40 else "warn", top["short"], top["weight"],
        "，单一标的对组合影响显著，建议设置单只上限" if top["weight"] > 40 else "，集中度尚在可控范围"))

    lock = [x for x in funds if x["liqGroup"] == "lock"]
    lockmv = sum(x["mv"] for x in lock)
    parts.append('<div class="insight %s"><b>③ 流动性：</b>可随时申赎资产占 <b>%.2f%%</b>；%s</div>' % (
        "risk" if lockmv / totals["mv"] > 0.3 else "warn",
        sum(x["mv"] for x in funds if x["liqGroup"] == "open") / totals["mv"] * 100,
        ("已封闭不可变现的持仓占 %.2f%%（%s）" % (lockmv / totals["mv"] * 100,
                                          "、".join(x["short"] for x in lock))) if lock else "无封闭期品种"))

    eq = [x for x in funds if x["stock"] > 1]
    parts.append('<div class="insight %s"><b>④ 权益敞口：</b>%s</div>' % (
        "risk" if eq and max(x["stock"] for x in eq) > 10 else "",
        ("<b>%s</b> 股票市值占净值 <b>%.2f%%</b>，为组合主要的权益风险来源；其余品种股票仓位为 0。"
         % (eq[0]["short"], eq[0]["stock"])) if eq else "全部持仓股票仓位低于 1%，组合为纯固收结构。"))
    return "".join(parts)


def _insight_pnl(funds, totals):
    win = sorted([x for x in funds if x["pnl"] > 0], key=lambda x: -x["pnl"])
    los = sorted([x for x in funds if x["pnl"] < 0], key=lambda x: x["pnl"])
    tw = sum(x["pnl"] for x in win)
    tl = sum(x["pnl"] for x in los)
    parts = []
    if win:
        parts.append('<div class="insight good" style="margin-top:0"><b>盈利端：</b>%s，合计 <b>%s 万元</b>。%s</div>' % (
            "、".join("<b>%s</b> %s 万元（%s）" % (x["short"], fmt(x["pnl"], 2, True), pct(x["ret"], 2))
                      for x in win),
            fmt(tw, 2, True),
            ("利润高度集中于 <b>%s</b>（占全部正收益的 %.1f%%），盈利集中度风险明显。"
             % (win[0]["short"], win[0]["pnl"] / tw * 100)) if len(win) > 1 and win[0]["pnl"] / tw > 0.6 else ""))
    if los:
        parts.append('<div class="insight risk"><b>亏损端：</b>%s，合计 <b>%s 万元</b>，抵消了盈利的 %.1f%%。</div>' % (
            "、".join("<b>%s</b> %s 万元（%s）" % (x["short"], fmt(x["pnl"], 2, True), pct(x["ret"], 2))
                      for x in los),
            fmt(tl, 2, True), abs(tl) / tw * 100 if tw else 0))
    parts.append('<div class="insight"><b>净结果：</b>组合累计浮动盈亏 <b>%s 万元</b>，整体账面收益率 <b>%s</b>；'
                 '已实现损益 <b>%s 万元</b>。%s</div>' % (
                     fmt(totals["pnl"], 2, True), pct(totals["ret"], 2),
                     fmt(totals["realized"], 2, True),
                     "账面盈亏为浮动口径，未考虑申赎费用。" ))
    return "".join(parts)


def _insight_perf(funds, meta):
    if not funds:
        return ""
    d = sorted(funds, key=lambda x: -x["y1"])
    best, worst = d[0], d[-1]
    max_vol = max(x["vol"] for x in funds)
    max_mdd = max(x["mdd"] for x in funds)
    min_sh = min(x["sharpe"] for x in funds)
    parts = []

    parts.append('<div class="insight good" style="margin-top:18px">'
                 '<b>收益最高 · %s（%s）：</b>近1年 <b>%s</b>、今年以来 <b>%s</b>，'
                 '近3月 %s、近6月 %s，最大回撤 %.2f%%，Sharpe %.2f、Calmar %.2f。</div>' % (
                     best["name"], best["code"], pct(best["y1"], 2), pct(best["ytd"], 2),
                     pct(best["m3"], 2), pct(best["m6"], 2), best["mdd"], best["sharpe"], best["calmar"]))

    if len(d) > 2:
        parts.append('<div class="insight"><b>表现居中：</b>%s。</div>' % (
            "；".join("<b>%s</b> 近1年 %s、近1月 %s，最大回撤 %.2f%%"
                      % (x["short"], pct(x["y1"], 2), pct(x["m1"], 2), x["mdd"]) for x in d[1:-1])))

    # 风险调整收益最弱（Sharpe 最低）
    weak = min(funds, key=lambda x: x["sharpe"])
    flags = []
    if abs(weak["vol"] - max_vol) < 1e-6:
        flags.append("年化波动率 %.2f%% 为组合最高" % weak["vol"])
    if abs(weak["mdd"] - max_mdd) < 1e-6:
        flags.append("最大回撤 %.2f%% 为组合最高" % weak["mdd"])
    negs = [p for p in ("近1月 %s" % pct(weak["m1"], 2), "近3月 %s" % pct(weak["m3"], 2),
                        "近6月 %s" % pct(weak["m6"], 2)) if "-" in p]
    parts.append('<div class="insight %s"><b>风险调整最弱 · %s（%s）：</b>Sharpe 仅 <b>%.2f</b>、Calmar %.2f，'
                 '近1年 %s。%s%s</div>' % (
                     "risk" if weak["sharpe"] < 0.8 else "warn",
                     weak["name"], weak["code"], weak["sharpe"], weak["calmar"], pct(weak["y1"], 2),
                     ("多区间收益为负（%s）。" % "、".join(negs)) if negs else "",
                     ("；".join(flags) + "。" if flags else "")))

    nneg = len([x for x in funds if x["m1"] < 0])
    parts.append('<div class="insight warn"><b>短期动能：</b>%d 只基金中 <b>%d 只近 1 月为负</b>，'
                 '组合短期收益动能%s。</div>' % (
                     len(funds), nneg, "偏弱，建议提高跟踪频率" if nneg * 2 > len(funds) else "尚可"))
    return "".join(parts)


def _fcards(funds):
    r = ""
    for f in sorted(funds, key=lambda x: -x["weight"]):
        pc = RED if f["pnl"] >= 0 else GREEN
        rc = RED if f["ret"] >= 0 else GREEN
        sc = "#1f9c62" if f["sharpe"] >= 2 else ("#2a6ba8" if f["sharpe"] >= 1.4 else "#d93f3f")
        stcl = {"tag-g": "#1f7d51", "tag-r": "#b93b3b"}.get(
            {"open": "tag-g", "lock": "tag-r"}.get(f["liqGroup"], "tag-x"), "#a86a10")
        sttag = {"open": "tag-g", "lock": "tag-r"}.get(f["liqGroup"], "tag-x")
        barw = min(100, max(3, f["weight"]))
        r += '<div class="fcard">' + \
             '<div class="fh"><div><div class="fname">%s</div>' % f["name"] + \
             '<div class="fmeta">%s · %s · %s</div>' % (f["code"], f["cls"], f["mk"]) + \
             '<div style="margin-top:6px"><span class="tag %s">%s</span> <span class="tag %s">%s</span>' \
             '<span class="tag tag-x">管理费 %s</span>%s</div></div>' % (
                 "tag-b" if "ETF" in f["cls"] else "tag-e", f["cls"], sttag, f["liqShort"], f["fee"],
                 ('<span class="tag tag-x" style="margin-left:4px">机构持有 %.1f%%</span>' % f["instHold"])
                 if f["instHold"] else "") + \
             '<div style="text-align:right"><div class="fret" style="color:%s">%s</div>' % (rc, pct(f["ret"], 2)) + \
             '<div style="font-size:11px;color:#8794a8">账面收益率</div></div></div>' + \
             '<div class="frow"><span class="fl">持仓市值 / 组合权重</span><span class="fv">%s 万元 · %.2f%%</span></div>' % (fmt(f["mv"]), f["weight"]) + \
             '<div class="progbar"><i style="width:%.1f%%;background:%s"></i></div>' % (barw, "#2a6ba8" if f["weight"] > 40 else "#7bb6e3") + \
             '<div class="frow" style="margin-top:8px"><span class="fl">持仓成本 / 单位成本</span><span class="fv">%s 万元 · %.4f</span></div>' % (fmt(f["cost"]), f["navcost"]) + \
             '<div class="frow"><span class="fl">浮动盈亏</span><span class="fv" style="color:%s">%s 万元</span></div>' % (pc, fmt(f["pnl"], 4, True)) + \
             '<div class="frow"><span class="fl">最新单位净值 / 累计净值</span><span class="fv">%.4f / %.4f <span style="color:#8794a8">(%s)</span></span></div>' % (f["nav"], f["acc"], f["navDate"]) + \
             '<div class="frow"><span class="fl">基金规模 / 规模环比</span><span class="fv">%s · %s</span></div>' % (f["scale"], f["scaleMom"] or "—") + \
             '<div class="frow"><span class="fl">基金经理</span><span class="fv" style="font-family:inherit;font-weight:500">%s</span></div>' % f["mgrPerson"] + \
             '<div class="frow"><span class="fl">近1年 / 今年以来</span><span class="fv" style="color:%s">%s / %s</span></div>' % (
                 RED if f["y1"] >= 0 else GREEN, pct(f["y1"], 2), pct(f["ytd"], 2)) + \
             '<div class="frow"><span class="fl">近1月 / 近3月</span><span class="fv" style="color:%s">%s / %s</span></div>' % (
                 RED if f["m1"] >= 0 else GREEN, pct(f["m1"], 4), pct(f["m3"], 4)) + \
             '<div class="frow"><span class="fl">最大回撤 / 年化波动率</span><span class="fv">%.4f%% / %.4f%%</span></div>' % (f["mdd"], f["vol"]) + \
             '<div class="frow"><span class="fl">Sharpe / Calmar</span><span class="fv" style="color:%s">%.4f / %.4f</span></div>' % (sc, f["sharpe"], f["calmar"]) + \
             '<div class="frow"><span class="fl">近6月 / 今年以来</span><span class="fv" style="color:%s">%s / %s</span></div>' % (
                 RED if f["m6"] >= 0 else GREEN, pct(f["m6"], 2), pct(f["ytd"], 2)) + \
             '<div class="frow"><span class="fl">资产配置（股/债/现金）</span><span class="fv">%.2f%% / %.2f%% / %.2f%%</span></div>' % (f["stock"], f["bond"], f["deposit"]) + \
             '<div class="frow"><span class="fl">申赎状态</span><span class="fv" style="color:%s">%s</span></div>' % (stcl, f["status"]) + \
             '</div>'
    return r


def _act_table(funds, totals, meta, changes):
    rows = []

    weak = min(funds, key=lambda x: x["m1"]) if funds else None
    if weak and (weak["m1"] < 0 or weak["sharpe"] < 0.8):
        rows.append(("P0", "%s（%s）近1月 %s、近3月 %s，Sharpe %.2f、最大回撤 %.2f%% 为组合最弱" % (
            weak["name"], weak["code"], pct(weak["m1"], 2), pct(weak["m3"], 2), weak["sharpe"], weak["mdd"]),
            "核查其权益/转债仓位与行业暴露，判断是市场性回撤还是策略偏移；结合持有期约束决定是否在开放窗口减仓或替换"))

    top = max(funds, key=lambda x: x["weight"]) if funds else None
    if top and top["weight"] > 35:
        rows.append(("P0" if top["weight"] > 45 else "P1",
                     "%s 占组合 %.2f%%，单一标的集中度偏高%s" % (
                         top["name"], top["weight"],
                         "，且处于封闭期无法赎回" if top["liqGroup"] == "lock" else ""),
                     ("登记其下一个开放期，提前规划流动性；开放日评估是否将占比降至 30%% 以内"
                      if top["liqGroup"] == "lock" else "评估逐步分散至其他管理人/策略，降低单一产品依赖")))

    lock = [x for x in funds if x["liqGroup"] == "lock"]
    if lock and sum(x["mv"] for x in lock) / totals["mv"] > 0.25:
        rows.append(("P1", "封闭期不可申赎资产占 %.2f%%（%s），组合流动性受限" % (
            sum(x["mv"] for x in lock) / totals["mv"] * 100, "、".join(x["short"] for x in lock)),
            "建立开放期日历提醒；预留一定比例高流动性 ETF 作为组合缓冲垫"))

    win = sorted([x for x in funds if x["pnl"] > 0], key=lambda x: -x["pnl"])
    if win and sum(x["pnl"] for x in win) > 0 and win[0]["pnl"] / sum(x["pnl"] for x in win) > 0.6:
        rows.append(("P1", "组合正收益高度依赖单只基金——%s 贡献正收益的 %.1f%%" % (
            win[0]["short"], win[0]["pnl"] / sum(x["pnl"] for x in win) * 100),
            "评估是否部分了结或再平衡，避免盈亏同源；关注该品种久期与利率敏感性"))

    noeq = all(x["stock"] < 1 for x in funds)
    if noeq:
        rows.append(("P1", "组合为纯固收结构（股票敞口≈0），收益弹性有限",
                     "若组合定位允许，可考虑加入少量含权或转债策略以提升长期收益中枢；若定位低波动固收，则维持现状并明确收益预期"))

    hi = [x for x in funds if num(x["fee"].replace("%", "")) >= 0.3]
    if hi and len(funds) > len(hi):
        rows.append(("P2", "%s 管理费率 %s，高于组合内 ETF 的 0.15%%" % (
            "、".join(x["short"] for x in hi), "、".join(x["fee"] for x in hi)),
            "跟踪费后收益差异，评估主动管理是否带来对应超额，必要时用低费率指数化工具替代"))

    nneg = len([x for x in funds if x["m1"] < 0])
    if funds and nneg * 2 > len(funds):
        rows.append(("P2", "近1月 %d/%d 只基金为负，组合短期动能减弱" % (nneg, len(funds)),
                     "建立月度净值跟踪表，对连续 2 个月跑输同类中位数的品种启动替换评估"))

    if changes and changes["changed"]:
        rows.append(("P2", "本期发生持仓调整：%s" % "、".join(
            "%s %s" % (x["short"], x["type"]) for x in changes["changed"]),
            "调整后复核单只权重上限、管理人分散度与整体久期分布，更新组合目标权重表"))

    rows.append(("P2", "缺少建仓日期，无法计算组合年化收益与持有期表现",
                 "补充每笔买入日期与金额，形成完整的成本-时间序列，便于后续归因与再平衡决策"))

    cls = {"P0": "tag-r", "P1": "tag-e", "P2": "tag-x"}
    r = ""
    for p, t, a in rows:
        r += ('<tr><td class="l"><span class="tag %s">%s</span></td>'
              '<td class="l" style="line-height:1.65">%s</td>'
              '<td class="l" style="line-height:1.65;color:#4a5c73;font-size:12.5px">%s</td></tr>'
              % (cls[p], p, t, a))
    return r


def render_html(funds, totals, meta, changes, out_path, template_path=TEMPLATE_PATH):
    with open(template_path, encoding="utf-8") as fh:
        tpl = fh.read()
    with open(ECHARTS_PATH, encoding="utf-8") as fh:
        ec = fh.read()

    navseries = []
    palette = ["#2a6ba8", "#7bb6e3", "#1f9c62", "#d98a1f", "#8a5cd6", "#c93636"]
    for i, f in enumerate(sorted(funds, key=lambda x: -x["mv"])):
        s = nav_series_for_chart(f)
        if s:
            navseries.append({"name": "%s(%s)" % (f["short"], f["code"]),
                              "color": palette[i % len(palette)], "d": s})

    data = {
        # 剔除以下划线开头的内部字段，避免把全量净值历史塞进报告
        "funds": [{k: v for k, v in x.items() if not k.startswith("_")} for x in funds],
        "total": totals,
        "navseries": navseries,
        "changes": changes,
        "valDate": meta["valDate"],
    }

    note = ("<b>数据说明与局限性：</b>持仓成本、份额与已实现损益取自您的持仓表"
            "（估值日 %s）；单位净值、区间收益、风险指标、规模与资产配置取自公开基金数据接口"
            "（最新净值日 %s）。"
            "风险指标为近 1 年窗口、按官方日增长率复权后计算：波动率为日频年化，"
            "Sharpe 采用无风险利率 %.1f%%，Calmar = 近1年收益 / 最大回撤。"
            "该表<b>未包含建仓日期</b>，因此「账面收益率」为累计浮动收益率，"
            "<b>无法换算为年化收益或与市场基准做同期对比</b>。") % (
                meta["valDate"], meta["navDate"], RF)

    reps = {
        "{{VAL_DATE}}": meta["valDate"],
        "{{NAV_DATE}}": meta["navDate"],
        "{{GEN_TIME}}": meta["genTime"],
        "{{HOLD_COUNT}}": str(totals["count"]),
        "{{ACCT_COUNT}}": str(meta["accounts"]),
        "{{PERIOD_LABEL}}": meta.get("periodLabel", "周度跟踪"),
        "{{KPI_HTML}}": _kpi_html(totals, changes),
        "{{CHANGES_SECTION}}": _changes_section(changes, totals),
        "{{TBL_HOLD}}": _tbl_hold(funds, totals),
        "{{INSIGHT_SCALE}}": _insight_scale(funds, totals, meta),
        "{{INSIGHT_STRUCT}}": _insight_struct(funds, totals),
        "{{INSIGHT_PNL}}": _insight_pnl(funds, totals),
        "{{INSIGHT_PERF}}": _insight_perf(funds, meta),
        "{{FCARDS}}": _fcards(funds),
        "{{ACT_TABLE}}": _act_table(funds, totals, meta, changes),
        "{{DATA_NOTE}}": note,
        "/*__ECHARTS__*/": ec,
        "/*__DATA__*/": js(data),
    }
    for k, v in reps.items():
        tpl = tpl.replace(k, v)

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as fh:
        fh.write(tpl)
    return out_path


# ================================================================== 6. Excel 输出
def write_excel(funds, totals, meta, out_path):
    if openpyxl is None:
        log("! 缺少 openpyxl，跳过 Excel 输出")
        return None
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "持仓投后更新"

    cols = [
        ("基金代码", 10), ("名称", 30), ("产品分类", 10), ("基金市场", 12), ("资产类型", 12),
        ("内部证券账户", 14), ("持仓份额", 18), ("持仓份额(万)", 14),
        ("持仓成本(万元)", 15), ("单位持仓成本(元)", 15),
        ("最新单位净值", 13), ("净值日期", 12), ("全价市值(万元)", 15),
        ("浮动盈亏(万元)", 15), ("总盈亏(万元)", 15), ("已实现损益(万元)", 16),
        ("账面收益率(%)", 14), ("市值权重(%)", 13),
        ("近1周(%)", 11), ("近1月(%)", 11), ("近3月(%)", 11), ("近6月(%)", 11),
        ("今年以来(%)", 12), ("近1年(%)", 11),
        ("最大回撤(%)", 12), ("年化波动率(%)", 13), ("Sharpe", 10), ("Calmar", 10),
        ("同类排名百分位(%)", 16), ("股票占净值(%)", 13), ("债券占净值(%)", 13),
        ("机构持有(%)", 12), ("基金规模(亿)", 12), ("管理费率", 10),
        ("基金经理", 20), ("申购状态", 16), ("赎回状态", 16), ("流动性分组", 16),
        ("持仓日期", 12),
    ]
    hdr_fill = PatternFill("solid", fgColor="1F4E79")
    hdr_font = Font(color="FFFFFF", bold=True, size=10)
    thin = Side(style="thin", color="DDE5F0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for i, (c, w) in enumerate(cols, start=1):
        cell = ws.cell(row=1, column=i, value=c)
        cell.fill, cell.font = hdr_fill, hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "C2"

    r = 2
    for f in sorted(funds, key=lambda x: -x["mv"]):
        vals = [f["code"], f["name"], f["cls"], f["mk"], f["atype"], f["acct"],
                f["sharesRaw"], round(f["shares"], 6), round(f["cost"], 2), round(f["navcost"], 4),
                round(f["nav"], 4), f["navDate"], round(f["mv"], 2),
                round(f["pnl"], 4), round(f["pnl"], 4), round(f["realized"], 4),
                round(f["ret"], 4), round(f["weight"], 4),
                round(f["w1"], 4), round(f["m1"], 4), round(f["m3"], 4), round(f["m6"], 4),
                round(f["ytd"], 4), round(f["y1"], 4),
                round(f["mdd"], 4), round(f["vol"], 4), round(f["sharpe"], 4), round(f["calmar"], 4),
                round(f["rank"], 4), round(f["stock"], 4), round(f["bond"], 4),
                round(f["instHold"], 4), round(f["scaleNum"], 4), f["fee"],
                f["mgrPerson"], f["sgzt"], f["shzt"], f["liqShort"], meta["valDate"]]
        for i, v in enumerate(vals, start=1):
            c = ws.cell(row=r, column=i, value=v)
            c.border = border
            c.font = Font(size=10)
            c.alignment = Alignment(horizontal="left" if i <= 6 or i == 36 else "right", vertical="center")
        if f["pnl"] >= 0:
            ws.cell(row=r, column=14).font = Font(size=10, bold=True, color="C00000")
        else:
            ws.cell(row=r, column=14).font = Font(size=10, bold=True, color="1F7D51")
        ws.cell(row=r, column=18).number_format = "0.00"
        r += 1

    ws.cell(row=r, column=2, value="合计（%d 只）" % len(funds)).font = Font(bold=True, size=10)
    for col, val in ((8, round(sum(x["shares"] for x in funds), 6)), (9, round(totals["cost"], 2)),
                     (13, round(totals["mv"], 2)), (14, round(totals["pnl"], 4)),
                     (16, round(totals["realized"], 4)), (17, round(totals["ret"], 4))):
        c = ws.cell(row=r, column=col, value=val)
        c.font = Font(bold=True, size=10, color="C00000" if totals["pnl"] >= 0 else "1F7D51")
    ws.cell(row=r, column=18, value=100.0).font = Font(bold=True, size=10)

    # 第二张表：组合汇总
    ws2 = wb.create_sheet("组合汇总")
    summary = [
        ("估值日", meta["valDate"]), ("最新净值日", meta["navDate"]),
        ("生成时间", meta["genTime"]), ("持仓基金数", len(funds)),
        ("组合总市值(万元)", round(totals["mv"], 2)),
        ("持仓总成本(万元)", round(totals["cost"], 2)),
        ("累计浮动盈亏(万元)", round(totals["pnl"], 4)),
        ("账面收益率(%)", round(totals["ret"], 4)),
        ("已实现损益(万元)", round(totals["realized"], 4)),
        ("", ""),
        ("— 按产品分类 —", ""),
    ]
    agg = {}
    for f in funds:
        agg.setdefault(f["cls"], [0, 0])
        agg[f["cls"]][0] += f["mv"]
        agg[f["cls"]][1] += f["pnl"]
    for k, v in sorted(agg.items(), key=lambda x: -x[1][0]):
        summary.append((k, "市值 %.2f 万元 / 权重 %.2f%% / 盈亏 %+.4f 万元"
                        % (v[0], v[0] / totals["mv"] * 100, v[1])))
    summary.append(("", ""))
    summary.append(("— 按流动性 —", ""))
    agg2 = {}
    for f in funds:
        agg2.setdefault(f["liqText"], 0.0)
        agg2[f["liqText"]] += f["mv"]
    for k, v in sorted(agg2.items(), key=lambda x: -x[1]):
        summary.append((k, "市值 %.2f 万元 / 权重 %.2f%%" % (v, v / totals["mv"] * 100)))

    for i, (a, b) in enumerate(summary, start=1):
        ws2.cell(row=i, column=1, value=a).font = Font(bold=str(a).startswith("—"), size=10)
        ws2.cell(row=i, column=2, value=b).font = Font(size=10)
    ws2.column_dimensions["A"].width = 22
    ws2.column_dimensions["B"].width = 52

    wb.save(out_path)
    return out_path


# ================================================================== 7. 截图
def find_chrome():
    for p in CHROME_CANDIDATES:
        if os.path.exists(p):
            return p
    return None


def screenshot(html_path, out_prefix, scale=2, chunk_px=3000, verbose=True):
    """整页高清截图 + 分片。返回 (full_png, [chunk_png...])"""
    chrome = find_chrome()
    if not chrome:
        log("! 未找到 Chrome/Edge，跳过截图")
        return None, []
    if Image is None:
        log("! 缺少 Pillow，跳过分片")
    full = os.path.abspath(out_prefix + "_全图.png")
    url = "file:///" + os.path.abspath(html_path).replace("\\", "/").replace("#", "%23")
    cmd = [chrome, "--headless=new", "--disable-gpu", "--no-sandbox", "--hide-scrollbars",
           "--force-device-scale-factor=%d" % scale, "--virtual-time-budget=15000",
           "--window-size=1400,9000", "--screenshot=" + full, url]
    try:
        subprocess.run(cmd, capture_output=True, timeout=180)
    except Exception as e:                  # noqa: BLE001
        log("! 截图失败：%s" % e)
        return None, []
    if not os.path.exists(full):
        log("! 截图未生成")
        return None, []

    chunks = []
    if Image is not None:
        im = Image.open(full).convert("RGB")
        w, h = im.size

        # 自动裁掉底部空白（页面背景色 #eef1f6），避免长图尾部大面积留白
        bg = im.getpixel((5, h - 5))
        x_probe = [int(w * r) for r in (0.12, 0.35, 0.5, 0.65, 0.88)]
        last = h
        y = h - 1
        while y > 0:
            hit = False
            for x in x_probe:
                px = im.getpixel((x, y))
                if abs(px[0] - bg[0]) + abs(px[1] - bg[1]) + abs(px[2] - bg[2]) > 12:
                    hit = True
                    break
            if hit:
                last = min(h, y + 90)
                break
            y -= 4
        if last < h:
            im = im.crop((0, 0, w, last))
        w, h = im.size

        n = max(1, math.ceil(h / chunk_px))
        step = math.ceil(h / n)
        for i in range(n):
            top = i * step
            bot = min(h, top + step)
            if bot - top < 60:
                break
            p = "%s_%02d.png" % (out_prefix, i + 1)
            im.crop((0, top, w, bot)).save(p, optimize=True)
            chunks.append(p)
        im.save(full, optimize=True)
        im.close()
    if verbose:
        log("截图完成：%s（%d 张分片）" % (os.path.basename(full), len(chunks)))
    return full, chunks


# ================================================================== 8. main
def main():
    ap = argparse.ArgumentParser(description="FOF 组合投后管理周更工具")
    ap.add_argument("--holdings", default=DEFAULT_HOLDINGS, help="持仓 Excel 路径")
    ap.add_argument("--out", default=DEFAULT_OUTDIR, help="输出目录")
    ap.add_argument("--date", default=None, help="估值日标签 YYYY-MM-DD，默认取最新净值日")
    ap.add_argument("--nav-json", default=None, help="外部净值覆盖 JSON：{code: {nav:..., ...}}")
    ap.add_argument("--no-shot", action="store_true", help="不生成截图")
    ap.add_argument("--no-excel", action="store_true", help="不生成 Excel")
    ap.add_argument("--check", action="store_true", help="只体检，不写文件")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args()

    v = not args.quiet
    global LOG_FILE
    LOG_FILE = os.path.join(args.out, HIST_SUBDIR, "更新日志.txt")
    ov = {}
    if args.nav_json and os.path.exists(args.nav_json):
        with open(args.nav_json, encoding="utf-8") as fh:
            ov = json.load(fh)
        log("使用外部净值覆盖：%d 只" % len(ov))

    header, recs, funds, totals, meta = build(ov, args.holdings, verbose=v)
    if args.date:
        meta["valDate"] = args.date
    if not funds:
        log("没有有效持仓（份额均 <= 0），退出")
        return 1

    d0 = datetime.strptime(meta["valDate"], "%Y-%m-%d")
    prev = [s for s in load_snapshots(args.out) if s.get("date") != meta["valDate"]]
    changes = diff_snapshots(prev[-1], funds, meta["closed"], totals) if prev else None

    log("=" * 62)
    log("估值日 %s ｜ 最新净值日 %s ｜ 持仓 %d 只" % (meta["valDate"], meta["navDate"], len(funds)))
    for f in sorted(funds, key=lambda x: -x["mv"]):
        log("  %-10s %-28s 市值 %12s 万  盈亏 %10s 万  %8s  权重 %5.2f%%"
            % (f["code"], f["name"][:26], fmt(f["mv"]), fmt(f["pnl"], 4, True),
               pct(f["ret"], 4), f["weight"]))
    log("  合计        市值 %s 万  成本 %s 万  盈亏 %s 万  %s"
        % (fmt(totals["mv"]), fmt(totals["cost"]), fmt(totals["pnl"], 4, True), pct(totals["ret"], 4)))
    if changes and changes["changed"]:
        log("  本期变动：" + "、".join("%s %s" % (x["short"], x["type"]) for x in changes["changed"]))
    log("=" * 62)

    if args.check:
        log("--check 模式，未写入任何文件")
        return 0

    os.makedirs(args.out, exist_ok=True)
    tag = meta["valDate"].replace("-", "")

    excel_path = None
    if not args.no_excel:
        excel_path = os.path.join(args.out, "基金_投后更新_%s.xlsx" % tag)
        write_excel(funds, totals, meta, excel_path)
        log("Excel 已更新：%s" % excel_path)

    html_path = os.path.join(args.out, "FOF组合投后管理分析报告_%s.html" % tag)
    render_html(funds, totals, meta, changes, html_path)
    log("报告已生成：%s" % html_path)

    shot = None
    chunks = []
    if not args.no_shot:
        shot, chunks = screenshot(html_path, os.path.join(args.out, "FOF组合投后管理分析报告_%s" % tag))

    snap_p, _ = save_snapshot(args.out, funds, totals, meta)
    log("历史快照：%s" % snap_p)

    log("")
    log("===== 完成 =====")
    log("Excel     : %s" % (excel_path or "（已跳过）"))
    log("报告 HTML : %s" % html_path)
    if shot:
        log("高清全图  : %s" % shot)
        for c in chunks:
            log("分片      : %s" % c)
    return 0


if __name__ == "__main__":
    sys.exit(main())
