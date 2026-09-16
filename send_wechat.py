# -*- coding: utf-8 -*-
"""
推送 FOF 周报摘要到微信（Server酱 / PushPlus 二选一，自动识别）。

环境变量：
  SERVERCHAN_SENDKEY  Server酱 Turbo 的 SendKey（sct.ftqq.com 获取）
  PUSHPLUS_TOKEN      PushPlus 的 token（pushplus.plus 获取）
  REPORT_DIR          报告输出目录（site/）
  PAGE_BASE           GitHub Pages 基础 URL，如 https://xxx.github.io/fof-report
"""
import io
import os
import re
import json
import glob
import sys

try:
    import urllib.request
    import urllib.parse
except ImportError:
    sys.exit("need python3")

RED = "🔴"
GRN = "🟢"


def sgn(v, nd=2):
    return ("%+." + str(nd) + "f%%") % v if v is not None else "—"


def find_report_dir():
    d = os.environ.get("REPORT_DIR") or "site"
    return d


def pick(pattern):
    fs = sorted(glob.glob(os.path.join(find_report_dir(), pattern)))
    return os.path.basename(fs[-1]) if fs else None


def build_desp():
    d = find_report_dir()
    # 从生成的 HTML 中抽 REPORT 数据，避免重复计算
    htmls = sorted(glob.glob(os.path.join(d, "*分析报告*.html")))
    if not htmls:
        return "⚠️ 未找到生成的报告 HTML，请检查 Actions 运行日志。"
    s = io.open(htmls[-1], encoding="utf-8").read()
    m = re.search(r"var REPORT\s*=\s*(\{.*?\});\s*\n", s, re.S)
    if not m:
        return "⚠️ 报告数据解析失败。"
    rep = json.loads(m.group(1))
    tot = rep["total"]
    funds = rep["funds"]
    chg = rep.get("changes") or {}

    lines = []
    lines.append("**FOF组合投后周报**（估值日 %s）" % rep.get("valDate", ""))
    lines.append("")
    lines.append("- 组合市值：**%s 万元**" % format(tot["mv"], ",.2f"))
    pl = tot["pnl"]
    lines.append("- 浮动盈亏：%s **%s 万元**（%s）" % (
        RED if pl >= 0 else GRN, format(pl, ",.2f"), sgn(tot.get("ret"))))
    lines.append("- 持仓基金：**%d 只**" % len(funds))
    lines.append("")
    lines.append("| 基金 | 权重 | 盈亏(万) | 近1年 |")
    lines.append("|---|---|---|---|")
    for f in sorted(funds, key=lambda x: -x["mv"]):
        lines.append("| %s | %.1f%% | %s%s | %s |" % (
            f["short"], f["mv"] / tot["mv"] * 100.0,
            RED if f["pnl"] >= 0 else GRN, format(f["pnl"], "+,.2f"),
            sgn(f.get("y1"))))
    lines.append("")

    items = (chg.get("items") if isinstance(chg, dict) else None) or []
    if items:
        lines.append("**⚠️ 本期持仓变动**")
        for it in items:
            lines.append("- %s：%s" % (it.get("code", ""), it.get("text", "")))
        lines.append("")

    base = (os.environ.get("PAGE_BASE") or "").rstrip("/")
    if base:
        lines.append("---")
        full = pick("*全图.png")
        html_name = os.path.basename(htmls[-1])
        quoted = urllib.parse.quote(html_name)
        lines.append("[📊 打开完整交互报告](%s/%s)" % (base, quoted))
        for i in range(1, 9):
            p = pick("*_%02d.png" % i)
            if not p:
                break
            lines.append("![周报第%d页](%s/%s)" % (i, base, urllib.parse.quote(p)))
        if full:
            lines.append("[🖼 高清长图](%s/%s)" % (base, urllib.parse.quote(full)))
    return "\n".join(lines)


def push_serverchan(key, title, desp):
    url = "https://sctapi.ftqq.com/%s.send" % key
    data = urllib.parse.urlencode({"title": title, "desp": desp}).encode()
    req = urllib.request.Request(url, data=data)
    with urllib.request.urlopen(req, timeout=30) as r:
        body = json.loads(r.read().decode())
    return body.get("code") == 0, str(body)


def push_pushplus(token, title, desp):
    url = "https://www.pushplus.plus/send"
    payload = {"token": token, "title": title, "content": desp, "template": "markdown"}
    data = json.dumps(payload).encode()
    req = urllib.request.Request(url, data=data,
                                 headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=30) as r:
        body = json.loads(r.read().decode())
    return body.get("code") == 200, str(body)


def main():
    title = "FOF投后周报"
    desp = build_desp()
    # 内容默认不落 Actions 日志（公开仓库日志人人可见），调试时设 DEBUG_PUSH=1
    if os.environ.get("DEBUG_PUSH"):
        print(desp)
    key = os.environ.get("SERVERCHAN_SENDKEY", "").strip()
    tok = os.environ.get("PUSHPLUS_TOKEN", "").strip()
    if key:
        ok, info = push_serverchan(key, title, desp)
    elif tok:
        ok, info = push_pushplus(tok, title, desp)
    else:
        print("!! 未配置 SERVERCHAN_SENDKEY / PUSHPLUS_TOKEN，跳过推送")
        return 0
    print("push:", "OK" if ok else info)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
