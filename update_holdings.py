# -*- coding: utf-8 -*-
"""
持仓更新一键脚本：用新的 基金.xls 覆盖云端加密持仓并推送。

用法：
  python update_holdings.py --xls "D:\\...\\基金.xls" [--push]

- 默认读取同目录 .holdings_key 作为加密密钥（与仓库 secret HOLDINGS_KEY 一致）
- --push 时会自动 commit + push（需要 GITHUB_TOKEN 环境变量，或仓库 remote 已带 token）
- 不传 --push 时只在本地生成 data/holdings.bin，便于人工核对后再推
"""
import argparse
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
KEY_FILE = os.path.join(HERE, ".holdings_key")
ENC = os.path.join(HERE, "data", "holdings.bin")
CRYPT = os.path.join(HERE, "crypt_holdings.py")


def run(cmd, **kw):
    print("$", " ".join(cmd))
    return subprocess.run(cmd, cwd=HERE, check=True, **kw)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xls", required=True, help="新的持仓 Excel 路径")
    ap.add_argument("--push", action="store_true", help="提交并推送到 GitHub")
    a = ap.parse_args()

    if not os.path.exists(a.xls):
        sys.exit("找不到文件：%s" % a.xls)
    if not os.path.exists(KEY_FILE):
        sys.exit("缺少密钥文件 %s" % KEY_FILE)
    key = open(KEY_FILE).read().strip()

    py = sys.executable
    run([py, CRYPT, "encrypt", a.xls, ENC, key])
    print("已生成加密持仓：", ENC)

    # 同步维护分红起点登记表：新基金登记本次「持仓日期」，已有基金保留原值
    try:
        import json as _json
        import re as _re
        from datetime import datetime, timedelta
        import xlrd
        wb = xlrd.open_workbook(a.xls)
        sh = wb.sheet_by_index(0)
        hd = [str(sh.cell_value(1, c)).strip() for c in range(sh.ncols)]
        try:
            ic, idt = hd.index("基金代码"), hd.index("持仓日期")
        except ValueError:
            ic = idt = None
        fd_path = os.path.join(HERE, "first_dates.json")
        fd = {}
        if ic is not None and os.path.exists(fd_path):
            fd = _json.load(open(fd_path, encoding="utf-8"))
        for r in range(2, sh.nrows):
            code = str(sh.cell_value(r, ic)).strip()
            if not code:
                continue
            if code in fd:
                continue                     # 已登记，保留最早日期
            raw = sh.cell_value(r, idt)
            if isinstance(raw, (int, float)) and raw:
                raw = (datetime(1899, 12, 30) + timedelta(days=float(raw))).strftime("%Y-%m-%d")
            m = _re.match(r"(\d{4}-\d{2}-\d{2})", str(raw))
            if m:
                fd[code] = m.group(1)
                print("登记分红起点：%s %s" % (code, fd[code]))
        if fd:
            _json.dump(fd, open(fd_path, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    except Exception as e:                   # noqa: BLE001
        print("分红起点登记失败（不影响持仓加密）：%s" % e)

    if not a.push:
        print("（未推送，确认无误后加 --push 重新运行）")
        return

    run(["git", "add", "data/holdings.bin"])
    try:
        run(["git", "-c", "user.name=fof-bot", "-c", "user.email=fof-bot@local",
             "commit", "-m", "update holdings"])
    except subprocess.CalledProcessError:
        print("没有变化，无需提交")
        return
    run(["git", "push", "origin", "main"])
    print("已推送。云端下次定时运行（或手动触发 Actions）即使用新持仓。")


if __name__ == "__main__":
    main()
