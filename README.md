# FOF 投后周报 · 云端自动版

每周六早上 09:30（北京时间）由 GitHub Actions 云端自动运行，**不依赖你的电脑开机**，
报告发布到 GitHub Pages，并推送摘要+图片链接到你的微信（Server酱）。

## 每周流程

```
GitHub Actions 定时触发（周六 09:30 北京时间）
  → 拉取最新净值（东方财富公开接口）
  → 读取 data/基金.xls 重算市值/盈亏/变动
  → 生成报告 HTML + 高清长图 + 分片图 + 更新版 Excel
  → 发布到 GitHub Pages（可在线打开）
  → Server酱 推送摘要到你的微信
```

## 首次部署（网页 3 步 + 自动运行）

**步骤 1 · 建仓库**：登录 github.com → 右上角 + → New repository →
名字填 `fof-report`，选 **Public** → Create。

**步骤 2 · 配置微信推送密钥**：仓库页 → Settings → Secrets and variables →
Actions → New repository secret：
- Name: `SERVERCHAN_SENDKEY`
- Secret: 粘贴你的 Server酱 SendKey

**步骤 3 · 通知助手推送代码**（git 走 github.com 主站，可直连），
推送完成后 Actions 会自动运行一次完整流程（生成报告 → 发布 gh-pages 分支 → 推微信）。

**步骤 4 · 开启报告网页（仅首次）**：首次运行成功后，仓库 Settings → Pages →
Build and deployment → Source 选 `Deploy from a branch` → Branch 选 `gh-pages` /(root) → Save。
以后每次运行自动更新，无需再动。

## 持仓变了怎么办

仓库中的持仓文件是**加密的**（`data/holdings.bin`，密钥在 Actions secret `HOLDINGS_KEY`，明文 `基金.xls` 不入库）。

持仓变更两种方式：
- **推荐**：直接告诉我（WorkBuddy），我帮你改份额 → 重新加密 → 推送；
- **自助**：本地改 `基金.xls` 后执行：
  ```bash
  python crypt_holdings.py encrypt data/基金.xls data/holdings.bin <HOLDINGS_KEY>
  git add data/holdings.bin && git commit -m "update holdings" && git push
  ```

下次定时运行自动使用新持仓；若当周就想立刻更新，Actions 页手动 Run workflow。

## 隐私说明

- 持仓数据在仓库中加密存储，密钥仅存在于 GitHub Secrets；
- Actions 日志不输出任何持仓/收益数据（调试需手动设 DEBUG_PUSH=1）；
- 仍公开的部分：gh-pages 上的报告网页与图片（含持仓明细）。这是"维持公开网页"方案的固有暴露面，链接不易被猜测；如需彻底私密，转私有仓库并关闭 Pages 即可。

## 本地推送（首次配置仓库用）

```bash
cd fof_cloud
git init -b main
git add .
git commit -m "FOF weekly report pipeline"
git remote add origin https://github.com/<你的用户名>/fof-report.git
git push -u origin main
```

## 说明

- 免费额度：公开仓库 Actions 每月 2000 分钟（本任务每次约 2-3 分钟，绰绰有余）。
- 隐私：Private 仓库也能用 Pages（GitHub 免费 Private 仓库支持 Pages，但 Pages 站点对拥有链接的人可见）。报告链接含随机性不高，介意可在 Settings → Pages 随时关闭站点，只保留微信推送的摘要文字。
- Server酱免费版每天 5 条，每周 1 条完全够用；若改用 PushPlus，把 secret 名换成 `PUSHPLUS_TOKEN` 即可，脚本自动识别。
