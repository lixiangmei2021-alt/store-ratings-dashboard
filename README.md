# 双端商店评分看板

公开网址：https://lixiangmei2021-alt.github.io/store-ratings-dashboard/


Frontier City 与 Vice Throne 的逐国家/地区 Google Play、App Store 评分看板。

- 公开网址由 GitHub Pages 提供。
- GitHub Actions 每 4 小时自动抓取一次，并刷新页面与 Excel。
- 每次刷新都会保存逐国评分快照，可查看趋势图、评分涨跌和完整刷新时间线。
- 无公开地区评分时显示“暂无评分”，不使用全球评分或默认评分回退。
- 数据仅来自公开商店页面，不需要登录，也不会收集访问者信息。

## 手动更新

点看板上的 **手动更新**，首次使用按弹窗里的步骤配置一次令牌：

1. 点 **打开 GitHub 创建令牌**，GitHub 会打开经典令牌页面并自动勾选 `repo` 与 `workflow`（建议同时设置较短有效期）。
2. 点页面底部的 **Generate token**，复制以 `ghp_` 开头的令牌。
3. 回到看板粘贴，点 **保存并启用**。

之后每次点 **立即更新**，都会直接在本页启动 GitHub Actions，等待新数据后自动刷新，不需要进仓库点 Run workflow。

权限更小的做法：改用 Fine-grained token，仓库范围只选 `store-ratings-dashboard`，仓库权限设为 `Actions: Read and write`。

令牌只保存在当前浏览器的 localStorage，不写入公开网页或仓库，可随时在看板中清除或在 GitHub 撤销。其他访客仍可查看看板，但没有令牌不能启动更新。

### 页面看起来是旧版

如果点 **手动更新** 后直接跳到 GitHub 仓库、看不到弹窗，说明浏览器用了缓存页面，按 `Ctrl+F5`（Mac：`⌘+Shift+R`）强制刷新。工具栏会显示当前页面版本号，可与 `version.json` 对照。

## 本地运行

```bash
pip install -r requirements.txt
python update_ratings.py
```

生成的 `index.html` 为看板，`商店评分汇总_latest.xlsx` 为最新 Excel。
