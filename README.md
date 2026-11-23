# 双端商店评分看板

Frontier City 与 Vice Throne 的逐国家/地区 Google Play、App Store 评分看板。

- 公开网址由 GitHub Pages 提供。
- GitHub Actions 每 4 小时自动抓取一次，并刷新页面与 Excel。
- 无公开地区评分时显示“暂无评分”，不使用全球评分或默认评分回退。
- 数据仅来自公开商店页面，不需要登录，也不会收集访问者信息。

## 手动更新

进入仓库的 **Actions** 页面，选择 **Update store ratings**，点击 **Run workflow**。

## 本地运行

```bash
pip install -r requirements.txt
python update_ratings.py
```

生成的 `index.html` 为看板，`商店评分汇总_latest.xlsx` 为最新 Excel。