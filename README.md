# 双端商店评分看板

公开网址：https://lixiangmei2021-alt.github.io/store-ratings-dashboard/


Frontier City 与 Vice Throne 的逐国家/地区 Google Play、App Store 评分看板。

- 公开网址由 GitHub Pages 提供。
- GitHub Actions 每 4 小时自动抓取一次，并刷新页面与 Excel。
- 每次刷新都会保存逐国评分快照，可查看趋势图、评分涨跌和完整刷新时间线。
- 无公开地区评分时显示“暂无评分”，不使用全球评分或默认评分回退。
- 数据仅来自公开商店页面，不需要登录，也不会收集访问者信息。

## 手动更新

点击看板上的 **手动更新**，首次使用时保存一个仅限本仓库的 GitHub Fine-grained token：

1. 在弹窗中点击 **创建更新令牌**。
2. 将仓库访问范围设为仅 `store-ratings-dashboard`。
3. 在仓库权限中授予 **Actions: Read and write**。
4. 复制令牌并粘贴回看板，点击 **保存并启用**。

之后点击 **立即更新** 即可从页面直接启动 GitHub Actions；页面会等待新数据并自动刷新。令牌只保存在当前浏览器，不写入公开网页，可随时在页面中清除或在 GitHub 撤销。其他访客仍可查看看板，但没有令牌时不能启动更新。

## 本地运行

```bash
pip install -r requirements.txt
python update_ratings.py
```

生成的 `index.html` 为看板，`商店评分汇总_latest.xlsx` 为最新 Excel。
