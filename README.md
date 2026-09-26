# 六分身兼職備料

瑪奇 Mobile 每日兼職備料工具：6 隻角色分開勾選兼職，自動算出全部原料總數，並列出每隻角色的交付路線。

網站：https://hsingshiue.github.io/mabi-jobs/

## 更新資料

```
pip install openpyxl
python sync.py
```

會重新抓原網站與試算表，產生 `jobs-data.js`，之後 commit 並 push 即可。

## 資料來源

- [瑪奇M 兼職列表](https://docs.google.com/spreadsheets/d/1dY03kiMC4x3jBB4gA71jdo2YqsLAgHC7kp-WmbsNSnU/)（墨音 整理）
- [瑪奇M 每日兼職備料器](https://mabinogi-mobile-jobs.vtuberparrot2021.chatgpt.site/)

感謝原作者整理。
