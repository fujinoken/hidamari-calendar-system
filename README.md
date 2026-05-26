# ひだまり帳 Ver1.1

## 追加内容

- 「今日は何ある」メニューを追加
- 今日の予定だけを一覧表示
- カテゴリ別件数サマリー
- 重要マーク件数表示
- 予定ボタンを押すと同じページ内に詳細表示
- 写真メモ・Excel添付も同じ画面で確認可能

## 起動方法

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 保存方式

- 本体：SQLite（hidamari_calendar.db）
- 写真：uploads/
- Excel添付：attached_files/
