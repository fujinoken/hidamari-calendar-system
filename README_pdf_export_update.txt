ひだまり帳 Ver1.3.1 PDFカレンダー出力版

変更点:
- Excel出力メニューを「Excel・PDF出力」に変更
- A4横の月間カレンダーPDF出力を追加
- PDF 1ページ目: 月間カレンダー
- PDF 2ページ目以降: 予定詳細一覧
- 日本語表示用にReportLabのCIDフォントを使用
- 重要予定は薄赤背景で表示

配置方法:
1. app.py として配置してください。
2. requirements.txt に reportlab を追加してください。
3. Streamlit Cloudで再デプロイまたはRebootしてください。

requirements.txt:
streamlit
pandas
openpyxl
psycopg2-binary
reportlab

注意:
- PDF枠内は要約表示です。
- 詳細は2ページ目以降の予定詳細一覧で確認できます。
- 絵文字はPDFで文字化けしやすいため、PDF内では【通院】などカテゴリ名で表示します。
