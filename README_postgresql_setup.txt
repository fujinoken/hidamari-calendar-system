ひだまり帳 PostgreSQL版 セットアップメモ

1. app.py として配置
   hidamari_calendar_postgresql_app.py を app.py にリネームして使ってください。

2. requirements.txt
   requirements_postgresql.txt の内容を requirements.txt に入れてください。
   必須: psycopg2-binary

3. Streamlit secrets の設定例
   DATABASE_URL = "postgresql://ユーザー名:パスワード@ホスト名:ポート/DB名"

   Supabaseの場合は、Project Settings → Database → Connection string の
   Transaction pooler または Session pooler のURLを使います。
   パスワード部分は自分のDBパスワードに置き換えてください。

4. 注意
   予定・利用者・職員・カテゴリ・写真/ファイルの紐づけ情報はPostgreSQLに保存されます。
   ただし、写真やExcelそのものは uploads / attached_files フォルダ保存のままです。
   Streamlit Cloud等ではファイル本体も永続化したい場合、次の段階でSupabase Storage化してください。
