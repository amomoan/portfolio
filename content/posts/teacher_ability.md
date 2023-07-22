+++
title = "学力に関する証明書作成アプリ"
date = "2023-07-19"
author = ""
authorTwitter = "" #do not include @
cover = "/portfolio/img/gakuryoku.png"
tags = ["Python", "pandas","Flask","Excel","SQL","ポートフォリオ","業務改善"]
keywords = ["", ""]
description = "「学力に関する証明書」を発行するためにウェブアプリを作成。"
showFullContent = false
readingTime = false
hideComments = false
color = "" #color from the theme settings
+++
## 概要

大学には教職課程に関して「学力に関する証明書」という証明書がある。これは文部科学省が様式を公開しており、決まったルールどおりに入力する必要がある。
しかし、入学した年度や各種条件により現行法の科目に読み替える必要がある。

この操作は大変複雑で、かつ、法改正が高頻度であるため、本証明書もそれに対応するため様式を変更する必要がる。とても流動的な証明書である。
多くの大学では、この証明書の発行に多大な労力を払っている。

この証明書を発行するために今までは以下の手順を踏んでいた。

1. 成績証明書を出力
2. 当時の講義要項を印刷
3. 読替表を印刷
4. 読替のチェック（蛍光ペンで紙にマーク）
5. その結果を文部科学省の様式に転記
6. 備考欄に特記事項を記載

以上5までのステップはロジカルなものであるため、機械化できると判断した。

当初Excelで完結するつもりだったが、これまでの成績データが 110万件を超えたため断念。
職場の制約を満たすため、成績データや個人情報などはDB化しNASに配置し、ローカルホストでサーバーを立ち上げウェブアプリからDBに接続する方法を採用した。

機密情報があるためすべては公開できないが、上記の5までのステップは1分かからず実行可能。
また、確認資料とできるように生データに近い形をExcelのシートに書き出すように工夫した。
ただし、成績情報が電子データとして保存されている場合に限定される。

## 効果

現時点ではテンプレートや読替表を整備中のため稼働していないので予測値。

3時間 → 3分程度

## ライブラリ

- `openpyxl` は、Excel 2010 xlsx/xlsm/xltx/xltmファイルを読み書きするためのPythonライブラリ。
- `pandas` は、データ操作と分析のためのPythonライブラリで、データフレームというデータ構造を提供している。
- `flask` は、軽量なWSGIウェブアプリケーションフレームワークで、簡単なアプリケーションから複雑なアプリケーションまでスケールアップすることができる。
- `flask_sqlalchemy` は、FlaskアプリケーションにSQLAlchemyのサポートを追加する拡張機能。FlaskとSQLAlchemyを使用することを簡単にするために、便利なデフォルトと追加のヘルパーが提供されている。
- `sqlalchemy` は、Python SQLツールキットおよびオブジェクト関係マッパーであり、データベースとの対話を容易にする。
- `dotenv` は、`.env` ファイルから環境変数を読み込み、`process.env` にロードするためのゼロ依存モジュール。コードから環境設定を分離して保存することは、The Twelve-Factor App方法論に基づいている。

## flask

```python
# app.py
import io
import os
import re
import sqlite3
from urllib.parse import quote

import openpyxl
import pandas as pd
from flask import (Flask, Response, current_app, flash, g, redirect,
                   render_template, request, session)
from flask_sqlalchemy import SQLAlchemy
from openpyxl.utils.dataframe import dataframe_to_rows
from sqlalchemy import text
from sqlalchemy.exc import SQLAlchemyError
from dotenv import load_dotenv

# Flaskアプリケーションを作成
app = Flask(__name__)
load_dotenv()
app.secret_key = os.getenv('SECRET_KEY', 'for dev')

import os
from dotenv import load_dotenv

load_dotenv()

app.secret_key = os.getenv('SECRET_KEY', 'for dev')

# app.secret_key = "fdasf123124fasdf"
# データベースファイルのパスを設定
app.config["DATABASE"] = r"path/to/sqlite.db"

db_path = os.path.abspath("instance/df_db.sqlite")
# データベースURIを設定
app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{db_path}"
df_db = SQLAlchemy(app)


class Score(df_db.Model):
    id = df_db.Column(df_db.Integer, primary_key=True)
    student_id = df_db.Column(df_db.String(80))
    student_name = df_db.Column(df_db.String(80))
    birthdate = df_db.Column(df_db.String(80))
    student_affiliation = df_db.Column(df_db.String(80))
    subject_code = df_db.Column(df_db.String(80))
    subject_name = df_db.Column(df_db.String(80))
    evaluation = df_db.Column(df_db.String(80))
    certification_year = df_db.Column(df_db.String(80))
    yomikae_to = df_db.Column(df_db.String(80))
    license_subject = df_db.Column(df_db.String(80))
    faculty_department = df_db.Column(df_db.String(80))

    def __init__(
        self,
        student_id,
        student_name,
        birthdate,
        student_affiliation,
        subject_code,
        subject_name,
        evaluation,
        certification_year,
        yomikae_to,
        license_subject,
        faculty_department,
    ):
        self.student_id = student_id
        self.student_name = student_name
        self.birthdate = birthdate
        self.student_affiliation = student_affiliation
        self.subject_code = subject_code
        self.subject_name = subject_name
        self.evaluation = evaluation
        self.certification_year = certification_year
        self.yomikae_to = yomikae_to
        self.license_subject = license_subject
        self.faculty_department = faculty_department


# データベース接続を取得する関数
def get_db():
    # gオブジェクトにsqlite_db属性がない場合
    if not hasattr(g, "sqlite_db"):
        # データベース接続を作成してgオブジェクトに保存
        g.sqlite_db = sqlite3.connect(current_app.config["DATABASE"])
    # データベース接続を返す
    return g.sqlite_db


# Flaskアプリケーションのコンテキスト内でデータベース操作を行う
with app.app_context():
    # データベースが存在しない場合、データベースとテーブルを作成する
    if not os.path.exists(db_path):
        os.makedirs("instance", exist_ok=True)
    try:
        with app.app_context():
            with df_db.engine.connect() as conn:
                trans = conn.begin()
                try:
                    result = conn.execute(text("DROP TABLE score;"))
                    trans.commit()
                except:
                    trans.rollback()
                    raise
    except SQLAlchemyError as e:
        print(f"An error occurred while executing the SQL statement: {e}")

    # データベースとテーブルを作成する。もしなくてもディレクトリごと作成される。
    df_db.create_all()


# アプリケーションコンテキストが終了するときに呼び出される関数
@app.teardown_appcontext
def close_db(exception):
    # gオブジェクトにsqlite_db属性がある場合
    if hasattr(g, "sqlite_db"):
        # データベース接続を閉じる
        g.sqlite_db.close()


# ルートURLにアクセスしたときに呼び出される関数
@app.route("/", methods=["GET", "POST"])
def index():
    # print("start!")
    # print(df_db.engine)
    # for table in df_db.metadata.sorted_tables:
    # print(table.name)

    try:
        with app.app_context():
            with df_db.engine.connect() as conn:
                trans = conn.begin()
                try:
                    result = conn.execute(text("DELETE FROM score;"))
                    trans.commit()
                except:
                    trans.rollback()
                    raise
    except SQLAlchemyError as e:
        print(f"An error occurred while executing the SQL statement: {e}")

    # 変数を初期化
    student_id = ""
    yomikae_from = ""
    yomikae_to = ""
    ability = ""
    name = ""
    belonging = ""
    yomikae_belonging = ""
    scores = []

    select_columns = []
    ability_columns = []
    yomikae_belonging_columns = []
    error = ""

    db = get_db()
    df = pd.read_sql_query("SELECT * FROM yomikae LIMIT 1", db)
    select_columns = [col for col in df.columns if re.match(r"\d{4}科目名", col)]
    select_columns.sort(reverse=True)
    df = pd.read_sql_query(
        'SELECT "免許教科" FROM yomikae GROUP BY "免許教科" HAVING "免許教科"!="教職"', db
    )
    ability_columns = df["免許教科"]
    df = pd.read_sql_query(
        'SELECT "学部学科" FROM yomikae GROUP BY "学部学科" HAVING 学部学科 != "教職教職" ORDER BY "学部学科"',
        db,
    )
    yomikae_belonging_columns = df["学部学科"]
    session["yomikae_to"] = yomikae_to
    session["yomikae_from"] = yomikae_from
    session["ability"] = ability
    session["yomikae_belonging"] = yomikae_belonging

    # POSTリクエストの場合
    if request.method == "POST":
        # フォームデータを取得
        student_id = request.form.get("student_id", "")
        yomikae_from = request.form.get("yomikae_from", "")
        yomikae_to = request.form.get("yomikae_to", "")
        ability = request.form.get("ability", "")
        yomikae_belonging = request.form.get("yomikae_belonging", "")
        session["yomikae_to"] = yomikae_to
        session["yomikae_from"] = yomikae_from
        session["ability"] = ability
        session["yomikae_belonging"] = yomikae_belonging

        sSQL = f"""
            SELECT
                A."学籍番号",
                A."学生氏名",
                A."生年月日",
                A."学生所属",
                A."科目コード",
                A."科目名",
                A."評価",
                A."認定年度",
                Y."{yomikae_to}",
                Y."免許教科",
                Y."学部学科"
            FROM
                (
                    all_score
                    LEFT JOIN student ON all_score."学籍番号" = student."学籍番号"
                ) AS A
                LEFT JOIN (
                    SELECT
                        "{yomikae_from}",
                        "{yomikae_to}",
                        "免許教科",
                        "学部学科"
                    FROM
                        yomikae
                    WHERE
                        (
                            "学部学科" = "{yomikae_belonging}"
                            OR "学部学科" = "教職教職"
                        )
                        AND
                        (
                            "免許教科" = "{ability}"
                            OR "免許教科" = "教職"
                            )
                        AND "{yomikae_from}" IS NOT NULL
                ) AS Y ON A."科目名" = REPLACE (Y."{yomikae_from}", '●', '')
            WHERE
                A."学籍番号" = ?
                """
        # student_idが入力されていない場合
        if not student_id:
            error = "学籍番号を入力してください。"
        else:
            # データベース接続を取得
            db = get_db()
        try:
            # SQLクエリを実行してデータフレームを取得
            df = pd.read_sql_query(sSQL, db, params=(student_id,))

        except pd.errors.DatabaseError as e:
            error = f"データベースエラー: {e}"
        else:
            df = df.sort_values([yomikae_to, "科目コード"]).copy()
            # データフレームが空でない場合
            if not df.empty:
                # 氏名と成績データを取得
                name = df.iloc[0]["学生氏名"]
                belonging = df.iloc[0]["学生所属"]
                columns = [
                    "学籍番号",
                    "学生氏名",
                    "生年月日",
                    "学生所属",
                    "科目コード",
                    "科目名",
                    "評価",
                    "認定年度",
                    yomikae_to,
                    "免許教科",
                    "学部学科",
                ]
                scores = df[columns].to_dict("records")

                for score in scores:
                    s = Score(
                        student_id=score["学籍番号"],
                        student_name=score["学生氏名"],
                        birthdate=score["生年月日"],
                        student_affiliation=score["学生所属"],
                        subject_code=score["科目コード"],
                        subject_name=score["科目名"],
                        evaluation=score["評価"],
                        certification_year=score["認定年度"],
                        yomikae_to=score[yomikae_to],
                        license_subject=score["免許教科"],
                        faculty_department=score["学部学科"],
                    )
                    df_db.session.add(s)
                df_db.session.commit()

        # テンプレートをレンダリングしてレスポンスを返す
    return render_template(
        "index.html",
        name=name,
        student_id=student_id,
        belonging=belonging,
        yomikae_from=yomikae_from,
        yomikae_to=yomikae_to,
        ability=ability,
        yomikae_belonging=yomikae_belonging,
        scores=scores,
        select_columns=select_columns,
        ability_columns=ability_columns,
        yomikae_belonging_columns=yomikae_belonging_columns,
        error=error,
    )


@app.route("/download")
def download():
    scores = Score.query.all()
    data = []
    for score in scores:
        data.append(
            {
                "学籍番号": score.student_id,
                "学生氏名": score.student_name,
                "生年月日": score.birthdate,
                "学生所属": score.student_affiliation,
                "科目コード": score.subject_code,
                "科目名": score.subject_name,
                "評価": score.evaluation,
                "認定年度": score.certification_year,
                session.get("yomikae_to"): score.yomikae_to,
                "免許教科": score.license_subject,
                "学部学科": score.faculty_department,
            }
        )
    df_db = pd.DataFrame(data)
    ## セッションからデータフレームを取得
    # print("download:")
    # print(df_db.head())
    template_path = r"xltemplate\tmp.xlsx"
    output_sheet_name = "output"
    # ワークブックを読み込む
    wb = openpyxl.load_workbook(template_path)
    # 出力シートを取得
    ws = wb[output_sheet_name]
    # データフレームを行に変換
    rows = dataframe_to_rows(df_db, index=False, header=True)
    # 行をシートに書き込む
    for row_index, row in enumerate(rows, 1):
        for column_index, value in enumerate(row, 1):
            ws.cell(row=row_index, column=column_index, value=value)
    # メモリ上に一時的なバッファを作成
    output = io.BytesIO()
    # ワークブックをバッファに保存
    wb.save(output)
    # バッファのデータを取得
    data = output.getvalue()
    # レスポンスを作成
    studentID = str(df_db["学籍番号"].iloc[0])
    student_name = str(df_db["学生氏名"].iloc[0])
    filename = f"output_{studentID}_{student_name}.xlsx"
    response = Response(data, content_type="application/vnd.ms-excel")
    response.headers.set("Content-Disposition", "attachment", filename=quote(filename))
    return response


@app.route("/yomikae", methods=["GET"])
def yomikae():
    db = get_db()
    yomikae_to = session.get("yomikae_to")
    yomikae_from = session.get("yomikae_from")
    ability = session.get("ability")
    yomikae_belonging = session.get("yomikae_belonging")

    # SQLクエリを定義
    sSQL = f"""SELECT "学部学科","免許教科","{yomikae_from}","{yomikae_to}" FROM yomikae WHERE "免許教科"="{ability}" AND "学部学科"="{yomikae_belonging}"  """

    try:
        # SQLクエリを実行してデータフレームを取得
        yomikae_df = pd.read_sql_query(sSQL, db)
    except pd.errors.DatabaseError as e:
        # エラー処理
        error = f"データベースエラー: {e}"
        yomikae_df = pd.DataFrame()
    error = ""
    # テンプレートをレンダリングしてレスポンスを返す
    return render_template("yomikae.html", data=yomikae_df.to_html(), error=error)


@app.route("/import_yomikae")
def import_yomikae():
    # print("インポート開始")
    db_path = current_app.config["DATABASE"]
    conn = sqlite3.connect(db_path)
    # Excelファイルからデータフレームを作成
    df = pd.read_excel("xltemplate\yomikae.xlsx")

    # データフレームをデータベースにインポート
    df.to_sql("yomikae", conn, if_exists="replace", index=False)

    # データベース接続を閉じる
    conn.close()
    # print("インポート成功")
    flash("取込が完了しました。")

    return redirect("/")


@app.route("/export_yomikae")
def export_yomikae():
    # print("エクスポート")
    db_path = current_app.config["DATABASE"]
    conn = sqlite3.connect(db_path)
    # データベースからデータを取得してExcelファイルにエクスポートする処理
    df = pd.read_sql_query("SELECT * FROM yomikae", conn)
    df.to_excel("xltemplate\yomikae.xlsx", index=False)
    conn.close()
    flash(r"出力が完了しました。")
    flash(r"出力したファイルは \xltemplate\yomikae.xlsx です。")
    flash(r"取り込む際は元の読替表が削除され全件分取り込まれます。")
    return index()


# スクリプトとして実行された場合
if __name__ == "__main__":
    # Flaskアプリケーションを起動（デバッグモード）
    app.run(debug=True)
```