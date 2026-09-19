import io
import os
from datetime import date
import pandas as pd
import streamlit as st

# ページ設定
st.set_page_config(
    page_title="法人会組織・イベント管理システム", layout="wide"
)

st.title("🏛️ 法人会 組織・イベント管理システム")
st.markdown(
    "ローカルPython / Streamlit UI版（役員・部会ハイブリッド対応）"
)

# タブの作成
tab1, tab2, tab3, tab4 = st.tabs(
    [
        "タブ1: 組織リスト作成（役員・部会選択）",
        "タブ2: 最終リスト編集 ＆ コンビニCSV",
        "タブ3: 領収書PDF発行",
        "タブ4: スマホQR回答状況の確認",
    ]
)

# ==========================================
# タブ1: 組織リスト作成（役員・部会選択式）
# ==========================================
with tab1:
  header_title = st.text_input(
      "組織名（タイトルとして表示・出力用）",
      value="第X回 交流会",
      key="t1_header_title",
  )
  st.header(f"1. 組織ピックアップ ＆ リスト作成：【 {header_title} 】")

  # リスト種類の選択ラジオボタン
  list_type = st.radio(
      "📁 読み込むリストの種類を選択してください",
      ["役員リスト", "部会リスト"],
      horizontal=True,
      key="t1_list_type",
  )

  uploaded_file = st.file_uploader(
      f"DBからエクスポートした【 {list_type} 】のExcelファイルをアップロード",
      type=["xlsx", "xls"],
      key="t1_file",
  )

  if uploaded_file is not None:
    try:
      df_members = pd.read_excel(uploaded_file)
      st.success(
          f"データを正常に読み込みました（総件数: {len(df_members)}件）"
      )

      # 列名の前後の空白を削除
      df_members.columns = df_members.columns.str.strip()

      # ==========================================
      # 【パターンA】 役員リストの場合
      # ==========================================
      if list_type == "役員リスト":
        required_cols = [
            "組織-大",
            "組織-小",
            "法人名",
            "役職",
            "役員名",
            "就任日",
            "ブロック",
            "支部",
            "郵便番号",
            "住所",
        ]
        missing_cols = [
            col for col in required_cols if col not in df_members.columns
        ]

        if missing_cols:
          st.warning(
              "⚠️ 以下の列が見つかりません。エクセル側の列名を確認してください:"
              f" {missing_cols}"
          )
          st.write("現在のエクセル列名一覧:", list(df_members.columns))
        else:
          st.subheader("🔍 組織・エリアによる絞り込み（役員）")
          col1, col2, col3, col4 = st.columns(4)

          with col1:
            org_large_list = ["すべて"] + list(
                df_members["組織-大"].dropna().unique()
            )
            selected_org_large = st.selectbox(
                "組織-大で絞り込み", org_large_list, key="t1_y_org_l"
            )

          with col2:
            org_small_list = ["すべて"] + list(
                df_members["組織-小"].dropna().unique()
            )
            selected_org_small = st.selectbox(
                "組織-小で絞り込み", org_small_list, key="t1_y_org_s"
            )

          with col3:
            blocks = ["すべて"] + list(df_members["ブロック"].dropna().unique())
            selected_block = st.selectbox(
                "ブロックで絞り込み", blocks, key="t1_y_b"
            )

          with col4:
            branches = ["すべて"] + list(df_members["支部"].dropna().unique())
            selected_branch = st.selectbox(
                "支部で絞り込み", branches, key="t1_y_br"
            )

          # フィルタリング
          df_filtered = df_members.copy()
          if selected_org_large != "すべて":
            df_filtered = df_filtered[
                df_filtered["組織-大"] == selected_org_large
            ]
          if selected_org_small != "すべて":
            df_filtered = df_filtered[
                df_filtered["組織-小"] == selected_org_small
            ]
          if selected_block != "すべて":
            df_filtered = df_filtered[df_filtered["ブロック"] == selected_block]
          if selected_branch != "すべて":
            df_filtered = df_filtered[df_filtered["支部"] == selected_branch]

          st.subheader("⚙️ 出力パラメータ設定")
          col_e1, _ = st.columns(2)
          with col_e1:
            fee = st.number_input(
                "参加費（円） ※9フィールド目（相当）に出力されます",
                value=5000,
                step=500,
                key="t1_y_fe",
            )

          st.subheader("✅ 対象メンバーの選択")
          df_filtered["選択"] = False
          edited_df = st.data_editor(
              df_filtered,
              column_config={
                  "選択": st.column_config.CheckboxColumn(
                      "選択", default=False
                  )
              },
              disabled=[c for c in df_filtered.columns if c != "選択"],
              hide_index=True,
          )

          selected_rows = edited_df[edited_df["選択"] == True].copy()
          st.info(f"現在選択されている役員数: {len(selected_rows)}件")

          if st.button("📥 役員組織リスト（Excel）を作成", key="t1_y_btn"):
            if len(selected_rows) == 0:
              st.warning("メンバーが1人も選択されていません。")
            else:
              selected_rows["組織名タイトル"] = header_title
              selected_rows["参加費"] = fee

              output = io.BytesIO()
              with pd.ExcelWriter(output, engine="openpyxl") as writer:
                selected_rows.to_excel(
                    writer, index=False, sheet_name="役員リスト"
                )
              excel_data = output.getvalue()

              st.download_button(
                  label="💾 役員組織リストExcelをダウンロード",
                  data=excel_data,
                  file_name="officer_selected_list.xlsx",
                  mime=(
                      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                  ),
              )
              st.success("役員組織リストのExcelファイルを作成しました！")

      # ==========================================
      # 【パターンB】 部会リストの場合
      # ==========================================
      elif list_type == "部会リスト":
        required_cols = [
            "部会名",
            "法人名",
            "部会担当者",
            "入部日",
            "退部日",
            "部員区分",
            "郵便番号",
            "住所",
            "整理番号CO",
        ]
        missing_cols = [
            col for col in required_cols if col not in df_members.columns
        ]

        if missing_cols:
          st.warning(
              "⚠️ 以下の列が見つかりません。エクセル側の列名を確認してください:"
              f" {missing_cols}"
          )
          st.write("現在のエクセル列名一覧:", list(df_members.columns))
        else:
          st.subheader("🔍 部会による絞り込み")
          col1, _ = st.columns(2)

          with col1:
            club_list = ["すべて"] + list(df_members["部会名"].dropna().unique())
            selected_club = st.selectbox(
                "部会名で絞り込み", club_list, key="t1_b_club"
            )

          # フィルタリング
          df_filtered = df_members.copy()
          if selected_club != "すべて":
            df_filtered = df_filtered[df_filtered["部会名"] == selected_club]

          st.subheader("⚙️ 出力パラメータ設定")
          col_e1, _ = st.columns(2)
          with col_e1:
            fee = st.number_input(
                "参加費（円） ※出力データに追加されます",
                value=5000,
                step=500,
                key="t1_b_fe",
            )

          st.subheader("✅ 対象メンバーの選択")
          df_filtered["選択"] = False
          edited_df = st.data_editor(
              df_filtered,
              column_config={
                  "選択": st.column_config.CheckboxColumn(
                      "選択", default=False
                  )
              },
              disabled=[c for c in df_filtered.columns if c != "選択"],
              hide_index=True,
          )

          selected_rows = edited_df[edited_df["選択"] == True].copy()
          st.info(f"現在選択されている部会員数: {len(selected_rows)}件")

          if st.button("📥 部会組織リスト（Excel）を作成", key="t1_b_btn"):
            if len(selected_rows) == 0:
              st.warning("メンバーが1人も選択されていません。")
            else:
              selected_rows["組織名タイトル"] = header_title
              selected_rows["参加費"] = fee

              output = io.BytesIO()
              with pd.ExcelWriter(output, engine="openpyxl") as writer:
                selected_rows.to_excel(
                    writer, index=False, sheet_name="部会リスト"
                )
              excel_data = output.getvalue()

              st.download_button(
                  label="💾 部会組織リストExcelをダウンロード",
                  data=excel_data,
                  file_name="club_selected_list.xlsx",
                  mime=(
                      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                  ),
              )
              st.success("部会組織リストのExcelファイルを作成しました！")

    except Exception as e:
      st.error(f"ファイルの読み込みまたは処理中にエラーが発生しました: {e}")

# ==========================================
# タブ2: 最終リスト編集 ＆ コンビニCSV出力
# ==========================================
with tab2:
  st.header("2. 最終参加者リスト編集 ＆ コンビニ収納用WEB-EB CSV出力")
  st.markdown(
      "スマホ回答分およびFAX回答分をまとめた最終リストを読み込み・編集し、コンビニ収納用CSVを出力します。"
  )

  uploaded_file_tab2 = st.file_uploader(
      "最終参加者リスト（Excel）をアップロード",
      type=["xlsx", "xls"],
      key="tab2_file",
  )

  if uploaded_file_tab2 is not None:
    df_tab2 = pd.read_excel(uploaded_file_tab2)

    st.subheader("📝 リストの最終確認・手動調整")
    edited_tab2 = st.data_editor(df_tab2, num_rows="dynamic", hide_index=True)

    if st.button("📤 コンビニ収納用WEB-EB CSVを生成", key="t2_csv_btn"):
      csv_data = edited_tab2.to_csv(
          index=False, encoding="cp932", errors="replace"
      )

      st.download_button(
          label="💾 WEB-EB用 CSVファイルをダウンロード",
          data=csv_data,
          file_name="convenience_store_web_eb.csv",
          mime="text/csv",
          key="t2_dl_btn",
      )
      st.success("CSVファイルを生成しました（Shift_JIS形式）")

# ==========================================
# タブ3: 領収書PDF発行
# ==========================================
with tab3:
  st.header("3. 領収書PDF発行（A4用紙上下2段・切り取り形式）")
  st.markdown(
      "参加者リストのExcelをアップロードして、ブルー基調の綺麗な領収書をA4用紙1枚につき2件ずつ作成できます。"
  )

  uploaded_file_tab3 = st.file_uploader(
      "領収書発行用リスト（Excel）をアップロード",
      type=["xlsx", "xls"],
      key="tab3_file",
  )

  st.subheader("⚙️ 領収書共通設定")
  col_r1, col_r2 = st.columns(2)
  with col_r1:
    receipt_event_name = st.text_input(
        "但し書き（イベント名等）",
        value="第X回 交流会 参加費として",
        key="t3_ev",
    )
  with col_r2:
    receipt_date = st.date_input("発行日", value=date.today(), key="t3_dt")

  default_fee = st.number_input(
      "既定の金額（円）", value=5000, step=500, key="t3_fee"
  )

  if uploaded_file_tab3 is not None:
    df_rec = pd.read_excel(uploaded_file_tab3)
    df_rec.columns = df_rec.columns.str.strip()

    st.subheader("📝 発行対象の確認・金額調整")
    if "金額" not in df_rec.columns:
      df_rec["金額"] = default_fee

    edited_rec_df = st.data_editor(df_rec, num_rows="dynamic", hide_index=True)

    if st.button("📄 A4 2段切り取り式 領収書PDFを生成"):
      from reportlab.lib.pagesizes import A4, mm
      from reportlab.pdfgen import canvas

      pdf_buffer = io.BytesIO()
      a4_w, a4_h = A4
      c = canvas.Canvas(pdf_buffer, pagesize=A4)

      records = list(edited_rec_df.iterrows())
      total_records = len(records)

      for i, (idx, row) in enumerate(records):
        pos_in_page = i % 2
        y_base = (
            a4_h - 10 * mm if pos_in_page == 0 else (a4_h / 2.0) - 5 * mm
        )

        corp_name = str(row["法人名"]) if "法人名" in row else "宛名不明"
        amt = (
            row["金額"]
            if "金額" in row and pd.notna(row["金額"])
            else default_fee
        )

        c.drawString(20 * mm, y_base - 28 * mm, "領 収 書")
        c.drawString(20 * mm, y_base - 42 * mm, f"{corp_name}  様")
        c.drawString(25 * mm, y_base - 55 * mm, f"金額: ￥{int(amt):,} - (税込)")

        if pos_in_page == 0 and i < total_records - 1:
          mid_y = a4_h / 2.0
          c.setDash(4, 4)
          c.line(15 * mm, mid_y, a4_w - 15 * mm, mid_y)
          c.setDash()

        if pos_in_page == 1 or i == total_records - 1:
          c.showPage()

      c.save()
      pdf_buffer.seek(0)
      st.download_button(
          label="📥 2段切り取り式 領収書PDFをダウンロード",
          data=pdf_buffer,
          file_name="receipts_2up.pdf",
          mime="application/pdf",
      )
      st.success("領収書PDFの生成が完了しました！")

# ==========================================
# タブ4: スマホQR回答状況の確認 ＆ CSV出力
# ==========================================
with tab4:
  st.header("4. スマホQR回答状況の確認 ＆ CSV出力")
  st.markdown(
      "スマホからQRコード経由で回答されたリアルタイムの集計結果をスプレッドシートから自動取得し、確認・CSV出力できます。"
  )

  sheet_csv_url = st.text_input(
      "Googleスプレッドシート「回答結果」のCSV公開リンク",
      value="",
      placeholder="https://docs.google.com/spreadsheets/d/.../export?format=csv&gid=0",
      key="t4_url",
  )

  df_tab4 = None
  if sheet_csv_url:
    if st.button("🔄 スプレッドシートから回答結果を読み込む", key="t4_load_btn"):
      try:
        df_tab4 = pd.read_csv(sheet_csv_url)
        st.session_state["df_qr_responses"] = df_tab4
        st.success(f"回答データを取得しました（件数: {len(df_tab4)}件）")
      except Exception as e:
        st.error(f"データの読み込みに失敗しました: {e}")

  if "df_qr_responses" in st.session_state:
    df_tab4 = st.session_state["df_qr_responses"]

  if df_tab4 is not None:
    st.subheader("📝 スマホ回答一覧の確認・編集")
    edited_tab4 = st.data_editor(df_tab4, num_rows="dynamic", hide_index=True)

    if st.button("📤 スマホ回答分のコンビニ収納CSVを生成", key="t4_csv_btn"):
      csv_data_t4 = edited_tab4.to_csv(
          index=False, encoding="cp932", errors="replace"
      )
      st.download_button(
          label="💾 スマホ回答用 WEB-EB CSVファイルをダウンロード",
          data=csv_data_t4,
          file_name="qr_responses_web_eb.csv",
          mime="text/csv",
          key="t4_dl_btn",
      )
      st.success("CSVファイルを生成しました（Shift_JIS形式）")
  else:
    st.markdown(
        "上の入力欄にスプレッドシートのリンクを入力し、読み込みボタンを押してください。"
    )