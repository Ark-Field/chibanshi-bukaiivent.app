import io
import os
from datetime import date
import pandas as pd
import streamlit as st

# ReportLabのインポート
from reportlab.lib.colors import HexColor
from reportlab.lib.pagesizes import A4, landscape, portrait
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# --- 日本語フォント（Noto Sans JP）の登録 ---
try:
  font_path = "NotoSansJP-Regular.ttf"
  if os.path.exists(font_path):
    pdfmetrics.registerFont(TTFont("JapaneseFont", font_path))
  else:
    alt_path = r"C:\Windows\Fonts\meiryo.ttc"
    if os.path.exists(alt_path):
      pdfmetrics.registerFont(TTFont("JapaneseFont", alt_path, subfontIndex=0))
except Exception as e:
  print(f"フォント登録エラー: {e}")

# ページ設定
st.set_page_config(
    page_title="法人会組織・イベント管理システム", layout="wide"
)

st.title("🏛️ 法人会 組織・イベント管理システム")
st.markdown("ローカルPython / Streamlit UI版（タブ1自由絞り込み対応）")

# タブの作成
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
    [
        "タブ1: 組織リスト作成",
        "タブ2: 最終リスト ＆ CSV",
        "タブ3: 領収書PDF",
        "タブ4: スマホQR確認",
        "タブ5: カスタム表作成 (PDF&Excel)",
        "タブ6: 宛名ラベル印刷 (12面)",
    ]
)

# ==========================================
# タブ1: 組織リスト作成（自由絞り込み・全選択対応）
# ==========================================
with tab1:
  header_title = st.text_input(
      "組織名（タイトルとして表示・出力用）",
      value="イベント名",
      key="t1_header_title",
  )
  st.header(f"1. 組織ピックアップ ＆ リスト作成：【 {header_title} 】")

  list_type_t1 = st.radio(
      "📁 読み込むリストの種類を選択してください",
      ["役員リスト", "部会リスト", "会員リスト"],
      horizontal=True,
      key="t1_list_type",
  )

  uploaded_file_t1 = st.file_uploader(
      f"DBからエクスポートした【 {list_type_t1} 】のExcelファイルをアップロード",
      type=["xlsx", "xls"],
      key="t1_file",
  )

  if uploaded_file_t1 is not None:
    try:
      df_t1 = pd.read_excel(uploaded_file_t1)
      df_t1.columns = df_t1.columns.str.strip()
      st.success(
          f"データを正常に読み込みました（総件数: {len(df_t1)}件）"
      )

      # --- 自由絞り込み機能の追加 ---
      st.subheader("🔍 データの絞り込み検索")
      col_f1, col_f2 = st.columns(2)

      all_t1_cols = list(df_t1.columns)
      filter_col_options = ["(絞り込みなし)"] + all_t1_cols

      with col_f1:
        selected_filter_col = st.selectbox(
            "絞り込みを行うフィールド名",
            filter_col_options,
            key="t1_filter_col",
        )

      df_filtered = df_t1.copy()
      if selected_filter_col != "(絞り込みなし)":
        # 選択された列のユニークな値（空欄除外）を取得
        unique_vals = [
            str(v)
            for v in df_t1[selected_filter_col].dropna().unique()
            if str(v).strip() != ""
        ]
        val_options = ["(すべて表示)"] + unique_vals

        with col_f2:
          selected_filter_val = st.selectbox(
              f"「{selected_filter_col}」の値を選択",
              val_options,
              key="t1_filter_val",
          )

        if selected_filter_val != "((すべて表示))" and selected_filter_val != "(すべて表示)":
          # 数値型か文字列型かに配慮して一致する行を抽出
          df_filtered = df_filtered[
              df_filtered[selected_filter_col].astype(str)
              == str(selected_filter_val)
          ]
        st.info(
            f"絞り込み結果: {len(df_filtered)}件 （全{len(df_t1)}件中）"
        )

      # タブ1用セッションステート（全選択・全解除の制御）
      if "select_all_t1" not in st.session_state:
        st.session_state["select_all_t1"] = True

      col_b1, col_b2, _ = st.columns([1, 1, 4])
      with col_b1:
        if st.button("☑️ すべて選択", key="t1_btn_all"):
          st.session_state["select_all_t1"] = True
          st.rerun()
      with col_b2:
        if st.button("☐ すべて解除", key="t1_btn_clear"):
          st.session_state["select_all_t1"] = False
          st.rerun()

      df_filtered["選択"] = st.session_state["select_all_t1"]

      edited_t1 = st.data_editor(
          df_filtered,
          column_config={
              "選択": st.column_config.CheckboxColumn("選択", default=True)
          },
          disabled=[c for c in df_filtered.columns if c != "選択"],
          hide_index=True,
          num_rows="dynamic",
      )

      selected_t1 = edited_t1[edited_t1["選択"] == True].copy()
      st.info(f"現在選択されている件数: {len(selected_t1)}件")

      fee_t1 = st.number_input(
          "参加費（円） ※出力データ用", value=5000, step=500, key="t1_fee"
      )

      if st.button("📥 組織リストExcelを作成", key="t1_excel_btn"):
        if len(selected_t1) == 0:
          st.warning("メンバーが1人も選択されていません。")
        else:
          selected_t1["組織名タイトル"] = header_title
          selected_t1["参加費"] = fee_t1
          output = io.BytesIO()
          with pd.ExcelWriter(output, engine="openpyxl") as writer:
            selected_t1.to_excel(writer, index=False, sheet_name="リスト")
          st.download_button(
              label="💾 組織リストExcelをダウンロード",
              data=output.getvalue(),
              file_name="organization_list.xlsx",
              mime=(
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              ),
          )
          st.success("Excelファイルを作成しました！")
    except Exception as e:
      st.error(f"エラー: {e}")

# ==========================================
# タブ2: 最終リスト編集 ＆ コンビニCSV出力
# ==========================================
with tab2:
  st.header("2. 最終参加者リスト編集 ＆ コンビニ収納用WEB-EB CSV出力")
  uploaded_file_tab2 = st.file_uploader(
      "最終参加者リスト（Excel）をアップロード",
      type=["xlsx", "xls"],
      key="tab2_file",
  )

  if uploaded_file_tab2 is not None:
    df_tab2 = pd.read_excel(uploaded_file_tab2)
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
  uploaded_file_tab3 = st.file_uploader(
      "領収書発行用リスト（Excel）をアップロード",
      type=["xlsx", "xls"],
      key="tab3_file",
  )
  receipt_event_name = st.text_input(
      "但し書き", value="第X回 交流会 参加費として", key="t3_ev"
  )
  receipt_date = st.date_input("発行日", value=date.today(), key="t3_dt")
  default_fee = st.number_input(
      "既定の金額（円）", value=5000, step=500, key="t3_fee"
  )

  if uploaded_file_tab3 is not None:
    df_rec = pd.read_excel(uploaded_file_tab3)
    df_rec.columns = df_rec.columns.str.strip()
    if "金額" not in df_rec.columns:
      df_rec["金額"] = default_fee
    edited_rec_df = st.data_editor(df_rec, num_rows="dynamic", hide_index=True)

    if st.button("📄 A4 2段切り取り式 領収書PDFを生成"):
      pdf_buffer = io.BytesIO()
      a4_w_v, a4_h_v = A4
      c = canvas.Canvas(pdf_buffer, pagesize=A4)
      records = list(edited_rec_df.iterrows())
      total_records = len(records)

      for i, (idx, row) in enumerate(records):
        pos_in_page = i % 2
        y_base = (
            a4_h_v - 10 * mm
            if pos_in_page == 0
            else (a4_h_v / 2.0) - 5 * mm
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
          mid_y = a4_h_v / 2.0
          c.setDash(4, 4)
          c.line(15 * mm, mid_y, a4_w_v - 15 * mm, mid_y)
          c.setDash()
        if pos_in_page == 1 or i == total_records - 1:
          c.showPage()

      c.save()
      pdf_buffer.seek(0)
      st.download_button(
          label="📥 領収書PDFをダウンロード",
          data=pdf_buffer,
          file_name="receipts_2up.pdf",
          mime="application/pdf",
      )
      st.success("領収書PDFを作成しました！")

# ==========================================
# タブ4: スマホQR回答状況の確認 ＆ CSV出力
# ==========================================
with tab4:
  st.header("4. スマホQR回答状況の確認 ＆ CSV出力")
  sheet_csv_url = st.text_input(
      "Googleスプレッドシート「回答結果」のCSV公開リンク",
      value="",
      placeholder="https://docs.google.com/spreadsheets/d/.../export?format=csv&gid=0",
      key="t4_url",
  )

  df_tab4 = None
  if sheet_csv_url and st.button(
      "🔄 スプレッドシートから読み込む", key="t4_load_btn"
  ):
    try:
      df_tab4 = pd.read_csv(sheet_csv_url)
      st.session_state["df_qr_responses"] = df_tab4
      st.success(f"取得成功（件数: {len(df_tab4)}件）")
    except Exception as e:
      st.error(f"エラー: {e}")

  if "df_qr_responses" in st.session_state:
    df_tab4 = st.session_state["df_qr_responses"]

  if df_tab4 is not None:
    edited_tab4 = st.data_editor(df_tab4, num_rows="dynamic", hide_index=True)
    if st.button("📤 スマホ回答分CSVを生成", key="t4_csv_btn"):
      csv_data_t4 = edited_tab4.to_csv(
          index=False, encoding="cp932", errors="replace"
      )
      st.download_button(
          label="💾 WEB-EB CSVをダウンロード",
          data=csv_data_t4,
          file_name="qr_responses_web_eb.csv",
          mime="text/csv",
          key="t4_dl_btn",
      )
      st.success("CSVを作成しました！")

# ==========================================
# タブ5: カスタム表作成（PDF & Excel両対応）
# ==========================================
with tab5:
  st.header("5. カスタム表作成（PDF ＆ Excel出力）")
  st.markdown(
      "役員・部会・会員のデータを読み込み、表示する項目、グループ化、並び順を自由にカスタムしてPDFとExcelを出力できます。"
  )

  custom_title = st.text_input(
      "表のタイトル",
      value="令和X年度 千葉西法人会 カスタム名簿",
      key="t5_title",
  )

  list_type_t5 = st.radio(
      "📁 読み込むデータの種類を選択",
      ["役員リスト", "部会リスト", "会員リスト"],
      horizontal=True,
      key="t5_type",
  )

  uploaded_file_t5 = st.file_uploader(
      f"【 {list_type_t5} 】のExcelファイルをアップロード",
      type=["xlsx", "xls"],
      key="t5_file",
  )

  if uploaded_file_t5 is not None:
    try:
      df_custom = pd.read_excel(uploaded_file_t5)
      df_custom.columns = df_custom.columns.str.strip()
      st.success(
          f"データを正常に読み込みました（総件数: {len(df_custom)}件）"
      )

      st.subheader("⚙️ 1. 表示項目（フィールド）の選択と並び替え・非表示")
      st.markdown(
          "出力したい項目にチェックを入れ、順番を整えてください（不要な項目はチェックを外します）。"
      )

      all_cols = list(df_custom.columns)

      col_config_df = pd.DataFrame({
          "フィールド名": all_cols,
          "表示": [True] * len(all_cols),
      })
      edited_config = st.data_editor(
          col_config_df, hide_index=True, key="t5_col_conf"
      )

      visible_cols = edited_config[edited_config["表示"] == True][
          "フィールド名"
      ].tolist()

      st.subheader("⚙️ 2. グループ化（見出し）と並び順の設定")
      col_s1, col_s2 = st.columns(2)

      with col_s1:
        group_options = ["なし"] + visible_cols
        selected_group = st.selectbox(
            "グループ化する項目（大見出し）", group_options, key="t5_group"
        )

      with col_s2:
        sort_options = ["指定なし"] + visible_cols
        selected_sort = st.selectbox(
            "並び順の基準とする項目", sort_options, key="t5_sort"
        )

      df_process = df_custom.copy()
      if selected_sort != "指定なし" and selected_sort in df_process.columns:
        df_process = df_process.sort_values(by=selected_sort)

      st.markdown("---")
      col_dl1, col_dl2 = st.columns(2)

      # --- Excel出力処理 ---
      with col_dl1:
        if st.button("📥 カスタムExcelをダウンロード", key="t5_excel_btn"):
          if not visible_cols:
            st.warning("表示する項目が1つも選択されていません。")
          else:
            export_df = df_process[visible_cols]
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
              export_df.to_excel(
                  writer, index=False, sheet_name="カスタムリスト"
              )
            st.download_button(
                label="💾 Excel (.xlsx) ファイルをダウンロード",
                data=output.getvalue(),
                file_name="custom_organization_list.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
            )
            st.success("カスタムExcelファイルを作成しました！")

      # --- PDF出力処理 ---
      with col_dl2:
        if st.button("📄 カスタムA4横PDFをダウンロード", key="t5_pdf_btn"):
          if not visible_cols:
            st.warning("表示する項目が1つも選択されていません。")
          else:
            try:
              pdf_buffer = io.BytesIO()
              doc = SimpleDocTemplate(
                  pdf_buffer,
                  pagesize=landscape(A4),
                  rightMargin=10 * mm,
                  leftMargin=10 * mm,
                  topMargin=15 * mm,
                  bottomMargin=15 * mm,
              )

              story = []
              styles = getSampleStyleSheet()

              title_style = ParagraphStyle(
                  'TitleStyle',
                  parent=styles['Normal'],
                  fontName='JapaneseFont',
                  fontSize=14,
                  leading=18,
                  alignment=1,
                  textColor=HexColor('#1A5276'),
              )
              group_header_style = ParagraphStyle(
                  'GroupHeaderStyle',
                  parent=styles['Normal'],
                  fontName='JapaneseFont',
                  fontSize=10,
                  leading=14,
                  textColor=HexColor('#FFFFFF'),
              )
              cell_style = ParagraphStyle(
                  'CellStyle',
                  parent=styles['Normal'],
                  fontName='JapaneseFont',
                  fontSize=8,
                  leading=11,
                  textColor=HexColor('#2C3E50'),
              )
              header_cell_style = ParagraphStyle(
                  'HeaderCellStyle',
                  parent=styles['Normal'],
                  fontName='JapaneseFont',
                  fontSize=8,
                  leading=11,
                  alignment=1,
                  textColor=HexColor('#FFFFFF'),
              )

              story.append(Paragraph(custom_title, title_style))
              story.append(Spacer(1, 10))

              page_width = 277 * mm
              col_w = page_width / max(len(visible_cols), 1)
              col_widths = [col_w] * len(visible_cols)

              if selected_group != "なし" and selected_group in df_process.columns:
                grouped = df_process.groupby(selected_group)
                for g_name, group_df in grouped:
                  g_title = (
                      f"■ {selected_group}: {g_name} （人数:"
                      f" {len(group_df)}名）"
                  )
                  g_table_data = [[
                      Paragraph(g_title, group_header_style)
                  ] + [""] * (len(visible_cols) - 1)]
                  g_table_data.append(
                      [
                          Paragraph(h, header_cell_style)
                          for h in visible_cols
                      ]
                  )

                  for _, r in group_df.iterrows():
                    g_table_data.append([
                        Paragraph(str(r.get(h, '')), cell_style)
                        for h in visible_cols
                    ])

                  t = Table(g_table_data, colWidths=col_widths, repeatRows=2)
                  t.setStyle(TableStyle([
                      ('SPAN', (0, 0), (-1, 0)),
                      ('BACKGROUND', (0, 0), (-1, 0), HexColor('#1A5276')),
                      ('BACKGROUND', (0, 1), (-1, 1), HexColor('#2980B9')),
                      ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                      ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                      ('GRID', (0, 0), (-1, -1), 0.5, HexColor('#BDC3C7')),
                      ('TOPPADDING', (0, 0), (-1, -1), 3),
                      ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                      ('LEFTPADDING', (0, 0), (-1, -1), 4),
                      ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                  ]))
                  story.append(t)
                  story.append(Spacer(1, 8))
              else:
                table_data = [[
                    Paragraph(h, header_cell_style) for h in visible_cols
                ]]
                for _, r in df_process.iterrows():
                  table_data.append([
                      Paragraph(str(r.get(h, '')), cell_style)
                      for h in visible_cols
                  ])

                t = Table(table_data, colWidths=col_widths, repeatRows=1)
                t.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), HexColor('#1A5276')),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('GRID', (0, 0), (-1, -1), 0.5, HexColor('#BDC3C7')),
                    ('TOPPADDING', (0, 0), (-1, -1), 3),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                    ('LEFTPADDING', (0, 0), (-1, -1), 4),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                ]))
                story.append(t)

              doc.build(story)
              pdf_buffer.seek(0)

              st.download_button(
                  label="📥 カスタムA4横PDFをダウンロード",
                  data=pdf_buffer,
                  file_name="custom_roster.landscape.pdf",
                  mime="application/pdf",
              )
              st.success("カスタムA4横PDFを生成しました！")
            except Exception as pdf_ex:
              st.error(f"PDF生成エラー: {pdf_ex}")

    except Exception as e:
      st.error(f"ファイル読み込みエラー: {e}")

# ==========================================
# タブ6: 宛名ラベル印刷（A4・12面付け）
# ==========================================
with tab6:
  st.header("6. 宛名ラベル印刷（A4サイズ・12面付け）")
  st.markdown(
      "会員リストなどのExcelを読み込み、市販の12面タックシール（86×42mm）にぴったりの宛名ラベルPDFを作成します。"
  )

  uploaded_file_t6 = st.file_uploader(
      "宛名ラベル用Excelファイルをアップロード",
      type=["xlsx", "xls"],
      key="t6_file",
  )

  if uploaded_file_t6 is not None:
    try:
      df_label = pd.read_excel(uploaded_file_t6)
      df_label.columns = df_label.columns.str.strip()
      st.success(
          f"ラベルデータを読み込みました（総件数: {len(df_label)}件）"
      )

      all_label_cols = list(df_label.columns)
      sub_field_options = ["(なし)"] + all_label_cols

      col_sub1, _ = st.columns(2)
      with col_sub1:
        selected_sub_field = st.selectbox(
            "🏷️ 名前（様）の下にカッコ書きで追加表示するフィールド",
            sub_field_options,
            key="t6_sub_field",
        )

      st.subheader("📝 宛名データの確認・選択")
      df_label["印刷"] = True
      edited_label_df = st.data_editor(
          df_label, hide_index=True, num_rows="dynamic", key="t6_editor"
      )
      selected_labels = edited_label_df[
          edited_label_df.get("印刷", True) == True
      ].copy()

      st.info(f"印刷対象のラベル数: {len(selected_labels)}件")

      if st.button("🏷️ 12面宛名ラベルPDFを生成する", key="t6_pdf_btn"):
        if len(selected_labels) == 0:
          st.warning("印刷対象が選択されていません。")
        else:
          try:
            pdf_buffer = io.BytesIO()
            doc = SimpleDocTemplate(
                pdf_buffer,
                pagesize=portrait(A4),
                rightMargin=8 * mm,
                leftMargin=8 * mm,
                topMargin=12 * mm,
                bottomMargin=12 * mm,
            )

            story = []
            styles = getSampleStyleSheet()

            label_style = ParagraphStyle(
                'LabelStyle',
                parent=styles['Normal'],
                fontName='JapaneseFont',
                fontSize=9,
                leading=13,
                textColor=HexColor('#000000'),
            )

            cell_width = 62 * mm
            cell_height = 42 * mm
            col_widths = [cell_width, cell_width, cell_width]

            current_page_data = []
            row_cells = []

            for idx, r in selected_labels.iterrows():
              post_code = (
                  str(r.get("宛郵便番号", r.get("郵便番号", "")))
                  if pd.notna(r.get("宛郵便番号", r.get("郵便番号", "")))
                  else ""
              )
              address = (
                  str(r.get("住所1（新）", r.get("住所", "")))
                  if pd.notna(r.get("住所1（新）", r.get("住所", "")))
                  else ""
              )
              corp = (
                  str(r.get("法人名", ""))
                  if pd.notna(r.get("法人名", ""))
                  else ""
              )
              member_name = (
                  str(r.get("役員名", r.get("部会担当者", "")))
                  if pd.notna(r.get("役員名", r.get("部会担当者", "")))
                  else ""
              )

              sub_text_str = ""
              if (
                  selected_sub_field != "(なし)"
                  and selected_sub_field in r
                  and pd.notna(r[selected_sub_field])
              ):
                val = str(r[selected_sub_field]).strip()
                if val:
                  sub_text_str = f"（{val}）"

              text_content = f"<b>〒 {post_code}</b><br/>{address}<br/><br/><b>{corp}</b><br/>{member_name} 様 {sub_text_str}"
              p = Paragraph(text_content, label_style)
              row_cells.append(p)

              if len(row_cells) == 3:
                current_page_data.append(row_cells)
                row_cells = []

                if len(current_page_data) == 4:
                  t = Table(
                      current_page_data,
                      colWidths=col_widths,
                      rowHeights=[cell_height] * 4,
                  )
                  t.setStyle(TableStyle([
                      ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                      ('LEFTPADDING', (0, 0), (-1, -1), 4),
                      ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                      ('TOPPADDING', (0, 0), (-1, -1), 4),
                      ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                      ('BOX', (0, 0), (-1, -1), 0.2, HexColor('#E5E7E9')),
                      ('GRID', (0, 0), (-1, -1), 0.2, HexColor('#E5E7E9')),
                  ]))
                  story.append(t)
                  story.append(Spacer(1, 0))
                  current_page_data = []

            if len(row_cells) > 0:
              while len(row_cells) < 3:
                row_cells.append(Paragraph("", label_style))
              current_page_data.append(row_cells)

            if len(current_page_data) > 0:
              while len(current_page_data) < 4:
                current_page_data.append(
                    [
                        Paragraph("", label_style),
                        Paragraph("", label_style),
                        Paragraph("", label_style),
                    ]
                )
              t = Table(
                  current_page_data,
                  colWidths=col_widths,
                  rowHeights=[cell_height] * 4,
              )
              t.setStyle(TableStyle([
                  ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                  ('LEFTPADDING', (0, 0), (-1, -1), 4),
                  ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                  ('TOPPADDING', (0, 0), (-1, -1), 4),
                  ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                  ('BOX', (0, 0), (-1, -1), 0.2, HexColor('#E5E7E9')),
                  ('GRID', (0, 0), (-1, -1), 0.2, HexColor('#E5E7E9')),
              ]))
              story.append(t)

            doc.build(story)
            pdf_buffer.seek(0)

            st.download_button(
                label="📥 12面宛名ラベルPDFをダウンロード",
                data=pdf_buffer,
                file_name="address_labels_12up.pdf",
                mime="application/pdf",
            )
            st.success("12面宛名ラベルのPDF生成が完了しました！")
          except Exception as label_ex:
            st.error(f"ラベルPDF生成中にエラーが発生しました: {label_ex}")

    except Exception as e:
      st.error(f"会員ファイル読み込みエラー: {e}")