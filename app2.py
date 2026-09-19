import io
import os
import streamlit as st
import pandas as pd

# ページ設定
st.set_page_config(page_title="法人会組織・ピックアップシステム", layout="wide")

st.title("🏛️ 法人会 組織・ピックアップシステム")
st.markdown("Python / Streamlit UI版")

# タブの作成（今回はご要望のタブ1をメインに構築）
tab1, tab2, tab3, tab4 = st.tabs([
    "タブ1: 法人会組織ピックアップ",
    "タブ2: 最終リスト編集 ＆ コンビニCSV",
    "タブ3: 領収書PDF発行",
    "タブ4: スマホQR回答状況の確認"
])

# ==========================================
# タブ1: 法人会組織ピックアップ ＆ リスト作成
# ==========================================
with tab1:
    header_title = st.text_input("組織名（タイトルとして表示・出力用）", value="第X回 交流会", key="t1_header_title")
    st.header(f"1. 組織ピックアップ ＆ リスト作成：【 {header_title} 】")

    uploaded_file = st.file_uploader(
        "DBからエクスポートした会員リスト（Excel）をアップロード",
        type=["xlsx", "xls"],
        key="t1_file"
    )

    if uploaded_file is not None:
        try:
            df_members = pd.read_excel(uploaded_file)
            st.success(f"会員データを正常に読み込みました（総件数: {len(df_members)}件）")

            # 列名の前後の空白を削除
            df_members.columns = df_members.columns.str.strip()

            # 必須列の確認
            required_cols = ["組織-大", "組織-小", "法人名", "役職", "役員名", "就任日", "ブロック", "支部"]
            missing_cols = [col for col in required_cols if col not in df_members.columns]

            if missing_cols:
                st.warning(f"⚠️ 以下の列が見つかりません。エクセル側の列名を確認してください: {missing_cols}")
                st.write("現在のエクセル列名一覧:", list(df_members.columns))
            else:
                st.subheader("🔍 組織・エリアによる絞り込み")
                
                # 絞り込み用セレクトボックスの作成
                col1, col2, col3, col4 = st.columns(4)

                with col1:
                    org_large_list = ["すべて"] + list(df_members["組織-大"].dropna().unique())
                    selected_org_large = st.selectbox("組織-大で絞り込み", org_large_list, key="t1_org_l")

                with col2:
                    org_small_list = ["すべて"] + list(df_members["組織-小"].dropna().unique())
                    selected_org_small = st.selectbox("組織-小で絞り込み", org_small_list, key="t1_org_s")

                with col3:
                    blocks = ["すべて"] + list(df_members["ブロック"].dropna().unique())
                    selected_block = st.selectbox("ブロックで絞り込み", blocks, key="t1_b")

                with col4:
                    branches = ["すべて"] + list(df_members["支部"].dropna().unique())
                    selected_branch = st.selectbox("支部で絞り込み", branches, key="t1_br")

                # データのフィルタリング処理
                df_filtered = df_members.copy()
                if selected_org_large != "すべて":
                    df_filtered = df_filtered[df_filtered["組織-大"] == selected_org_large]
                if selected_org_small != "すべて":
                    df_filtered = df_filtered[df_filtered["組織-小"] == selected_org_small]
                if selected_block != "すべて":
                    df_filtered = df_filtered[df_filtered["ブロック"] == selected_block]
                if selected_branch != "すべて":
                    df_filtered = df_filtered[df_filtered["支部"] == selected_branch]

                st.subheader("⚙️ 出力パラメータ設定")
                col_e1, _ = st.columns(2)
                with col_e1:
                    fee = st.number_input("参加費（円） ※9フィールド目に出力されます", value=5000, step=500, key="t1_fe")

                st.subheader("✅ 対象メンバーの選択")
                st.markdown("絞り込まれたリストから、今回対象とするメンバーにチェックを入れてください。")

                # チェックボックス用の列を追加
                df_filtered["選択"] = False
                edited_df = st.data_editor(
                    df_filtered,
                    column_config={
                        "選択": st.column_config.CheckboxColumn("選択", default=False)
                    },
                    disabled=[c for c in df_filtered.columns if c != "選択"],
                    hide_index=True,
                )

                selected_rows = edited_df[edited_df["選択"] == True].copy()
                st.info(f"現在選択されているメンバー数: {len(selected_rows)}件")

                if st.button("📥 組織リスト（Excel）を作成"):
                    if len(selected_rows) == 0:
                        st.warning("メンバーが1人も選択されていません。")
                    else:
                        # 参加費や組織名をデータフレームに追加（9フィールド目等としての出力用）
                        selected_rows["組織名タイトル"] = header_title
                        selected_rows["参加費"] = fee

                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine="openpyxl") as writer:
                            selected_rows.to_excel(writer, index=False, sheet_name="組織リスト")
                        excel_data = output.getvalue()

                        st.download_button(
                            label="💾 組織リストExcelをダウンロード",
                            data=excel_data,
                            file_name="organization_selected_list.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                        st.success("組織リストのExcelファイルが正常に作成されました！")

        except Exception as e:
            st.error(f"ファイルの読み込みまたは処理中にエラーが発生しました: {e}")