import streamlit as st
import pandas as pd
import io

# --- Streamlit 應用程式介面 ---
st.set_page_config(page_title="Excel 合併工具", page_icon="🧩", layout="wide")

st.title("🧩 Excel 合併工具")

# 上傳多個 Excel 檔案
uploaded_files = st.file_uploader("請上傳您的 Excel 檔案（可上傳多個檔案）", type=["xlsx"], accept_multiple_files=True)

# 用來存儲用戶選擇的工作表
selected_sheets = {}

if uploaded_files:
    # 為每個上傳的檔案顯示多選框供用戶選擇工作表
    for uploaded_file in uploaded_files:
        try:
            # 讀取每個 Excel 檔案的工作表名稱
            excel_data = pd.ExcelFile(uploaded_file)
            sheet_names = excel_data.sheet_names

            # 顯示多選框供用戶選擇多個工作表
            selected_sheets[uploaded_file.name] = st.multiselect(
                f"請選擇檔案 `{uploaded_file.name}` 中的工作表進行合併：",
                options=sheet_names,
                default=sheet_names  # 預設選中所有工作表
            )
        except Exception as e:
            st.error(f"無法讀取檔案 {uploaded_file.name} 的工作表，請檢查檔案格式。錯誤：{e}")

    # 合併模式選擇
    merge_mode = st.radio(
        "合併模式設定",
        options=["縱向合併 (上下堆疊)", "橫向合併 (左右拼接)"],
        index=0
    )

    # 動態選項：根據合併模式顯示不同的選項
    if merge_mode == "縱向合併 (上下堆疊)":
        # 縱向合併：選擇表頭所在行
        st.write("請選擇表頭所在的行（1 表示第一行）：")
        header_row = st.number_input("表頭所在行：", min_value=1, max_value=100, value=1, step=1)

    elif merge_mode == "橫向合併 (左右拼接)":
        # 橫向合併：選擇需要合併的欄位
        try:
            # 從第一個檔案中提取第一個工作表的欄位
            first_file = uploaded_files[0]
            first_sheet = selected_sheets[first_file.name][0]  # 預設取第一個選擇的工作表
            sample_df = pd.read_excel(first_file, sheet_name=first_sheet, header=0)
            columns = list(sample_df.columns)

            # 用戶選擇需要用於橫向合併的欄位
            selected_column = st.selectbox(
                "請選擇一個欄位作為主要合併依據（例如：會計科目）：",
                options=columns,
            )
        except Exception as e:
            st.error(f"無法讀取檔案的欄位，請檢查檔案格式。錯誤：{e}")

    # 合併資料
    if st.button("執行合併"):
        merged_df = None

        try:
            for uploaded_file in uploaded_files:
                # 獲取用戶選定的多個工作表名稱
                sheets_to_merge = selected_sheets[uploaded_file.name]

                for selected_sheet in sheets_to_merge:
                    try:
                        if merge_mode == "縱向合併 (上下堆疊)":
                            # 獲取用戶選定的表頭行
                            df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=header_row - 1)

                            # 清理欄位名稱
                            df.columns = df.columns.str.strip()

                            # 添加來源檔案與工作表資訊
                            df["來源檔案"] = uploaded_file.name
                            df["來源工作表"] = selected_sheet

                            # 合併資料
                            if merged_df is None:
                                merged_df = df
                            else:
                                merged_df = pd.concat([merged_df, df], ignore_index=True)

                        elif merge_mode == "橫向合併 (左右拼接)":
                            # 橫向合併：根據用戶選定的主欄位
                            df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=0)

                            # 清理主欄位
                            df.columns = df.columns.str.strip()
                            if selected_column in df.columns:
                                df[selected_column] = df[selected_column].astype(str).str.strip().fillna("N/A")

                                # 設置索引，並確保索引名稱為 "會計科目"
                                df.set_index(selected_column, inplace=True)
                                #df.index.name = "會計科目"
                                df.index.name = selected_column  # 動態設置索引名稱為用戶選擇的合併依據
                            else:
                                st.warning(f"檔案 {uploaded_file.name} 的工作表 {selected_sheet} 缺少主欄位 {selected_column}，跳過該工作表。")
                                continue

                            # 添加來源檔案與工作表資訊
                            df = df.add_suffix(f"_{uploaded_file.name}_{selected_sheet}")

                            # 合併資料
                            if merged_df is None:
                                merged_df = df
                            else:
                                merged_df = pd.concat([merged_df, df], axis=1, join="outer")

                    except Exception as e:
                        st.warning(f"檔案 {uploaded_file.name} 的工作表 {selected_sheet} 合併失敗，原因：{e}")
                        continue

            # 縱向合併完成後，重置索引
            if merge_mode == "縱向合併 (上下堆疊)" and merged_df is not None:
                merged_df.reset_index(drop=True, inplace=True)

            # 填補空值
            if merged_df is not None:
                for column in merged_df.columns:
                    if merged_df[column].dtype in ["float64", "int64"]:
                        # 數字型欄位填補空值為 0
                        merged_df[column] = merged_df[column].fillna(0)
                    else:
                        # 其他型別欄位填補空值為 "N/A"
                        merged_df[column] = merged_df[column].fillna("N/A")

                # **修正：確保索引重置為欄位（橫向合併時適用）**
                if merge_mode == "橫向合併 (左右拼接)":
                    merged_df.reset_index(inplace=True)

                # 顯示合併完成的結果
                st.success("合併完成！")
                st.write(merged_df)

                # 提供下載選項
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                    merged_df.to_excel(writer, index=False, sheet_name="合併結果")
                output.seek(0)

                st.download_button(
                    label="下載合併結果",
                    data=output,
                    file_name="合併結果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("未生成任何合併結果，請檢查檔案與工作表格式是否正確。")

        except Exception as e:
            st.error(f"合併過程中發生錯誤：{e}")

else:
    st.info("請上傳至少一個 Excel 檔案以開始。")

