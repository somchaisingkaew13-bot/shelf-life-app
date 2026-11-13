import streamlit as st
import pandas as pd
import numpy as np
import io

# ตั้งค่าหน้าเว็บ
st.set_page_config(page_title="Shelf Life Checker", layout="wide")
st.title("📦 โปรแกรมคำนวณอายุสินค้าคงเหลือ")
st.markdown("อัปโหลดไฟล์ Excel ทั้ง 3 ไฟล์เพื่อประมวลผลหาอายุสินค้าคงเหลือ และดาวน์โหลดรายงานที่ไฮไลต์สีแดงอัตโนมัติ")

st.header("1. อัปโหลดไฟล์ข้อมูล")
col1, col2, col3 = st.columns(3)
with col1:
    file_limit = st.file_uploader("📂 1. ไฟล์เกณฑ์อายุลูกค้ารับได้", type=['xlsx', 'xls'])
with col2:
    file_item = st.file_uploader("📂 2. ไฟล์รายการสินค้า (Item)", type=['xlsx', 'xls'])
with col3:
    file_inventory = st.file_uploader("📂 3. ไฟล์สินค้าคงเหลือ (Daily Cycle Count)", type=['xlsx', 'xls'])

if st.button("🚀 ประมวลผลสร้างรายงาน", use_container_width=True):
    if not (file_limit and file_item and file_inventory):
        st.warning("⚠️ กรุณาอัปโหลดไฟล์ให้ครบทั้ง 3 ไฟล์ก่อนกดประมวลผลครับ")
    else:
        try:
            with st.spinner('กำลังประมวลผลข้อมูล...'):
                # 1. อ่านไฟล์เกณฑ์
                df_limit = pd.read_excel(file_limit)
                df_limit = df_limit.rename(columns={'Description:': 'CustomerGroup', 'Shelf Life': 'Acceptable_Limit'})

                # 2. อ่านไฟล์ Item
                df_item = pd.read_excel(file_item)
                df_item = df_item.drop(0).copy()
                df_item = df_item[['SKU', 'SKUGROUP2']].copy()
                df_item['SKU'] = df_item['SKU'].astype(str).str.strip()
                df_item = df_item.rename(columns={'SKUGROUP2': 'CustomerGroup'})

                # 3. อ่านไฟล์ Inventory
                df_inv = pd.read_excel(file_inventory, header=2)
                df_inv = df_inv.dropna(subset=['Item'])
                item_str = df_inv['Item'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

                cols_to_drop = ['EXPIREDATE', 'Qty Allocated', 'Qty Picked', 'Lot', 'หมายเหตุ', 'Vendor', 'Checked', 'Unnamed: 18']
                for col in cols_to_drop:
                    if col in df_inv.columns:
                        df_inv = df_inv.drop(columns=[col])

                # 4. ประกอบร่าง (Merge)
                df_merged = df_inv.merge(df_item, left_on=item_str, right_on='SKU', how='left')
                df_merged = df_merged.merge(df_limit, on='CustomerGroup', how='left')
                
                if 'key_0' in df_merged.columns: df_merged = df_merged.drop(columns=['key_0'])
                if 'SKU' in df_merged.columns: df_merged = df_merged.drop(columns=['SKU'])

                # 5. คำนวณวันคงเหลือ
                df_merged['MfgDate'] = pd.to_datetime(df_merged['MfgDate'], errors='coerce')
                df_merged['ExpDate'] = pd.to_datetime(df_merged['ExpDate'], errors='coerce')
                df_merged['Total_Life_Days'] = (df_merged['ExpDate'] - df_merged['MfgDate']).dt.days
                df_merged['Remaining_Days'] = df_merged['Total_Life_Days'] - pd.to_numeric(df_merged['Stock Day'], errors='coerce')
                df_merged['Remaining_Percent'] = df_merged['Remaining_Days'] / df_merged['Total_Life_Days']
                df_merged['Is_Low_Shelf_Life'] = df_merged['Remaining_Percent'] < df_merged['Acceptable_Limit']

                # 6. สร้างไฟล์ Excel บนหน่วยความจำเว็บ (BytesIO)
                output = io.BytesIO()
                writer = pd.ExcelWriter(output, engine='xlsxwriter')
                df_merged.to_excel(writer, index=False, sheet_name='Shelf_Life_Report')

                workbook = writer.book
                worksheet = writer.sheets['Shelf_Life_Report']
                header_format = workbook.add_format({'bold': True, 'bg_color': '#1F497D', 'font_color': 'white', 'border': 1})
                red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
                percent_format = workbook.add_format({'num_format': '0.00%'})
                date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})

                for col_num, value in enumerate(df_merged.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                    worksheet.set_column(col_num, col_num, 14)

                if 'Description' in df_merged.columns:
                    idx_desc = df_merged.columns.get_loc('Description')
                    worksheet.set_column(idx_desc, idx_desc, 45)
                if 'MfgDate' in df_merged.columns:
                    idx_mfg = df_merged.columns.get_loc('MfgDate')
                    worksheet.set_column(idx_mfg, idx_mfg, 18, date_format)
                if 'ExpDate' in df_merged.columns:
                    idx_exp = df_merged.columns.get_loc('ExpDate')
                    worksheet.set_column(idx_exp, idx_exp, 18, date_format)

                if 'Acceptable_Limit' in df_merged.columns:
                    idx_limit = df_merged.columns.get_loc('Acceptable_Limit')
                    idx_remain = df_merged.columns.get_loc('Remaining_Percent')
                    idx_is_low = df_merged.columns.get_loc('Is_Low_Shelf_Life')
                    worksheet.set_column(idx_limit, idx_limit, 18, percent_format)
                    worksheet.set_column(idx_remain, idx_remain, 18, percent_format)
                    worksheet.set_column(idx_is_low, idx_is_low, 18)

                    def colnum_string(n):
                        string = ""
                        while n > 0:
                            n, remainder = divmod(n - 1, 26)
                            string = chr(65 + remainder) + string
                        return string

                    max_col_str = colnum_string(len(df_merged.columns))
                    max_row = len(df_merged) + 1
                    cell_range = f'A2:{max_col_str}{max_row}'
                    is_low_col_str = colnum_string(idx_is_low + 1)
                    formula = f'=${is_low_col_str}2=TRUE'
                    worksheet.conditional_format(cell_range, {'type': 'formula', 'criteria': formula, 'format': red_format})

                worksheet.freeze_panes(1, 0)
                worksheet.autofilter(f'A1:{max_col_str}{max_row}')
                writer.close()
                output.seek(0) # รีเซ็ตไฟล์เพื่อให้โหลดได้
                
            st.success("✅ สร้างรายงานสำเร็จแล้ว! สามารถดาวน์โหลดได้ที่ปุ่มด้านล่าง")
            
            # ปุ่มสำหรับให้ผู้ใช้งานคลิกเพื่อดาวน์โหลด
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์รายงาน (Excel)",
                data=output,
                file_name="Result_ShelfLife_Check.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาดในการประมวลผล: {e}")
