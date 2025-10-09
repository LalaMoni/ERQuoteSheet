import streamlit as st
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.drawing.image import Image as XLImage
from PIL import Image
from io import BytesIO
import datetime
import pandas as pd
from openpyxl.utils import get_column_letter


# ---------------- 函数 ----------------
def write_cell_safe(ws, row, col, value):
    cell = ws.cell(row=row, column=col)
    if not isinstance(cell, MergedCell):
        cell.value = value
    else:
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                top_left = merged_range.min_row, merged_range.min_col
                ws.cell(row=top_left[0], column=top_left[1], value=value)
                break

def calculate_prices(P, product_Q, total_Q, F, R):
    A_CNY = P + F / total_Q
    B_CNY = P / 1.13 * 1.02 + F / total_Q
    A_USD = A_CNY / R
    B_USD = B_CNY / R
    return round(B_CNY, 4), round(A_CNY, 4), round(B_USD, 4), round(A_USD, 4)

def excel_cell_size_to_pixels(ws, row, col):
    col_letter = get_column_letter(col)
    col_width = ws.column_dimensions[col_letter].width or 8.43
    row_height = ws.row_dimensions[row].height or 15
    px_width = int(col_width * 7)
    px_height = int(row_height * 1.33)
    return px_width, px_height

def insert_image(ws, img_file, cell, max_width, max_height, scale=0.7):
    img_pil = Image.open(BytesIO(img_file.getbuffer()))
    img_pil.thumbnail((max_width * scale, max_height * scale))
    img_bytes = BytesIO()
    img_pil.save(img_bytes, format="PNG")
    img_bytes.seek(0)
    img = XLImage(img_bytes)
    ws.add_image(img, cell)

# ---------------- Streamlit 页面 ----------------
st.title("报价单生成器")

st.header("上传 Excel 模板")
uploaded_template = st.file_uploader("请选择 Excel 模板文件", type=["xlsx"])

# 基本信息
st.header("基本信息")
purchaser = st.text_area("采购商信息")
order_no = st.text_input("编号（ERKJXXXXXXXXXX）")
date_input = st.text_input("日期（YYYY/MM/DD）", value=datetime.date.today().strftime("%Y/%m/%d"))
F_input = st.text_input("总费用（支持公式，例如 200+50*4）")
R_input = st.text_input("汇率（买入价）")
start_row = 13  # 默认13行

# --- 产品信息 ---
st.header("产品信息")
product_options = {
    "吸气片": ["SG-01", "SG-02", "SG-03"],
    "焊料": ["CB-01", "CB-02"],
}

# 动态添加产品行
if "product_rows" not in st.session_state:
    st.session_state.product_rows = 1
if st.button("添加产品"):
    st.session_state.product_rows += 1

products = []
for i in range(st.session_state.product_rows):
    st.subheader(f"产品 {i+1}")
    no = i + 1
    st.text(f"序号: {no}")
    name = st.selectbox(f"产品名称", list(product_options.keys()), key=f"name{i}")
    model = st.selectbox(f"型号", product_options[name], key=f"model{i}")
    P = st.number_input(f"净单价", format="%.4f", key=f"P{i}")
    Q = st.number_input(f"数量", min_value=0, key=f"Q{i}")
    uploaded_file = st.file_uploader(f"上传图片", type=["png","jpg","jpeg"], key=f"img{i}")
    products.append({"no": no, "name": name, "model": model, "P": P, "Q": Q, "img": uploaded_file})


# ---------------- 预览 ----------------
if st.button("预览报价单"):
    try:
        F = eval(F_input)
    except Exception as e:
        st.error(f"总费用输入错误: {e}")
        st.stop()
    try:
        R = float(R_input)
    except:
        st.error("汇率输入错误")
        st.stop()

    total_Q = sum(p["Q"] for p in products)
    preview_data = []
    for p in products:
        B_CNY, A_CNY, B_USD, A_USD = calculate_prices(p["P"], p["Q"], total_Q, F, R)
        preview_data.append({
            "序号": p["no"],
            "产品": p["name"],
            "型号": p["model"],
            "数量": p["Q"],
            "人民币单价(不含税)": B_CNY,
            "人民币单价(含税)": A_CNY,
            "美元单价(不含税)": B_USD,
            "美元单价(含税)": A_USD,
        })
    df_preview = pd.DataFrame(preview_data)
    st.table(df_preview)

# 生成报价单
if st.button("生成报价单"):
    try:
        F = eval(F_input)
        R = float(R_input)
    except Exception as e:
        st.error(f"F 或 R 输入错误: {e}")
        st.stop()

    total_Q = sum(p["Q"] for p in products)
    wb = load_workbook(uploaded_template)
    ws = wb.active
    
    # 写入基本信息
    write_cell_safe(ws, 4, 7, purchaser)
    ws.cell(row=8, column=7, value=order_no)
    ws.cell(row=9, column=7, value=date_input)

    # 列位置
    NO_COL = 1
    PRODUCT_COL = 2
    IMG_COL = 3
    MODEL_COL = 4
    QUANTITY_COL = 5
    RMB_COL_START = 7
    USD_COL_START = 9

    # 写入产品数据
    for idx, p in enumerate(products):
        row = start_row + idx
        ws.row_dimensions[row].height = 69
        max_w, max_h = excel_cell_size_to_pixels(ws, row, IMG_COL)

        ws.cell(row=row, column=NO_COL, value=p["no"])
        ws.cell(row=row, column=PRODUCT_COL, value=p["name"])
        ws.cell(row=row, column=MODEL_COL, value=p["model"])
        ws.cell(row=row, column=QUANTITY_COL, value=p["Q"])

        B_CNY, A_CNY, B_USD, A_USD = calculate_prices(p["P"], p["Q"], total_Q, F, R)
        write_cell_safe(ws, row, RMB_COL_START, B_CNY)
        write_cell_safe(ws, row, RMB_COL_START + 1, A_CNY)
        write_cell_safe(ws, row, USD_COL_START, B_USD)
        write_cell_safe(ws, row, USD_COL_START + 1, A_USD)

        # 插入图片
        if p["img"] is not None:
            insert_image(ws, p["img"], f"C{row}", max_width=max_w, max_height=max_h)

    # 保存新文件到 BytesIO 提供下载
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    new_file_name = f"报价单_{date_input.replace('/', '-')}.xlsx"

    st.download_button(
        "下载生成的报价单",
        data=output,
        file_name=new_file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
