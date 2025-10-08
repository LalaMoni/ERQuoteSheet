import streamlit as st
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.drawing.image import Image as XLImage
from PIL import Image
from io import BytesIO

# ---------------- 函数 ----------------
def write_cell_safe(ws, row, col, value):
    """安全写入单元格，处理合并单元格"""
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
    """计算含税/不含税人民币与美元单价，返回 float 保留四位小数"""
    A_CNY = P + F / total_Q
    B_CNY = P / 1.13 * 1.02 + F / total_Q
    A_USD = A_CNY / R
    B_USD = B_CNY / R
    return round(B_CNY, 4), round(A_CNY, 4), round(B_USD, 4), round(A_USD, 4)

def excel_cell_size_to_pixels(ws, row, col):
    """根据 Excel 列宽和行高计算像素大小"""
    col_letter = ws.cell(row=row, column=col).column_letter
    col_width = ws.column_dimensions[col_letter].width or 8.43
    row_height = ws.row_dimensions[row].height or 15
    px_width = int(col_width * 7)
    px_height = int(row_height * 1.33)
    return px_width, px_height

def insert_image(ws, img_file, cell, max_width, max_height, scale=0.7):
    """在 Excel 中插入图片，自动缩放到单元格大小"""
    img_pil = Image.open(BytesIO(img_file.getbuffer()))
    img_pil.thumbnail((max_width * scale, max_height * scale))
    img_bytes = BytesIO()
    img_pil.save(img_bytes, format="PNG")
    img_bytes.seek(0)
    img = XLImage(img_bytes)
    ws.add_image(img, cell)

# ---------------- Streamlit 页面 ----------------
st.title("报价单生成器")

# 基本信息
st.header("基本信息")
purchaser = st.text_area("采购商信息")
order_no = st.text_input("编号")
date_input = st.text_input("日期 (YYYY/MM/DD)")
F_input = st.text_input("总费用 F（支持加减运算）")
R_input = st.text_input("汇率 R")
start_row = st.number_input("模板开始填数据行", min_value=1, value=13)

# 上传模板
st.header("上传 Excel 模板")
uploaded_template = st.file_uploader("请选择 Excel 模板文件", type=["xlsx"])

# 产品信息
st.header("产品信息")
product_count = st.number_input("产品数量", min_value=1, value=1)
products = []
for i in range(product_count):
    st.subheader(f"产品{i+1}")
    no = st.number_input(f"序号", value=i+1, key=f"no{i}")
    name = st.text_input(f"产品类别", key=f"name{i}")
    model = st.text_input(f"型号", key=f"model{i}")
    P = st.number_input(f"净单价 (人民币)", format="%.4f", key=f"P{i}")
    Q = st.number_input(f"数量", value=0, key=f"Q{i}")
    uploaded_file = st.file_uploader(f"上传图片", type=["png","jpg","jpeg"], key=f"img{i}")
    products.append({"no": no, "name": name, "model": model, "P": P, "Q": Q, "img": uploaded_file})

# 生成报价单
if st.button("生成报价单"):
    if uploaded_template is None:
        st.error("请先上传 Excel 模板")
        st.stop()

    # 解析 F 和 R
    try:
        F = eval(F_input)
    except Exception as e:
        st.error(f"F 输入错误: {e}")
        st.stop()
    try:
        R = float(R_input)
    except:
        st.error("汇率 R 输入错误")
        st.stop()

    total_Q = sum(p["Q"] for p in products)

    # 读取模板
    wb = load_workbook(uploaded_template)
    ws = wb.active

    # 写入采购商、编号、日期
    write_cell_safe(ws, 4, 7, purchaser)
    ws.cell(row=8, column=7, value=order_no)
    ws.cell(row=9, column=7, value=date_input)

    # 假设列位置
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
