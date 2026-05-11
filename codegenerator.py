import streamlit as 
import qrcode
import barcode
from barcode.writer import ImageWriter
from io import BytesIO
import pandas as pd
import zipfile
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage

st.title("QR Code & Barcode Generator")

tab1, tab2 = st.tabs(["Single", "Batch"])

def generate_code(text, code_type):
    buffer = BytesIO()
    if code_type == "QR Code":
        img = qrcode.make(str(text))
        img.save(buffer, format="PNG")
    else:
        type_map = {
            "Code128": "code128",
            "Code39": "code39",
            "EAN13": "ean13",
            "EAN8": "ean8",
            "UPCA": "upca",
        }
        bc_class = barcode.get_barcode_class(type_map[code_type])
        bc = bc_class(str(text), writer=ImageWriter())
        bc.write(buffer)
    buffer.seek(0)
    return buffer

with tab1:
    text = st.text_input("Enter Text", "ER46")
    code_type = st.selectbox("Code Type", ["QR Code", "Code128", "Code39", "EAN13", "EAN8", "UPCA"])

    if st.button("Generate", key="single_generate"):
        try:
            buf = generate_code(text, code_type)
            st.image(buf, caption=f"{code_type}: {text}")
            st.download_button(
                "Download",
                data=buf.getvalue(),
                file_name=f"{text}_{code_type}.png",
                mime="image/png"
            )
        except Exception as e:
            st.error(f"Error: {e}")

with tab2:
    with st.expander("📋 File Format Guide", expanded=True):
        st.markdown("""
上传任意 Excel 或 CSV 文件，将自动使用**第一列**的数据生成码。

**各码型字符要求：**
| 码型 | 要求 |
|------|------|
| QR Code | 任意字符串（含中文） |
| Code128 | 任意 ASCII 字符串（字母、数字、符号，不含中文） |
| Code39 | 大写字母、数字，以及 ` - . $ / + % 空格` |
| EAN13 | 恰好 **12** 位数字（末位校验码自动生成） |
| EAN8 | 恰好 **7** 位数字（末位校验码自动生成） |
| UPCA | 恰好 **11** 位数字（末位校验码自动生成） |
""")
        template_df = pd.DataFrame({"SKU": ["ER46", "ER47", "ER48"]})
        st.dataframe(template_df, use_container_width=True, hide_index=True)
        st.download_button(
            "⬇️ Download template CSV",
            data=template_df.to_csv(index=False).encode(),
            file_name="batch_upload_template.csv",
            mime="text/csv",
        )

    uploaded = st.file_uploader("Upload an Excel or CSV file", type=["xlsx", "xls", "csv"])

    if uploaded:
        try:
            df = pd.read_csv(uploaded) if uploaded.name.endswith(".csv") else pd.read_excel(uploaded)
        except Exception as e:
            st.error(f"Failed to read file: {e}")
            st.stop()

        st.dataframe(df.head(), use_container_width=True)

        text_col = df.columns[0]
        st.info(f"Using column: **{text_col}**")

        batch_type = st.selectbox("Code Type", ["QR Code", "Code128", "Code39", "EAN13", "EAN8", "UPCA"], key="batch")
        output_format = st.radio("Output Format", ["ZIP (individual PNGs)", "Excel (images in column B)"], horizontal=True)

        if st.button("Generate", key="batch_generate"):
            values = df[text_col].dropna().tolist()
            if not values:
                st.warning("No valid data in the selected column.")
                st.stop()

            errors = []
            progress = st.progress(0, text="Generating...")

            if output_format == "Excel (images in column B)":
                wb = Workbook()
                ws = wb.active
                ws.column_dimensions["A"].width = 20
                ws.column_dimensions["B"].width = 40
                ws["A1"] = text_col
                ws["B1"] = batch_type

                for i, val in enumerate(values):
                    row = i + 2
                    ws.row_dimensions[row].height = 80
                    ws[f"A{row}"] = str(val)
                    try:
                        img_buf = generate_code(val, batch_type)
                        xl_img = XLImage(img_buf)
                        ratio = 70 / xl_img.height
                        xl_img.height = 70
                        xl_img.width = int(xl_img.width * ratio)
                        ws.add_image(xl_img, f"B{row}")
                    except Exception as e:
                        errors.append(f"Row {i+1} ({val}): {e}")
                    progress.progress((i + 1) / len(values), text=f"{i+1} / {len(values)}")

                xl_buf = BytesIO()
                wb.save(xl_buf)
                xl_buf.seek(0)

                st.success(f"Done! {len(values) - len(errors)} generated, {len(errors)} failed.")
                if errors:
                    with st.expander("View errors"):
                        for err in errors:
                            st.text(err)

                st.download_button(
                    "📊 Download Excel",
                    data=xl_buf.getvalue(),
                    file_name=f"batch_{batch_type}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            else:
                zip_buf = BytesIO()
                with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    for i, val in enumerate(values):
                        try:
                            img_buf = generate_code(val, batch_type)
                            zf.writestr(f"{val}_{batch_type}.png", img_buf.getvalue())
                        except Exception as e:
                            errors.append(f"Row {i+1} ({val}): {e}")
                        progress.progress((i + 1) / len(values), text=f"{i+1} / {len(values)}")

                zip_buf.seek(0)
                st.success(f"Done! {len(values) - len(errors)} generated, {len(errors)} failed.")
                if errors:
                    with st.expander("View errors"):
                        for err in errors:
                            st.text(err)

                st.download_button(
                    "📦 Download ZIP",
                    data=zip_buf.getvalue(),
                    file_name=f"batch_{batch_type}.zip",
                    mime="application/zip"
                )