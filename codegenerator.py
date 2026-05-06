import streamlit as st
import qrcode
import barcode
from barcode.writer import ImageWriter
from io import BytesIO

st.title("QR Code & Barcode Generator")

text = st.text_input("Enter Text", "ER46")
code_type = st.selectbox("Code Type", ["QR Code", "Code128", "Code39", "EAN13", "EAN8", "UPCA"])

if st.button("Generate"):
    buffer = BytesIO()

    if code_type == "QR Code":
        img = qrcode.make(text)
        img.save(buffer, format="PNG")

    else:
        type_map = {
            "Code128": "code128",
            "Code39": "code39",
            "EAN13": "ean13",
            "EAN8": "ean8",
            "UPCA": "upca",
        }
        try:
            bc_class = barcode.get_barcode_class(type_map[code_type])
            bc = bc_class(text, writer=ImageWriter())
            bc.write(buffer)
        except Exception as e:
            st.error(f"Error: {e}")
            st.stop()

    buffer.seek(0)
    st.image(buffer, caption=f"{code_type}: {text}")
    st.download_button(
        "Download",
        data=buffer.getvalue(),
        file_name=f"{text}_{code_type}.png",
        mime="image/png"
    )