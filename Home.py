import streamlit as st
from docx import Document
import pythoncom
import win32com.client as win32
from pathlib import Path
import os
from docx2pdf import convert

st.set_page_config(
    page_title="LoA Generator",
    page_icon="✨"
)

st.title("LoA Generator 😁")


st.markdown(
    """
    Selamat datang di aplikasi LoA Generator. Disini kamu bisa membuat LoA jurnal secara otomatis lho. Berikut adalah cara yang dapat kamu lakukan:
    1. Upload jurnal berekstensi .docx yang mau kamu buat LoA-nya   
    2. Jika sudah sesuai maka hasilnya akan secara otomatis ditampilkan
    3. Anda akan memperoleh file LoA word dan PDF sesuai dengan path tiap jurnal
    
    """
)

st.title("Letter of Acceptance")
st.write("Kamu bisa copy pesan ini, lalu paste ke pengiriman LoA.")

tab1, tab2, tab3, tab4,tab5, tab6, tab7 = st.tabs(["JAMSI", "IJPM", "JIPPM", "JPMII", "JUPIN", "JIKI", "KONTAK"])

with tab1:
    st.code(
    """
    Dear Author,
    Berikut ini kami lampirkan Letter of Acceptance dari paper Anda yang di submit ke Jurnal Abdi Masyarakat Indonesia (JAMSI).
    Saat ini paper Anda sedang dalam proses layout dan proofread untuk dipublikasikan.

    Terima kasih atas kontribusi yang Anda berikan di Jurnal Abdi Masyarakat Indonesia.
    Kami menunggu paper Anda selanjutnya di Jurnal Abdi Masyarakat Indonesia.
    
    """
    )
with tab2:
    st.code(
    """
    Dear Author,
    Berikut ini kami lampirkan Letter of Acceptance dari paper Anda yang di submit ke Inovasi Jurnal Pengabdian Masyarakat (IJPM).
    Saat ini paper Anda sedang dalam proses layout dan proofread untuk dipublikasikan.

    Terima kasih atas kontribusi yang Anda berikan di Inovasi Jurnal Pengabdian Masyarakat.
    Kami menunggu paper Anda selanjutnya di Inovasi Jurnal Pengabdian Masyarakat.
    
    """
    )
with tab3:
    st.code(
    """
    Dear Author,
    Berikut ini kami lampirkan Letter of Acceptance dari paper Anda yang di submit ke Jurnal Inovasi Pengabdian dan Pemberdayaan Masyarakat (JIPPM).
    Saat ini paper Anda sedang dalam proses layout dan proofread untuk dipublikasikan.

    Terima kasih atas kontribusi yang Anda berikan di Jurnal Inovasi Pengabdian dan Pemberdayaan Masyarakat.
    Kami menunggu paper Anda selanjutnya di Jurnal Inovasi Pengabdian dan Pemberdayaan Masyarakat.
        
    """
    )
with tab4:
    st.code(
    """
    Dear Author,
    Berikut ini kami lampirkan Letter of Acceptance dari paper Anda yang di submit ke Jurnal Pengabdian Masyarakat Inovatif Indonesia (JPMII).
    Saat ini paper Anda sedang dalam proses layout dan proofread untuk dipublikasikan.

    Terima kasih atas kontribusi yang Anda berikan di Jurnal Pengabdian Masyarakat Inovatif Indonesia.
    Kami menunggu paper Anda selanjutnya di Jurnal Pengabdian Masyarakat Inovatif Indonesia.
    
    """
    )
with tab5:
    st.code(
    """
    Dear Author,
    Berikut ini kami lampirkan Letter of Acceptance dari paper Anda yang di submit ke Jurnal Penelitian Inovatif (JUPIN).
    Saat ini paper Anda sedang dalam proses layout dan proofread untuk dipublikasikan. 

    Terima kasih atas kontribusi yang Anda berikan di Jurnal Penelitian Inovatif.
    Kami menunggu paper Anda selanjutnya di Jurnal Penelitian Inovatif.
    
    """
    )
with tab6:
    st.code(
    """
    Dear Author,
    Berikut ini kami lampirkan Letter of Acceptance dari paper Anda yang di submit ke Jurnal Ilmu Komputer dan Informatika (JIKI).
    Saat ini paper Anda sedang dalam proses layout dan proofread untuk dipublikasikan.

    Terima kasih atas kontribusi yang Anda berikan di Jurnal Ilmu Komputer dan Informatika.
    Kami menunggu paper Anda selanjutnya di Jurnal Ilmu Komputer dan Informatika.
    
    """
    )
with tab7:
    st.code(
    """
    Dear Author,
    Berikut ini kami lampirkan Letter of Acceptance dari paper Anda yang di submit ke Jurnal Komputer dan Teknik Informatika (KONTAK).
    Saat ini paper Anda sedang dalam proses layout dan proofread untuk dipublikasikan.

    Terima kasih atas kontribusi yang Anda berikan di Jurnal Komputer dan Teknik Informatika.
    Kami menunggu paper Anda selanjutnya di Jurnal Komputer dan Teknik Informatika.
    
    """
    )

# Function to convert DOCX to PDF
def convert_docx_to_pdf(docx_path, pdf_path):
    pythoncom.CoInitialize()  # Ensure COM initialization
    try:
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(docx_path)
        doc.SaveAs(pdf_path, FileFormat=17)  # 17 is the format code for PDF
        doc.Close()
        word.Quit()
    except Exception as e:
        return f"Error in PDF conversion: {e}"
    finally:
        pythoncom.CoUninitialize()
    print("PDF saved successfully.")

# Streamlit app
st.title("Word to PDF Converter")

# Function to convert DOCX to PDF
def convert_docx_to_pdf(docx_path, pdf_output_path):
    pythoncom.CoInitialize()  # Ensure COM initialization
    try:
        convert(docx_path, pdf_output_path)
    except Exception as e:
        return f"Error in PDF conversion: {e}"
    finally:
        pythoncom.CoUninitialize()
    return pdf_output_path  # Return the path of the generated PDF

# Upload Word document
uploaded_file = st.file_uploader("Choose a DOCX file", type="docx")

if uploaded_file is not None:
    # Save the uploaded DOCX file temporarily
    docx_path = "uploaded_file.docx"  # Correctly assign the file name
    with open(docx_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    # Define the PDF output path
    file_name = uploaded_file.name.replace(".docx", "")
    downloads_dir = Path.home() / "Downloads"
    pdf_output_path = downloads_dir / f"{file_name}.pdf"  # Correctly format the PDF filename
    
    # Convert DOCX to PDF
    conversion_result = convert_docx_to_pdf(docx_path, pdf_output_path)
    
    # Check if PDF was created successfully
    if os.path.exists(pdf_output_path):
        # Allow user to download the converted PDF
        # with open(pdf_output_path, "rb") as pdf_file:
        #     st.download_button(
        #         label="Download PDF",
        #         data=pdf_file,
        #         file_name=os.path.basename(pdf_output_path),
        #         mime="application/pdf"
        #     )
        st.success("File convert ada di berkas Download ya 😃")
        
        # Clean up temporary files
        os.remove(docx_path)
        # os.remove(pdf_output_path)
    else:
        st.error(f"Conversion: {conversion_result}")  # Use the error message from the function


