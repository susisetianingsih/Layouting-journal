# import fitz  # PyMuPDF
# import re
# import json
# import streamlit as st
# import io

# # Fungsi untuk ekstrak informasi jurnal dari halaman pertama PDF
# def extract_journal_info(pdf_file):
#     journal_info = {
#         "title": "",
#         "author": "",
#         "affiliation": ""
#     }

#     # Buka PDF dari objek bytes
#     pdf = fitz.open(stream=pdf_file.read(), filetype="pdf")
#     page = pdf[0]  # Halaman pertama
#     text = page.get_text()

#     # Ekstraksi judul
#     title_match = re.search(r"\d+\n(.+?)\n\n", text, re.DOTALL)
#     if title_match:
#         title_text = title_match.group(1).strip()
#         title_text = re.sub(r"^\d+\s*", "", title_text)
#         journal_info["title"] = title_text

#     # Ekstraksi penulis
#     author_match = re.search(r"(\n.+?(\*|\d))\n", text, re.DOTALL)
#     if author_match:
#         author_text = author_match.group(1).strip()
#         author_text = re.sub(r"^\d+\s*", "", author_text)
#         journal_info["author"] = author_text

#     # Ekstraksi afiliasi
#     affiliation_match = re.search(r"(\d\s*(.+?))\nEmail", text, re.DOTALL)
#     if affiliation_match:
#         affiliation_text = affiliation_match.group(2).strip()
#         journal_info["affiliation"] = affiliation_text

#     return journal_info

# # Membuat antarmuka Streamlit
# st.title("Journal Info Extractor")

# # Upload file PDF
# uploaded_file = st.file_uploader("Upload PDF file", type="pdf")

# # Tombol untuk melakukan ekstraksi
# if uploaded_file is not None:
#     if st.button("Extract Journal Information"):
#         # Ekstrak informasi dari PDF yang diunggah
#         journal_info = extract_journal_info(uploaded_file)
        
#         # Tampilkan hasil ekstraksi
#         st.subheader("Extracted Journal Information:")
#         st.write(f"**Title:** {journal_info['title']}")
#         st.write(f"**Author:** {journal_info['author']}")
#         st.write(f"**Affiliation:** {journal_info['affiliation']}")

#         # Simpan hasil ke file JSON (opsional)
#         if st.button("Save to JSON"):
#             with open("journal_info.json", "w") as json_file:
#                 json.dump(journal_info, json_file, indent=4)
#             st.success("Journal information saved to journal_info.json")

# import fitz  # PyMuPDF
# import re
# import json
# import streamlit as st

# # Fungsi untuk ekstrak informasi jurnal dari halaman pertama PDF
# def extract_journal_info(pdf_file):
#     journal_info = {
#         "title": "",
#         "author": "",
#         "affiliation": ""
#     }

#     try:
#         # Buka PDF dari objek bytes
#         pdf = fitz.open(stream=pdf_file.read(), filetype="pdf")
#         page = pdf[0]  # Halaman pertama
#         text = page.get_text()

#         return text, journal_info

#     except Exception as e:
#         st.error(f"An error occurred: {str(e)}")
#         return None, journal_info

# # Membuat antarmuka Streamlit
# st.title("Journal Info Extractor")

# # Upload file PDF
# uploaded_file = st.file_uploader("Upload PDF file", type="pdf")

# # Tombol untuk melakukan ekstraksi
# if uploaded_file is not None:
#     if st.button("Extract Journal Information"):
#         # Ekstrak informasi dari PDF yang diunggah
#         extracted_text, journal_info = extract_journal_info(uploaded_file)

#         # Tampilkan teks yang diekstrak secara keseluruhan
#         st.subheader("Extracted Text:")
#         st.text(extracted_text)  # Menampilkan seluruh teks yang diambil

#         # Ekstraksi informasi dari teks yang diambil
#         if extracted_text:
#             # Ekstraksi judul
#             title_match = re.search(r"(e-ISSN:\s*\d{4}-\d{4}\s*\n+\d+\s*\n)([^\n].+?)(?=\n)", extracted_text)
#             if title_match:
#                 journal_info["title"] = title_match.group(0).strip()
#                 title = journal_info["title"]

#             # Ekstraksi penulis
#             author_match = re.search(r"({title}\n\n)(.+?)(?=\n\n)", extracted_text)
#             if author_match:
#                 journal_info["author"] = author_match.group(0).strip()

#             # Ekstraksi afiliasi
#             affiliation_match = re.search(r"(Universitas.+?)(?=,+\s+Indonesia)", extracted_text)
#             if affiliation_match:
#                 journal_info["affiliation"] = affiliation_match.group(0).strip()

#         # Tampilkan hasil ekstraksi informasi jurnal
#         st.subheader("Extracted Journal Information:")
#         st.write(f"**Title:** {journal_info['title']}")
#         st.write(f"**Author:** {journal_info['author']}")
#         st.write(f"**Affiliation:** {journal_info['affiliation']}")

#         # Simpan hasil ke file JSON (opsional)
#         if st.button("Save to JSON"):
#             with open("journal_info.json", "w") as json_file:
#                 json.dump(journal_info, json_file, indent=4)
#             st.success("Journal information saved to journal_info.json")

import streamlit as st
from docx import Document

def extract_info_from_docx(file):
    # Membaca dokumen Word
    doc = Document(file)
    title = ""
    authors = ""
    affiliation = ""

    # Asumsi: Judul ada di paragraf pertama, penulis di paragraf kedua, dan afiliasi di paragraf ketiga
    paragraphs = doc.paragraphs
    if len(paragraphs) > 0:
        title = paragraphs[0].text
    if len(paragraphs) > 1:
        authors = paragraphs[1].text
    if len(paragraphs) > 3:
        affiliation = paragraphs[3].text
    
    return title, authors, affiliation

# Membuat antarmuka Streamlit
st.title("Ekstraksi Data dari Jurnal")
st.write("Unggah file jurnal Word (DOCX) untuk mengekstrak judul, penulis, dan afiliasi.")

# Upload file
uploaded_file = st.file_uploader("Pilih file Word", type=["docx"])

if uploaded_file is not None:
    # Ekstrak informasi dari file yang diunggah
    title, authors, affiliation = extract_info_from_docx(uploaded_file)

    # Tampilkan hasil ekstraksi
    st.subheader("Hasil Ekstraksi:")
    st.write("**Judul:**", title)
    st.write("**Penulis:**", authors)
    st.write("**Afiliasi:**", affiliation)
