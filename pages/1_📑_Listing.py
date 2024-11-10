import os
import re
import streamlit as st

st.set_page_config(
    page_title="Listing",
    page_icon="📑"
)

def extract_ids_from_files(directory, project_ke):
    journal_ids = {}

    # List all subdirectories in the main directory
    for journal in os.listdir(directory):
        journal_path = os.path.join(directory, journal)
        
        if os.path.isdir(journal_path):
            # Prepare to collect IDs
            journal_ids[journal] = []

            loa_path = os.path.join(journal_path, 'LoA', 'Finish')
            if os.path.isdir(loa_path):
                for month_folder in os.listdir(loa_path):
                    month_path = os.path.join(loa_path, month_folder)
                    
                    # Check if folder name contains '[bulan]-op'
                    if f'{project_ke}' in month_folder:
                        for file_name in os.listdir(month_path):
                            if file_name.endswith('.pdf'):
                                # Extract ID from file name using regex
                                match = re.search(r'_(\d+)', file_name)
                                if match:
                                    journal_ids[journal].append('ID' + match.group(1))

    return journal_ids

def display_ids(journal_ids, bulan):
    tot = sum(len(ids) for ids in journal_ids.values())

    # Create a string to hold the output for easy copying
    output_text = f"Layout 1 bulan terakhir ({bulan})\n"
    output_text += f"Total : {tot} Paper\n\n"

    for journal, ids in journal_ids.items():
        output_text += f"{journal}\n"
        for index, id in enumerate(ids):
            output_text += f"{index + 1}. {id}\n"
        output_text += "\n"

    # Display all results in a text area for easy copying
    st.text_area("Copyable Output", output_text, height=300)

if __name__ == "__main__":
    st.title("Journal ID Extractor")
    
    col1, col2 = st.columns(2)

    with col1:
        project_ke = st.text_input("Project ke:", "38")

    with col2:
        bulan = st.text_input("Bulan:", "Oktober")
    
    # Directory input for Streamlit
    main_directory = "D:\\PROJECT"
    
    # Extract and display IDs if directory path is provided
    if main_directory:
        journal_ids = extract_ids_from_files(main_directory, project_ke)
        display_ids(journal_ids, bulan)
