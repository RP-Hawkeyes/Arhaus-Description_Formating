#!/usr/bin/env python
# coding: utf-8

import streamlit as st
import pandas as pd
import openpyxl  # Ensure openpyxl is imported
import os
import re

# Debugging step to print installed packages
import subprocess
subprocess.run(["pip", "list"], stdout=subprocess.PIPE, text=True, check=True)

def format_descriptions(df):
    # Format the descriptions
    def format_description(description):
        # Split by periods followed by no space and strip each part
        parts = [part.strip() for part in re.split(r'\.(?=\S)', description)]
        # Filter out any empty parts
        parts = [part for part in parts if part]
        # Join with bullet points
        return '• ' + '\n• '.join(parts)
    
    df['Description'] = df['Description'].apply(format_description)
    return df

def main():
    st.title("Excel Description Formatter")

    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            st.write("File uploaded successfully. Here are the first few rows:")
            st.write(df.head())

            if 'Description' in df.columns:
                output_file_name = st.text_input("Enter the output file name (with .xlsx extension)", value="formatted_output.xlsx")

                if st.button("Format Descriptions"):
                    if output_file_name:
                        formatted_df = format_descriptions(df)
                        output_file_path = os.path.join("downloads", output_file_name)
                        os.makedirs("downloads", exist_ok=True)
                        formatted_df.to_excel(output_file_path, index=False)
                        st.success(f"Formatted data saved to {output_file_path}")

                        # Provide a download link
                        with open(output_file_path, "rb") as file:
                            st.download_button(label="Download formatted file", data=file, file_name=output_file_name)
                    else:
                        st.error("Please enter a valid output file name")
            else:
                st.error("The uploaded file does not contain a 'Description' column.")
        except Exception as e:
            st.error(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
