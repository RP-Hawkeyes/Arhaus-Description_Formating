#!/usr/bin/env python
# coding: utf-8

import streamlit as st
import pandas as pd
import openpyxl
import os
import re

def format_descriptions(df, column_name):
    # Format the descriptions in the specified column
    def format_description(description):
        # Split by periods followed by no space and strip each part
        parts = [part.strip() for part in re.split(r'\.(?=\S)', description)]
        # Filter out any empty parts
        parts = [part for part in parts if part]
        # Join with bullet points
        return '• ' + '\n• '.join(parts)
    
    df[column_name] = df[column_name].apply(format_description)
    return df

def main():
    st.title("Description Formatter - (Para to Bulletin)")

    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        st.write("File uploaded successfully. Here are the first few rows:")
        st.write(df.head())
        
        # User input for column name
        column_name = st.text_input("Enter the column name containing descriptions")
        
        # User input for output file name (without extension)
        output_file_name = st.text_input("Enter the output file name (without extension)")
        
        if st.button("Format Descriptions"):
            if column_name in df.columns:
                if output_file_name:
                    formatted_df = format_descriptions(df, column_name)
                    output_file_path = os.path.join("downloads", f"{output_file_name}.xlsx")
                    os.makedirs("downloads", exist_ok=True)
                    formatted_df.to_excel(output_file_path, index=False)
                    st.write("Here is the formatted data:")
                    st.write(formatted_df)
                    st.success(f"Formatted data saved to {output_file_path}")

                    # Provide a download link
                    with open(output_file_path, "rb") as file:
                        st.download_button(label="Download formatted file", data=file, file_name=f"{output_file_name}.xlsx")
                else:
                    st.error("Please enter a valid output file name without extension.")
            else:
                st.error(f"The uploaded file does not contain a column named '{column_name}'.")
                
if __name__ == "__main__":
    main()
