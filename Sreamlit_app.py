import xml.etree.ElementTree as ET
import xlwings as xw
import streamlit as st
import os
from datetime import datetime


def main():
    st.title("SEBI Data Processing")

    uploaded_file = st.file_uploader("Upload XML file", type=["xml"])
    if uploaded_file is not None:
        # Save the uploaded file temporarily
        with open("temp.xml", "wb") as temp_file:
            temp_file.write(uploaded_file.getvalue())

        tree = ET.parse("temp.xml")
        root = tree.getroot()

        # tree = ET.parse('D:\SEBI_DATA\sebi.xml')
        # root = tree.getroot()
        # Get the current month and financial year
        current_date = datetime.now()
        month = current_date.strftime("%B")  # Get the full month name
        financial_year = f"{current_date.year}-{current_date.year + 1}"

        path = "D:\\SEBI_DATA\\sebi_data.xlsx"
        wb_obj = xw.Book(path)
        wks = wb_obj.sheets.active

        wks.range('A3').value=f"Report for the month of {month} FY {financial_year}"

        #Whether the Service is Offered
        wks.range('N8').value=root[1][0].text
        wks.range('N9').value=root[1][1].text
        wks.range('N10').value=root[1][2].text

        #Data for Discretionary Services
        #PF/EPFO
        wks.range('B16').value=root[2][0][0][0].text
        wks.range('B17').value=root[2][0][1][0].text

        #Corporates
        wks.range('E16').value=root[2][0][0][1].text
        wks.range('E17').value=root[2][0][1][1].text

        #Non-Corporates
        wks.range('H16').value=root[2][0][0][2].text
        wks.range('H17').value=root[2][0][1][2].text

        #NRI
        wks.range('K16').value=root[2][0][0][3].text
        wks.range('K17').value=root[2][0][1][3].text

        #FPI
        wks.range('N16').value=root[2][0][0][4].text
        wks.range('N17').value=root[2][0][1][4].text

        #Others
        wks.range('Q16').value=root[2][0][0][5].text
        wks.range('Q17').value=root[2][0][1][5].text

        #Total
        wks.range('T16').value=root[2][0][0][6].text
        wks.range('T17').value=root[2][0][1][6].text

        #Break-up of assets under management of the Portfolio Manager 
        #Equity
        wks.range('B23').value=root[2][1][0][0][1][0].text
        wks.range('B24').value=root[2][1][0][1][0][0].text

        wks.range('D23').value=root[2][1][0][0][1][1].text
        wks.range('D24').value=root[2][1][0][1][0][1].text

        #Plain Debt
        wks.range('F23').value=root[2][1][0][0][2][0].text
        wks.range('F24').value=root[2][1][0][1][2][0].text

        wks.range('H23').value=root[2][1][0][0][2][1].text
        wks.range('H24').value=root[2][1][0][1][2][1].text

        #Structured Debt
        wks.range('J23').value=root[2][1][0][0][3][0].text
        wks.range('J24').value=root[2][1][0][1][3][0].text

        wks.range('L23').value=root[2][1][0][0][3][1].text
        wks.range('L24').value=root[2][1][0][1][3][1].text

        #Derivatives
        wks.range('N23').value=root[2][1][0][0][4][0].text
        wks.range('N24').value=root[2][1][0][1][3][0].text

        wks.range('O23').value=root[2][1][0][0][4][1].text
        wks.range('O24').value=root[2][1][0][1][3][1].text

        wks.range('Q23').value=root[2][1][0][0][4][2].text
        wks.range('Q24').value=root[2][1][0][1][3][2].text

        #Mutual Funds
        wks.range('R23').value=root[2][1][0][0][5].text
        wks.range('R24').value=root[2][1][0][1][4].text

        #Others
        wks.range('T23').value=root[2][1][0][0][6].text
        wks.range('T24').value=root[2][1][0][1][5].text

        #Total
        wks.range('U23').value=root[2][1][0][0][7].text
        wks.range('U24').value=root[2][1][0][1][6].text

        #Funds Inflow/ Outflow
        #Inflow during the month
        wks.range('B28').value=root[2][2][0][1].text
        wks.range('B29').value=root[2][2][1][0].text

        wks.range('F28').value=root[2][2][0][2].text
        wks.range('F29').value=root[2][2][1][1].text

        wks.range('I28').value=root[2][2][0][3].text
        wks.range('I29').value=root[2][2][1][2].text

        wks.range('M28').value=root[2][2][0][4].text
        wks.range('M29').value=root[2][2][1][3].text

        wks.range('P28').value=root[2][2][0][5].text
        wks.range('P29').value=root[2][2][1][4].text

        wks.range('T28').value=root[2][2][0][6].text
        wks.range('T29').value=root[2][2][1][5].text

        #Transaction Data
        wks.range('R33').value=root[2][3][0].text
        wks.range('R34').value=root[2][3][1].text
        wks.range('R35').value=root[2][3][2].text

        #Performance Data
        #AUM (in INR Cr.)
        wks.range('D40').value=root[2][4][0][0][1].text
        # wks.range('D41').value=

        #Return TWRR (%)/1 month
        wks.range('G40').value=root[2][4][0][0][2][0].text
        wks.range('G41').value=root[2][4][0][1][1][0].text

        #Return TWRR (%)/1 year
        wks.range('K40').value=root[2][4][0][0][2][1].text
        wks.range('K41').value=root[2][4][0][1][1][1].text

        #Portfolio/1 month
        wks.range('O40').value=root[2][4][0][0][3][0].text
        # wks.range('K41').value=

        #Portfolio/1 year
        wks.range('S40').value=root[2][4][0][0][3][1].text
        # wks.range('K41').value=

        # Save the workbook after applying formulas
        wb_obj.save(f"D:\\SEBI_DATA\\sebi_data2.xlsx")
        # Close the workbook
        wb_obj.close()

        # Remove the temporary file
        os.remove("temp.xml")

        # Provide a download button for the modified Excel file
        st.download_button(
            label="Download Modified Excel File",
            data=open("D:\\SEBI_DATA\\sebi_data2.xlsx", "rb").read(),
            file_name="sebi_data2.xlsx",
            key="download_button",
        )

if __name__ == "__main__":
    main()
