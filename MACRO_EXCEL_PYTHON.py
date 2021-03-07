""" 
pip --version
python -m pip install -U pip

pip install pywin32

python
"""
import os
import win32com.client as win32

# Ouvrir EXCEL
xl=win32.Dispatch('Excel.Application')
wb=xl.Workbooks.Open('MOD_A30.xlsb')
#%%
xl.Visible=True
# ----------------------------
# ws=wb.Worksheets('TCD_01')
wb.Sheets("TCD_01").Select()
#-----------------------------

# Create a new Excel Workbook  ( Format Binary )
wb=xl.Workbooks.Add()
# wb.SaveAs(os.path.join(os.getcwd(), 'text.xlsx'),FileFormat:=51)
wb.SaveAs(os.path.join(os.getcwd(), 'text.xlsb'),FileFormat:=50)

ws_1=wb.Worksheets('Feuil1')
"""
Write data to Excel
"""
# Cells (row,col)
ws_1.Cells(5,2).Value="Cell B5"
ws_1.Cells(5,3).Value="Cell C5"

# Range()
ws_1.Range('A1').Value='ghjgj'
ws_1.Range('A2').Value='azsas'

ws_1.Range("A1:E5").Select()
ws_1.Range("A1:E5").Copy()
# ws_1.Cells.Copy()
# ws_1.Selection.Copy()

# Delete Cells
# ws_1.Cells.ClearContents()

"""
Write Data to multiple Cells
"""
ws_1.Range("A1:E5").Value="Hel"
ws_1.Range(ws_1.Cells(1,1),ws_1.Cells(5,5)).Value="Hello"


"""
Read Data from Cells
"""
for i in range(1,6):
    print(ws_1.Range(ws_1.Cells(i,1),ws_1.Cells(i,5)).Value)

ws_1.Range("A1:E5").Select









