import win32com.client
import os

try:
    file_name = "CZ-PT-A0000 Maliyetler.xlsm"
    script_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(script_dir, file_name)

    excel = win32com.client.Dispatch("Excel.Application")
    workbook = excel.Workbooks.Open(file_path)

    # A kolonundaki verileri almak
    data_range = ('','ias', 'mekanik', 'elektrik', 'mekanik maliyet', 'elektrik maliyet', 'toplam maliyet')

    print(data_range)
    for worksheet in workbook.Worksheets:
        if worksheet.Name.startswith(('CT', 'MS', 'CZM')):
            # Veri doğrulama kısmında liste seçip belirli bir aralıktaki verileri P1 hücresinde gösterme
            data_validation = worksheet.Range('P1').Validation
            data_validation.Delete()
            data_validation.Add(Type=3, AlertStyle=1, Operator=1, Formula1=";".join(data_range))
            data_validation.IgnoreBlank = True
            data_validation.InCellDropdown = True
            data_validation.InputMessage = "Listeden Seçiniz"  # Set the default prompt


    excel.DisplayAlerts = False  # Prevent Excel from showing alerts when saving

except Exception as e:
    print(f"Error: {str(e)}")
