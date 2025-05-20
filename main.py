from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
import xlsxwriter
import os
from uuid import uuid4

app = FastAPI()

@app.post("/generate-excel/")
async def generate_excel(file: UploadFile = File(...)):
    categories = ['Jan', 'Feb', 'Mar']
    series = [10, 20, 30]

    filename = f"{uuid4().hex}.xlsx"
    filepath = f"/tmp/{filename}"

    workbook = xlsxwriter.Workbook(filepath)
    worksheet = workbook.add_worksheet()

    worksheet.write_column('A2', categories)
    worksheet.write_column('B2', series)
    worksheet.write('A1', 'Month')
    worksheet.write('B1', 'Value')

    chart = workbook.add_chart({'type': 'column'})
    chart.add_series({
        'name': 'Example',
        'categories': f'=Sheet1!$A$2:$A$4',
        'values':     f'=Sheet1!$B$2:$B$4',
    })
    chart.set_title({'name': 'Example Chart'})
    worksheet.insert_chart('D2', chart)

    workbook.close()
    return FileResponse(filepath, filename="chart.xlsx")
