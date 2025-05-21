from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
import xlsxwriter
import os
from uuid import uuid4
from openai import OpenAI
client = OpenAI()
import base64
import json

app = FastAPI()

# Load OpenAI API key from environment variable (in Railway)
openai.api_key = os.getenv("OPENAI_API_KEY")

@app.post("/generate-excel/")
async def generate_excel(file: UploadFile = File(...)):
    # Read the uploaded image
    contents = await file.read()
    base64_image = base64.b64encode(contents).decode('utf-8')

    # Prompt sent to GPT-4o
    prompt = """
You are an expert in reading charts. Based on this image of a chart, extract the chart data in JSON format using this schema:

{
  "title": "string",
  "xAxis": {
    "title": "string",
    "labels": ["string", ...]
  },
  "yAxis": {
    "title": "string"
  },
  "series": [
    {
      "name": "string",
      "data": [number, number, ...]
    }
  ]
}
Respond only in valid JSON.
"""

    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url", "image_url": {
                        "url": f"data:image/jpeg;base64,{base64_image}"}
                    }
                ]
            }
        ],
        max_tokens=1000
    )

    try:
        chart_data = json.loads(response.choices[0].message.content)
    except json.JSONDecodeError:
        return {"error": "Could not parse GPT response as JSON."}


    # Create Excel file
    filename = f"{uuid4().hex}.xlsx"
    filepath = f"/tmp/{filename}"
    workbook = xlsxwriter.Workbook(filepath)
    worksheet = workbook.add_worksheet("Chart Data")

    worksheet.write('A1', chart_data["xAxis"]["title"])
    worksheet.write_row('A2', ["Category"] + [s["name"] for s in chart_data["series"]])

    for i, label in enumerate(chart_data["xAxis"]["labels"]):
        row = [label] + [s["data"][i] for s in chart_data["series"]]
        worksheet.write_row(f'A{i+3}', row)

    # Add chart
    chart = workbook.add_chart({'type': 'column'})
    for i, s in enumerate(chart_data["series"]):
        chart.add_series({
            'name':       s["name"],
            'categories': f"='Chart Data'!$A$3:$A${len(chart_data['xAxis']['labels']) + 2}",
            'values':     f"='Chart Data'!${chr(66 + i)}$3:${chr(66 + i)}${len(chart_data['xAxis']['labels']) + 2}"
        })

    chart.set_title({'name': chart_data["title"]})
    chart.set_x_axis({'name': chart_data["xAxis"]["title"]})
    chart.set_y_axis({'name': chart_data["yAxis"]["title"]})

    worksheet.insert_chart('E2', chart)
    workbook.close()

    return FileResponse(filepath, filename="chart.xlsx")

@app.post("/generate-excel/")
async def generate_excel(file: UploadFile = File(...)):
    try:
        contents = await file.read()
        print("✅ Image read successfully")

        base64_image = base64.b64encode(contents).decode('utf-8')

        prompt = """... (same as before) ..."""

        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=[
                {"role": "user", "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url", "image_url": {
                        "url": f"data:image/jpeg;base64,{base64_image}"}}
                ]}
            ],
            max_tokens=1000
        )
        print("✅ GPT-4o responded")

        raw_data = response['choices'][0]['message']['content']
        print("📦 GPT raw content:", raw_data)

        graph_data = json.loads(raw_data)
        print("✅ Parsed JSON")

        # ... Excel writing code (same as before) ...

        return FileResponse(filepath, filename="chart.xlsx")

    except Exception as e:
        print("❌ Error:", str(e))
        return {"error": str(e)}

