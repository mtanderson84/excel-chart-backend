import logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)
from fastapi.middleware.cors import CORSMiddleware
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
import xlsxwriter
import os
from uuid import uuid4
from openai import OpenAI
import base64
import json
import re

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def get_openai_client():
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise HTTPException(status_code=500, detail="OPENAI_API_KEY environment variable not set")
    return OpenAI(api_key=api_key)

def clean_gpt_response(raw_response: str) -> dict:
    cleaned = raw_response.strip()
    if "```json" in cleaned:
        cleaned = cleaned.split("```json", 1)[1].split("```", 1)[0]
    elif "```" in cleaned:
        cleaned = cleaned.split("```", 1)[1]

    start_idx = cleaned.find('{')
    end_idx = cleaned.rfind('}') + 1
    json_str = cleaned[start_idx:end_idx]
    json_str = re.sub(r',\s*[}\]]', lambda m: m.group(0).lstrip(','), json_str)
    json_str = json_str.replace("'", '"')

    return json.loads(json_str)

def infer_y_axis_range(chart_data):
    all_values = [v for s in chart_data["series"] for v in s["data"] if isinstance(v, (int, float))]
    return min(all_values, default=0), max(all_values, default=100)

def build_excel_file(chart_data: dict, filepath: str):
    workbook = xlsxwriter.Workbook(filepath)
    worksheet = workbook.add_worksheet("Chart Data")

    worksheet.write("A1", chart_data["xAxis"].get("title", ""))
    worksheet.write_row("A2", ["Category"] + [s["name"] for s in chart_data["series"]])

    for i, label in enumerate(chart_data["xAxis"]["labels"]):
        row = [label] + [s["data"][i] if i < len(s["data"]) else 0 for s in chart_data["series"]]
        worksheet.write_row(f"A{i+3}", row)

    chart_type = chart_data.get("chartType", "column").lower()
    num_labels = len(chart_data["xAxis"]["labels"])
    base_chart = workbook.add_chart({'type': 'column'})

    for i, s in enumerate(chart_data["series"]):
        series_type_raw = s.get("type", "column").lower()
        if series_type_raw in ["line", "column"]:
            series_type = series_type_raw
        elif "line" in series_type_raw:
            series_type = "line"
        else:
            series_type = "column"

        is_stacked = s.get("isStacked", False)
        color = s.get("color")

        chart_args = {'type': series_type}
        if chart_type == "stacked_column" or is_stacked:
            if series_type == "column":
                chart_args["subtype"] = "stacked"

        chart = workbook.add_chart(chart_args)

        col_letter = chr(66 + i)
        series_conf = {
            'name': s["name"],
            'categories': f"='Chart Data'!$A$3:$A${num_labels + 2}",
            'values': f"='Chart Data'!${col_letter}$3:${col_letter}${num_labels + 2}"
        }

        if series_type == "line":
            series_conf["line"] = {"color": color or "#000000"}
            series_conf["marker"] = {"type": "circle", "size": 6}
        elif color:
            series_conf["fill"] = {"color": color}
            series_conf["border"] = {"color": "#FFFFFF"}

        chart.add_series(series_conf)
        base_chart.combine(chart)

    base_chart.set_title({
        'name': chart_data.get("title", "Chart"),
        'name_font': {'bold': True, 'size': 14}
    })
    base_chart.set_x_axis({
        'name': chart_data["xAxis"].get("title", ""),
        'name_font': {'bold': True}
    })

    y_axis = {
        'name': chart_data["yAxis"].get("title", ""),
        'name_font': {'bold': True}
    }
    if "min" in chart_data["yAxis"]:
        y_axis["min"] = chart_data["yAxis"]["min"]
    if "max" in chart_data["yAxis"]:
        y_axis["max"] = chart_data["yAxis"]["max"]
    base_chart.set_y_axis(y_axis)

    base_chart.set_legend({'position': chart_data.get("legendPosition", "bottom")})

    worksheet.insert_chart("E2", base_chart)
    workbook.close()

def generate_prompt():
    return '''
You are a chart interpretation expert. Analyze the uploaded image and extract chart data in a structured JSON format.

IMPORTANT:
- Identify the chart type: column, stacked_column, bar, line, or combination.
- Match exact values for every series and label.
- Detect axis titles, legends, series types, colors, and y-axis scale.

Respond in this schema:
{
  "title": "Chart title",
  "chartType": "column|stacked_column|line|bar|combination",
  "xAxis": {
    "title": "X-axis label",
    "labels": ["Jan", "Feb", ...]
  },
  "yAxis": {
    "title": "Y-axis label",
    "min": number,
    "max": number,
    "gridInterval": number
  },
  "legendPosition": "bottom|top|left|right|none",
  "series": [
    {
      "name": "Series name",
      "type": "line|column",
      "data": [numbers],
      "color": "#RRGGBB",
      "isStacked": true|false
    }
  ]
}

Only respond with valid JSON — no markdown, no commentary.
'''

@app.get("/")
async def root():
    return {"message": "Chart to Excel API is running"}

@app.post("/generate-excel/")
async def generate_excel(file: UploadFile = File(...)):
    try:
        client = get_openai_client()

        if not file.content_type.startswith("image/"):
            raise HTTPException(status_code=400, detail="File must be an image")

        contents = await file.read()
        base64_image = base64.b64encode(contents).decode("utf-8")

        prompt = generate_prompt()

        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}}
                ]
            }],
            max_tokens=3000,
            temperature=0.05
        )

        raw_response = response.choices[0].message.content
        logger.info("🧠 Raw GPT response (first 500 chars): %s", raw_response[:500])
        chart_data = clean_gpt_response(raw_response)

        # Fix label/series mismatch
        num_labels = len(chart_data["xAxis"]["labels"])
        for series in chart_data["series"]:
            if len(series["data"]) < num_labels:
                series["data"].extend([0] * (num_labels - len(series["data"])))
            elif len(series["data"]) > num_labels:
                series["data"] = series["data"][:num_labels]

        filename = f"chart_{uuid4().hex}.xlsx"
        filepath = f"/tmp/{filename}"
        logger.info("✅ Cleaned chart data:\n%s", json.dumps(chart_data, indent=2))
        build_excel_file(chart_data, filepath)

        return FileResponse(
            filepath,
            filename="chart.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except HTTPException:
        raise
    except Exception as e:
        import traceback
        logger.error("❌ Unexpected error: %s", str(e))
        logger.error(traceback.format_exc())
        raise HTTPException(status_code=500, detail=f"Unexpected error: {str(e)}")



