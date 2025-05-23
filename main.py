from fastapi.middleware.cors import CORSMiddleware
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
import xlsxwriter
import os
from uuid import uuid4
from openai import OpenAI
import base64
import json

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Set your frontend domain here in production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def get_openai_client():
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise HTTPException(status_code=500, detail="OPENAI_API_KEY environment variable not set")
    return OpenAI(api_key=api_key)

@app.get("/")
async def root():
    return {"message": "Chart to Excel API is running"}

@app.post("/generate-excel/")
async def generate_excel(file: UploadFile = File(...)):
    try:
        client = get_openai_client()

        if not file.content_type or not file.content_type.startswith("image/"):
            raise HTTPException(status_code=400, detail="File must be an image")

        contents = await file.read()
        base64_image = base64.b64encode(contents).decode("utf-8")

        prompt = """
You are an expert in reading charts. Based on this image of a chart, extract the chart data in strict JSON format. Follow this schema:

{
  "title": "string",
  "xAxis": {
    "title": "string",
    "labels": ["string", ...]
  },
  "yAxis": {
    "title": "string",
    "min": number,
    "max": number
  },
  "legendPosition": "bottom" | "right" | "none",
  "series": [
    {
      "name": "string",
      "type": "line" | "column",
      "color": "#RRGGBB",
      "data": [number, number, ...]
    }
  ]
}

Additional rules:
- If a line crosses the chart, include it as a series with "type": "line"
- If a series appears below zero, include negative values.
- Include "color" from bar/line color if visible.
- If the chart appears stacked, assume columns share the same x-axis label.

Respond with JSON only, no markdown or preamble.
"""

        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url", "image_url": {
                        "url": f"data:image/jpeg;base64,{base64_image}"}
                    }
                ]
            }],
            max_tokens=1000
        )

        raw_response = response.choices[0].message.content
        print("🧠 Raw GPT response:", raw_response)

        cleaned = raw_response.strip()
        if cleaned.startswith("```"):
            cleaned = cleaned.strip("`").split("\n", 1)[-1]

        try:
            chart_data = json.loads(cleaned)
        except json.JSONDecodeError as json_error:
            raise HTTPException(status_code=500, detail=f"Failed to parse OpenAI response as JSON: {str(json_error)}. Raw content: {raw_response}")

        required_keys = ["title", "xAxis", "yAxis", "series"]
        if not all(key in chart_data for key in required_keys):
            raise HTTPException(status_code=500, detail="Invalid chart data structure from OpenAI")

        filename = f"{uuid4().hex}.xlsx"
        filepath = f"/tmp/{filename}"

        workbook = xlsxwriter.Workbook(filepath)
        worksheet = workbook.add_worksheet("Chart Data")

        worksheet.write("A1", chart_data["xAxis"].get("title", ""))
        worksheet.write_row("A2", ["Category"] + [s["name"] for s in chart_data["series"]])

        for i, label in enumerate(chart_data["xAxis"]["labels"]):
            row = [label] + [s["data"][i] if i < len(s["data"]) else 0 for s in chart_data["series"]]
            worksheet.write_row(f"A{i+3}", row)

        base_chart = None

        for i, s in enumerate(chart_data["series"]):
            chart_type = s.get("type", "column")
            chart = workbook.add_chart({'type': chart_type})

            series_opts = {
                'name': s["name"],
                'categories': f"='Chart Data'!$A$3:$A${len(chart_data['xAxis']['labels']) + 2}",
                'values': f"='Chart Data'!${chr(66 + i)}$3:${chr(66 + i)}${len(chart_data['xAxis']['labels']) + 2}"
            }

            # Optional color if specified
            if "color" in s:
                series_opts['fill'] = {'color': s["color"]}
                series_opts['border'] = {'none': True}

            chart.add_series(series_opts)

            if base_chart is None:
                base_chart = chart
            else:
                base_chart.combine(chart)

        base_chart.set_title({
            'name': chart_data["title"],
            'name_font': {'bold': True, 'size': 14}
        })
        base_chart.set_x_axis({
            'name': chart_data["xAxis"].get("title", ""),
            'name_font': {'bold': True}
        })

        y_axis = {'name': chart_data["yAxis"].get("title", ""), 'name_font': {'bold': True}}
        if "min" in chart_data["yAxis"]:
            y_axis["min"] = chart_data["yAxis"]["min"]
        if "max" in chart_data["yAxis"]:
            y_axis["max"] = chart_data["yAxis"]["max"]

        base_chart.set_y_axis(y_axis)
        base_chart.set_legend({'position': chart_data.get("legendPosition", "bottom")})

        worksheet.insert_chart("E2", base_chart)
        workbook.close()

        return FileResponse(
            filepath,
            filename="chart.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Unexpected error: {str(e)}")
