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
    allow_origins=["*"],  # Replace with frontend domain in production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def get_openai_client():
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise HTTPException(status_code=500, detail="OPENAI_API_KEY environment variable not set")
    return OpenAI(api_key=api_key)

def clean_json_response(raw_response: str) -> str:
    """Clean and extract JSON from GPT response"""
    cleaned = raw_response.strip()

    # Remove code blocks if present
    if cleaned.startswith("```"):
        lines = cleaned.split("\n")
        start_idx = 1 if lines[0].startswith("```") else 0
        end_idx = len(lines)
        for i in range(len(lines) - 1, -1, -1):
            if lines[i].strip() == "```":
                end_idx = i
                break
        cleaned = "\n".join(lines[start_idx:end_idx])

    # Find JSON object boundaries
    json_match = re.search(r'\{.*\}', cleaned, re.DOTALL)
    if json_match:
        cleaned = json_match.group()

    return cleaned

def validate_and_fix_chart_data(chart_data: dict) -> dict:
    """Validate and fix common issues in chart data"""

    # Ensure required fields exist
    if "xAxis" not in chart_data:
        chart_data["xAxis"] = {"title": "", "labels": []}
    if "yAxis" not in chart_data:
        chart_data["yAxis"] = {"title": ""}
    if "series" not in chart_data:
        chart_data["series"] = []

    # Ensure all series have same number of data points as x-axis labels
    num_labels = len(chart_data["xAxis"].get("labels", []))
    for series in chart_data["series"]:
        if "data" not in series:
            series["data"] = [0] * num_labels
        elif len(series["data"]) < num_labels:
            # Pad with zeros if data is shorter
            series["data"].extend([0] * (num_labels - len(series["data"])))
        elif len(series["data"]) > num_labels:
            # Truncate if data is longer
            series["data"] = series["data"][:num_labels]

    # Set default values for missing fields
    for series in chart_data["series"]:
        if "type" not in series:
            series["type"] = "column"
        if "color" not in series:
            series["color"] = "#4472C4"  # Default Excel blue
        if "name" not in series:
            series["name"] = f"Series {len(chart_data['series'])}"

    return chart_data

def create_excel_chart(chart_data: dict, worksheet, workbook) -> None:
    """Create Excel chart with improved accuracy"""

    # Determine if we need to combine different chart types
    chart_types = list(set(s.get("type", "column") for s in chart_data["series"]))

    if len(chart_types) == 1:
        # Single chart type
        chart_type = chart_types[0]
        if chart_type == "line":
            chart = workbook.add_chart({'type': 'line'})
        else:
            chart = workbook.add_chart({'type': 'column'})

        # Add all series to the single chart
        for i, series in enumerate(chart_data["series"]):
            col_letter = chr(66 + i)  # B, C, D, etc.
            chart.add_series({
                'name': series["name"],
                'categories': f"='Chart Data'!$A$3:$A${len(chart_data['xAxis']['labels']) + 2}",
                'values': f"='Chart Data'!${col_letter}$3:${col_letter}${len(chart_data['xAxis']['labels']) + 2}",
                'line': {'color': series.get("color", "#4472C4")} if series.get("type") == "line" else None,
                'fill': {'color': series.get("color", "#4472C4")} if series.get("type") == "column" else None
            })
    else:
        # Mixed chart types - create combination chart
        primary_chart = None

        for i, series in enumerate(chart_data["series"]):
            col_letter = chr(66 + i)
            series_type = series.get("type", "column")

            if series_type == "line":
                chart = workbook.add_chart({'type': 'line'})
            else:
                chart = workbook.add_chart({'type': 'column'})

            chart.add_series({
                'name': series["name"],
                'categories': f"='Chart Data'!$A$3:$A${len(chart_data['xAxis']['labels']) + 2}",
                'values': f"='Chart Data'!${col_letter}$3:${col_letter}${len(chart_data['xAxis']['labels']) + 2}",
                'line': {'color': series.get("color", "#4472C4")} if series_type == "line" else None,
                'fill': {'color': series.get("color", "#4472C4")} if series_type == "column" else None
            })

            if primary_chart is None:
                primary_chart = chart
            else:
                primary_chart.combine(chart)

        chart = primary_chart

    # Set chart formatting
    chart.set_title({
        'name': chart_data.get("title", "Chart"),
        'name_font': {'bold': True, 'size': 14}
    })

    chart.set_x_axis({
        'name': chart_data["xAxis"].get("title", ""),
        'name_font': {'bold': True}
    })

    # Set Y-axis with min/max if specified
    y_axis_config = {
        'name': chart_data["yAxis"].get("title", ""),
        'name_font': {'bold': True}
    }
    if "min" in chart_data["yAxis"]:
        y_axis_config["min"] = chart_data["yAxis"]["min"]
    if "max" in chart_data["yAxis"]:
        y_axis_config["max"] = chart_data["yAxis"]["max"]

    chart.set_y_axis(y_axis_config)

    # Set legend position
    legend_position = chart_data.get("legendPosition", "bottom")
    if legend_position != "none":
        chart.set_legend({'position': legend_position})
    else:
        chart.set_legend({'none': True})

    # Insert chart into worksheet
    worksheet.insert_chart("E2", chart, {'x_scale': 1.5, 'y_scale': 1.2})

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

        # Enhanced prompt with better instructions
        prompt = '''
You are an expert chart data extraction specialist. Analyze this chart image and extract ALL visible data with maximum precision.

CRITICAL REQUIREMENTS:
1. Extract EXACT numerical values for each data point - read carefully from the chart
2. Include ALL visible data series (bars, lines, areas, etc.)
3. Identify negative values correctly (below zero line)
4. For stacked charts, provide individual component values, not cumulative
5. Match colors as closely as possible using hex codes
6. Preserve exact axis titles and series names as shown
7. Include precise Y-axis min/max values if visible on the chart

CHART ANALYSIS STEPS:
1. Identify chart type for each series (column/bar/line/area)
2. Read X-axis labels exactly as shown
3. For each data point, read the precise Y-value from the axis scale
4. Note if values are positive or negative
5. Identify series colors and names from legend
6. Record axis titles and chart title exactly

OUTPUT SCHEMA (JSON only, no commentary):
{
  "title": "exact chart title",
  "xAxis": {
    "title": "exact x-axis title",
    "labels": ["label1", "label2", ...]
  },
  "yAxis": {
    "title": "exact y-axis title",
    "min": actual_minimum_value_shown,
    "max": actual_maximum_value_shown
  },
  "legendPosition": "bottom" | "right" | "top" | "left" | "none",
  "series": [
    {
      "name": "exact series name from legend",
      "type": "column" | "line" | "bar" | "area",
      "color": "#RRGGBB",
      "data": [precise_value1, precise_value2, ...]
    }
  ]
}

RESPOND WITH ONLY THE JSON OBJECT.
'''

        # Increase token limit for complex charts
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
            max_tokens=2000,  # Increased token limit
            temperature=0.1   # Lower temperature for more consistent results
        )

        raw_response = response.choices[0].message.content
        print("🧠 Raw GPT response:", raw_response)

        # Clean and parse JSON response
        cleaned = clean_json_response(raw_response)

        try:
            chart_data = json.loads(cleaned)
        except json.JSONDecodeError as json_error:
            print(f"JSON parsing failed: {str(json_error)}")
            print(f"Cleaned content: {cleaned}")
            raise HTTPException(status_code=500, detail=f"Failed to parse OpenAI response as JSON: {str(json_error)}")

        # Validate and fix chart data
        chart_data = validate_and_fix_chart_data(chart_data)
        print("📊 Processed chart data:", json.dumps(chart_data, indent=2))

        # Generate Excel file
        filename = f"chart_{uuid4().hex}.xlsx"
        filepath = f"/tmp/{filename}"

        workbook = xlsxwriter.Workbook(filepath)
        worksheet = workbook.add_worksheet("Chart Data")

        # Write headers and data
        worksheet.write("A1", chart_data["xAxis"].get("title", "Category"))

        # Write column headers
        headers = ["Category"] + [s["name"] for s in chart_data["series"]]
        worksheet.write_row("A2", headers)

        # Write data rows
        for i, label in enumerate(chart_data["xAxis"]["labels"]):
            row_data = [label]
            for series in chart_data["series"]:
                value = series["data"][i] if i < len(series["data"]) else 0
                row_data.append(value)
            worksheet.write_row(f"A{i+3}", row_data)

        # Create chart with improved accuracy
        create_excel_chart(chart_data, worksheet, workbook)

        workbook.close()

        return FileResponse(
            filepath, 
            filename="extracted_chart.xlsx", 
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except HTTPException:
        raise
    except Exception as e:
        print(f"Unexpected error: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Unexpected error: {str(e)}")