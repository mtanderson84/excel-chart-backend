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

def extract_chart_data_advanced(client, base64_image: str) -> dict:
    """Advanced chart analysis with multiple specialized prompts"""

    # Enhanced prompt specifically designed for accurate data extraction
    prompt = '''
You are a professional data analyst. Extract ALL data from this chart with MAXIMUM PRECISION.

ANALYSIS STEPS:
1. Identify the exact chart type (stacked column, grouped column, combination line+column, etc.)
2. Read every data point precisely from the Y-axis gridlines
3. For stacked charts: Extract individual component values (not cumulative totals)
4. For negative values: Use negative numbers for bars below zero line
5. Match exact colors from the legend to hex codes
6. Copy exact text from titles and labels

SPECIFIC CHART READING RULES:
- If bars are stacked (colors on top of each other), extract each color segment separately
- If you see both bars and lines, it's a combination chart
- Read values where bars/lines intersect with horizontal gridlines
- For partially filled bars, estimate the exact value between gridlines
- Negative values extend downward from the zero line

OUTPUT ONLY THIS EXACT JSON STRUCTURE:
{
  "title": "exact chart title as shown",
  "chartType": "stacked_column|grouped_column|combination|line",
  "hasNegativeValues": true/false,
  "xAxis": {
    "title": "exact x-axis title",
    "labels": ["exact", "label", "names"]
  },
  "yAxis": {
    "title": "exact y-axis title",
    "min": actual_minimum_shown,
    "max": actual_maximum_shown,
    "gridInterval": grid_spacing_if_visible
  },
  "legendPosition": "bottom|right|top|left|none",
  "series": [
    {
      "name": "exact series name from legend",
      "type": "column|line|bar",
      "color": "#RRGGBB",
      "data": [precise_value_1, precise_value_2, ...],
      "isStacked": true/false,
      "stackGroup": "group_name_if_stacked"
    }
  ]
}

CRITICAL: Respond with ONLY valid JSON. No explanations, no code blocks, no extra text.
'''

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
        temperature=0.05  # Very low temperature for consistency
    )

    return response.choices[0].message.content

def clean_and_parse_json(raw_response: str) -> dict:
    """Robust JSON cleaning and parsing"""

    print(f"🔍 Raw response length: {len(raw_response)} characters")
    print(f"🔍 Raw response preview: {raw_response[:200]}...")

    # Step 1: Clean the response
    cleaned = raw_response.strip()

    # Remove code blocks
    if "```json" in cleaned:
        cleaned = cleaned.split("```json", 1)[1].split("```", 1)[0]
    elif "```" in cleaned:
        parts = cleaned.split("```")
        if len(parts) >= 2:
            cleaned = parts[1]

    # Step 2: Find JSON boundaries
    start_idx = cleaned.find('{')
    if start_idx == -1:
        raise ValueError("No JSON object found")

    # Find matching closing brace
    brace_count = 0
    end_idx = -1
    for i in range(start_idx, len(cleaned)):
        if cleaned[i] == '{':
            brace_count += 1
        elif cleaned[i] == '}':
            brace_count -= 1
            if brace_count == 0:
                end_idx = i + 1
                break

    if end_idx == -1:
        raise ValueError("Incomplete JSON object")

    json_str = cleaned[start_idx:end_idx]

    # Step 3: Fix common JSON issues
    json_str = re.sub(r',(\s*[}\]])', r'\1', json_str)  # Remove trailing commas
    json_str = re.sub(r'\s+', ' ', json_str)  # Normalize whitespace
    json_str = json_str.replace("'", '"')  # Fix quotes

    print(f"🧹 Cleaned JSON: {json_str}")

    # Step 4: Parse JSON
    try:
        return json.loads(json_str)
    except json.JSONDecodeError as e:
        print(f"❌ JSON Parse Error: {e}")
        print(f"❌ Problematic JSON: {json_str}")
        raise

def create_excel_chart_advanced(chart_data: dict, worksheet, workbook, data_range: dict):
    """Create highly accurate Excel chart with advanced formatting"""

    chart_type = chart_data.get("chartType", "column")
    has_negative = chart_data.get("hasNegativeValues", False)

    # Determine best Excel chart approach
    if chart_type == "combination":
        # Create combination chart (bars + lines)
        primary_chart = None

        for i, series in enumerate(chart_data["series"]):
            col_letter = chr(66 + i)

            if series.get("type") == "line":
                chart = workbook.add_chart({'type': 'line'})
                chart.add_series({
                    'name': series["name"],
                    'categories': f"='Data'!$A$3:$A${data_range['last_row']}",
                    'values': f"='Data'!${col_letter}$3:${col_letter}${data_range['last_row']}",
                    'line': {
                        'color': series.get("color", "#FF7F00"),
                        'width': 3
                    },
                    'marker': {
                        'type': 'circle',
                        'size': 6,
                        'border': {'color': series.get("color", "#FF7F00")},
                        'fill': {'color': series.get("color", "#FF7F00")}
                    }
                })
            else:
                chart = workbook.add_chart({'type': 'column'})
                chart.add_series({
                    'name': series["name"],
                    'categories': f"='Data'!$A$3:$A${data_range['last_row']}",
                    'values': f"='Data'!${col_letter}$3:${col_letter}${data_range['last_row']}",
                    'fill': {'color': series.get("color", "#4472C4")},
                    'border': {'color': '#FFFFFF', 'width': 1}
                })

            if primary_chart is None:
                primary_chart = chart
            else:
                primary_chart.combine(chart)

        chart = primary_chart

    elif chart_type == "stacked_column" or any(s.get("isStacked") for s in chart_data["series"]):
        # Create stacked column chart
        chart = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})

        for i, series in enumerate(chart_data["series"]):
            col_letter = chr(66 + i)
            chart.add_series({
                'name': series["name"],
                'categories': f"='Data'!$A$3:$A${data_range['last_row']}",
                'values': f"='Data'!${col_letter}$3:${col_letter}${data_range['last_row']}",
                'fill': {'color': series.get("color", "#4472C4")},
                'border': {'color': '#FFFFFF', 'width': 1}
            })

    else:
        # Standard column/line chart
        if chart_type == "line":
            chart = workbook.add_chart({'type': 'line'})
        else:
            chart = workbook.add_chart({'type': 'column'})

        for i, series in enumerate(chart_data["series"]):
            col_letter = chr(66 + i)

            series_config = {
                'name': series["name"],
                'categories': f"='Data'!$A$3:$A${data_range['last_row']}",
                'values': f"='Data'!${col_letter}$3:${col_letter}${data_range['last_row']}",
            }

            if chart_type == "line":
                series_config['line'] = {'color': series.get("color", "#4472C4"), 'width': 2}
                series_config['marker'] = {'type': 'circle', 'size': 5}
            else:
                series_config['fill'] = {'color': series.get("color", "#4472C4")}
                series_config['border'] = {'color': '#FFFFFF', 'width': 1}

            chart.add_series(series_config)

    # Advanced chart formatting
    chart.set_title({
        'name': chart_data.get("title", "Chart"),
        'name_font': {'name': 'Calibri', 'size': 16, 'bold': True, 'color': '#333333'}
    })

    chart.set_x_axis({
        'name': chart_data["xAxis"].get("title", ""),
        'name_font': {'name': 'Calibri', 'size': 12, 'bold': True, 'color': '#666666'},
        'num_font': {'name': 'Calibri', 'size': 10, 'color': '#666666'},
        'line': {'color': '#D9D9D9'},
        'major_tick_mark': 'outside',
        'minor_tick_mark': 'none'
    })

    # Y-axis with precise scaling
    y_axis_config = {
        'name': chart_data["yAxis"].get("title", ""),
        'name_font': {'name': 'Calibri', 'size': 12, 'bold': True, 'color': '#666666'},
        'num_font': {'name': 'Calibri', 'size': 10, 'color': '#666666'},
        'line': {'color': '#D9D9D9'},
        'major_gridlines': {'visible': True, 'line': {'color': '#E5E5E5', 'width': 0.75}},
        'minor_gridlines': {'visible': False}
    }

    # Set exact Y-axis range if available
    if "min" in chart_data["yAxis"] and "max" in chart_data["yAxis"]:
        y_axis_config["min"] = chart_data["yAxis"]["min"]
        y_axis_config["max"] = chart_data["yAxis"]["max"]

        # Set major unit if grid interval is specified
        if "gridInterval" in chart_data["yAxis"]:
            y_axis_config["major_unit"] = chart_data["yAxis"]["gridInterval"]

    chart.set_y_axis(y_axis_config)

    # Legend formatting
    legend_pos = chart_data.get("legendPosition", "bottom")
    if legend_pos != "none":
        chart.set_legend({
            'position': legend_pos,
            'font': {'name': 'Calibri', 'size': 10, 'color': '#666666'}
        })
    else:
        chart.set_legend({'none': True})

    # Chart area formatting
    chart.set_chartarea({
        'border': {'none': True},
        'fill': {'color': '#FFFFFF'}
    })

    chart.set_plotarea({
        'border': {'color': '#D9D9D9', 'width': 0.75},
        'fill': {'color': '#FFFFFF'}
    })

    # Insert chart with optimal sizing
    worksheet.insert_chart("F2", chart, {
        'x_scale': 2.0,
        'y_scale': 1.8,
        'x_offset': 15,
        'y_offset': 15
    })

def create_data_validation_sheet(workbook, chart_data: dict):
    """Create a sheet with data validation and instructions"""

    validation_sheet = workbook.add_worksheet("Instructions & Validation")

    # Header format
    header_format = workbook.add_format({
        'bold': True,
        'font_size': 14,
        'bg_color': '#4472C4',
        'font_color': 'white',
        'align': 'center'
    })

    # Info format
    info_format = workbook.add_format({
        'text_wrap': True,
        'valign': 'top',
        'font_size': 11
    })

    validation_sheet.write("A1", "Chart Recreation Instructions", header_format)
    validation_sheet.merge_range("A1:E1", "Chart Recreation Instructions", header_format)

    instructions = [
        "1. The 'Data' sheet contains the extracted chart data",
        "2. The chart has been recreated based on AI analysis of your image",
        "3. You can modify the data values to adjust the chart",
        "4. The chart will automatically update when you change data values",
        "5. Chart colors and formatting match the original as closely as possible"
    ]

    for i, instruction in enumerate(instructions):
        validation_sheet.write(f"A{i+3}", instruction, info_format)

    # Chart analysis summary
    validation_sheet.write("A10", "Extracted Chart Information:", header_format)
    validation_sheet.merge_range("A10:E10", "Extracted Chart Information:", header_format)

    summary_data = [
        ["Chart Title:", chart_data.get("title", "")],
        ["Chart Type:", chart_data.get("chartType", "")],
        ["X-axis Title:", chart_data["xAxis"].get("title", "")],
        ["Y-axis Title:", chart_data["yAxis"].get("title", "")],
        ["Y-axis Range:", f"{chart_data['yAxis'].get('min', 'auto')} to {chart_data['yAxis'].get('max', 'auto')}"],
        ["Number of Series:", str(len(chart_data["series"]))],
        ["Data Points:", str(len(chart_data["xAxis"]["labels"]))]
    ]

    for i, (label, value) in enumerate(summary_data):
        validation_sheet.write(f"A{i+12}", label, workbook.add_format({'bold': True}))
        validation_sheet.write(f"B{i+12}", value, info_format)

    # Set column widths
    validation_sheet.set_column("A:A", 20)
    validation_sheet.set_column("B:E", 30)

@app.get("/")
async def root():
    return {"message": "Advanced Excel Chart Recreation API"}

@app.post("/generate-excel/")
async def generate_excel(file: UploadFile = File(...)):
    try:
        client = get_openai_client()

        if not file.content_type or not file.content_type.startswith("image/"):
            raise HTTPException(status_code=400, detail="File must be an image")

        contents = await file.read()
        base64_image = base64.b64encode(contents).decode("utf-8")

        # Extract chart data with advanced analysis
        print("🔍 Starting advanced chart analysis...")
        raw_response = extract_chart_data_advanced(client, base64_image)

        # Parse the response
        chart_data = clean_and_parse_json(raw_response)

        # Validate data consistency
        num_labels = len(chart_data["xAxis"]["labels"])
        for series in chart_data["series"]:
            if len(series["data"]) != num_labels:
                print(f"⚠️ Length mismatch for series '{series['name']}': {len(series['data'])} vs {num_labels}")
                # Fix the mismatch
                if len(series["data"]) < num_labels:
                    series["data"].extend([0] * (num_labels - len(series["data"])))
                else:
                    series["data"] = series["data"][:num_labels]

        print("✅ Final chart data:", json.dumps(chart_data, indent=2))

        # Create advanced Excel file
        filename = f"advanced_chart_{uuid4().hex}.xlsx"
        filepath = f"/tmp/{filename}"

        workbook = xlsxwriter.Workbook(filepath, {
            'default_date_format': 'dd/mm/yy',
            'remove_timezone': True
        })

        # Create main data sheet
        data_sheet = workbook.add_worksheet("Data")

        # Formatting
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': 'white',
            'align': 'center',
            'border': 1
        })

        data_format = workbook.add_format({
            'align': 'center',
            'border': 1,
            'num_format': '#,##0.00'
        })

        # Write headers
        headers = [chart_data["xAxis"].get("title") or "Category"] + [s["name"] for s in chart_data["series"]]
        for col, header in enumerate(headers):
            data_sheet.write(1, col, header, header_format)

        # Write data
        last_row = 2
        for i, label in enumerate(chart_data["xAxis"]["labels"]):
            data_sheet.write(last_row + i, 0, label, data_format)
            for j, series in enumerate(chart_data["series"]):
                value = series["data"][i]
                data_sheet.write(last_row + i, j + 1, value, data_format)

        last_row = last_row + len(chart_data["xAxis"]["labels"]) - 1

        # Set column widths
        data_sheet.set_column(0, 0, 15)  # Category column
        data_sheet.set_column(1, len(chart_data["series"]), 12)  # Data columns

        # Create the advanced chart
        data_range = {"last_row": last_row}
        create_excel_chart_advanced(chart_data, data_sheet, workbook, data_range)

        # Create validation and instructions sheet
        create_data_validation_sheet(workbook, chart_data)

        workbook.close()

        return FileResponse(
            filepath,
            filename="advanced_chart_recreation.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except HTTPException:
        raise
    except Exception as e:
        print(f"💥 Unexpected error: {str(e)}")
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Unexpected error: {str(e)}")

