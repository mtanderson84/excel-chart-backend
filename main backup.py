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

def analyze_chart_with_multiple_attempts(client, base64_image: str) -> dict:
    """Use multiple AI calls to get better accuracy"""

    # First call: General structure and type detection
    structure_prompt = '''
Analyze this chart image and identify its structure. Focus on:

1. CHART TYPE: Is this a stacked bar chart, grouped bar chart, combination chart, or other?
2. STACKING: Are bars stacked on top of each other or side by side?
3. NEGATIVE VALUES: Are there bars that extend below the zero line?
4. SERIES RELATIONSHIPS: How do the different colored sections relate to each other?

Respond with a brief analysis of the chart structure and type.
'''

    structure_response = client.chat.completions.create(
        model="gpt-4o",
        messages=[{
            "role": "user",
            "content": [
                {"type": "text", "text": structure_prompt},
                {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}}
            ]
        }],
        max_tokens=500,
        temperature=0.1
    )

    chart_structure = structure_response.choices[0].message.content
    print("📊 Chart Structure Analysis:", chart_structure)

    # Second call: Detailed data extraction with context
    data_prompt = f'''
Based on this chart analysis: "{chart_structure}"

Now extract the precise data from this chart. This appears to be a budget chart with stacked bars.

CRITICAL INSTRUCTIONS:
1. If bars are STACKED (one color on top of another), extract the individual component values, not cumulative totals
2. For stacked bars with negative and positive components:
   - Negative values (spending/debt) should be negative numbers
   - Positive values (savings/gain) should be positive numbers
   - The visual stacking means they are separate data series
3. Read each value precisely from the Y-axis scale
4. Include the exact legend names as shown

EXACT OUTPUT FORMAT (JSON only):
{{
  "title": "exact title from image",
  "chartType": "stacked_column" | "grouped_column" | "combination",
  "xAxis": {{
    "title": "x-axis title",
    "labels": ["Jan", "Feb", "Mar", ...]
  }},
  "yAxis": {{
    "title": "y-axis title", 
    "min": minimum_value_on_scale,
    "max": maximum_value_on_scale
  }},
  "legendPosition": "position of legend",
  "series": [
    {{
      "name": "exact name from legend",
      "type": "column",
      "color": "#hexcolor",
      "data": [value1, value2, ...],
      "stack": "group_name_if_stacked"
    }}
  ]
}}

Extract ALL series visible in the legend. For stacked charts, use the same stack group name.
'''

    data_response = client.chat.completions.create(
        model="gpt-4o",
        messages=[{
            "role": "user", 
            "content": [
                {"type": "text", "text": data_prompt},
                {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}}
            ]
        }],
        max_tokens=2500,
        temperature=0.1
    )

    return data_response.choices[0].message.content

def clean_json_response(raw_response: str) -> str:
    """Enhanced JSON cleaning"""
    cleaned = raw_response.strip()

    # Remove markdown code blocks
    if "```json" in cleaned:
        cleaned = cleaned.split("```json")[1].split("```")[0]
    elif "```" in cleaned:
        cleaned = cleaned.split("```")[1].split("```")[0]

    # Find JSON object using regex
    json_pattern = r'\{(?:[^{}]|{(?:[^{}]|{[^{}]*})*})*\}'
    matches = re.findall(json_pattern, cleaned, re.DOTALL)

    if matches:
        # Take the largest/most complete JSON object
        cleaned = max(matches, key=len)

    return cleaned.strip()

def create_advanced_excel_chart(chart_data: dict, worksheet, workbook):
    """Create more accurate Excel chart"""

    chart_type = chart_data.get("chartType", "column")

    if chart_type == "stacked_column" or any(s.get("stack") for s in chart_data["series"]):
        # Create stacked column chart
        chart = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})

        for i, series in enumerate(chart_data["series"]):
            col_letter = chr(66 + i)

            # For stacked charts, we need to handle positive and negative values carefully
            chart.add_series({
                'name': series["name"],
                'categories': f"='Chart Data'!$A$3:$A${len(chart_data['xAxis']['labels']) + 2}",
                'values': f"='Chart Data'!${col_letter}$3:${col_letter}${len(chart_data['xAxis']['labels']) + 2}",
                'fill': {'color': series.get("color", "#4472C4")},
                'border': {'color': '#FFFFFF', 'width': 1}
            })

    elif chart_type == "combination":
        # Handle combination charts
        primary_chart = None

        for i, series in enumerate(chart_data["series"]):
            col_letter = chr(66 + i)
            series_type = series.get("type", "column")

            if series_type == "line":
                current_chart = workbook.add_chart({'type': 'line'})
                current_chart.add_series({
                    'name': series["name"],
                    'categories': f"='Chart Data'!$A$3:$A${len(chart_data['xAxis']['labels']) + 2}",
                    'values': f"='Chart Data'!${col_letter}$3:${col_letter}${len(chart_data['xAxis']['labels']) + 2}",
                    'line': {'color': series.get("color", "#FF7F00"), 'width': 3}
                })
            else:
                current_chart = workbook.add_chart({'type': 'column'})
                current_chart.add_series({
                    'name': series["name"],
                    'categories': f"='Chart Data'!$A$3:$A${len(chart_data['xAxis']['labels']) + 2}",
                    'values': f"='Chart Data'!${col_letter}$3:${col_letter}${len(chart_data['xAxis']['labels']) + 2}",
                    'fill': {'color': series.get("color", "#4472C4")}
                })

            if primary_chart is None:
                primary_chart = current_chart
            else:
                primary_chart.combine(current_chart)

        chart = primary_chart

    else:
        # Standard grouped column chart
        chart = workbook.add_chart({'type': 'column'})

        for i, series in enumerate(chart_data["series"]):
            col_letter = chr(66 + i)
            chart.add_series({
                'name': series["name"],
                'categories': f"='Chart Data'!$A$3:$A${len(chart_data['xAxis']['labels']) + 2}",
                'values': f"='Chart Data'!${col_letter}$3:${col_letter}${len(chart_data['xAxis']['labels']) + 2}",
                'fill': {'color': series.get("color", "#4472C4")}
            })

    # Chart formatting
    chart.set_title({
        'name': chart_data.get("title", "Chart"),
        'name_font': {'bold': True, 'size': 14}
    })

    chart.set_x_axis({
        'name': chart_data["xAxis"].get("title", ""),
        'name_font': {'bold': True, 'size': 11}
    })

    y_axis_config = {
        'name': chart_data["yAxis"].get("title", ""),
        'name_font': {'bold': True, 'size': 11}
    }

    if "min" in chart_data["yAxis"]:
        y_axis_config["min"] = chart_data["yAxis"]["min"]
    if "max" in chart_data["yAxis"]:
        y_axis_config["max"] = chart_data["yAxis"]["max"]

    chart.set_y_axis(y_axis_config)

    # Legend
    legend_pos = chart_data.get("legendPosition", "bottom")
    if legend_pos != "none":
        chart.set_legend({'position': legend_pos, 'font': {'size': 10}})

    # Insert with better sizing
    worksheet.insert_chart("E2", chart, {
        'x_scale': 1.8, 
        'y_scale': 1.5,
        'x_offset': 10,
        'y_offset': 10
    })

@app.get("/")
async def root():
    return {"message": "Enhanced Chart to Excel API is running"}

@app.post("/generate-excel/")
async def generate_excel(file: UploadFile = File(...)):
    try:
        client = get_openai_client()

        if not file.content_type or not file.content_type.startswith("image/"):
            raise HTTPException(status_code=400, detail="File must be an image")

        contents = await file.read()
        base64_image = base64.b64encode(contents).decode("utf-8")

        # Use enhanced multi-step analysis
        raw_response = analyze_chart_with_multiple_attempts(client, base64_image)
        print("🧠 Raw GPT response:", raw_response)

        # Clean and parse response
        cleaned = clean_json_response(raw_response)
        print("🧹 Cleaned JSON:", cleaned)

        try:
            chart_data = json.loads(cleaned)
        except json.JSONDecodeError as json_error:
            print(f"JSON parsing failed: {str(json_error)}")
            raise HTTPException(status_code=500, detail=f"Failed to parse response as JSON: {str(json_error)}")

        # Validate data consistency
        num_labels = len(chart_data["xAxis"]["labels"])
        for series in chart_data["series"]:
            if len(series["data"]) != num_labels:
                print(f"Warning: Series '{series['name']}' has {len(series['data'])} values but {num_labels} labels")
                # Fix length mismatch
                if len(series["data"]) < num_labels:
                    series["data"].extend([0] * (num_labels - len(series["data"])))
                else:
                    series["data"] = series["data"][:num_labels]

        print("📊 Final chart data:", json.dumps(chart_data, indent=2))

        # Create Excel file with enhanced chart
        filename = f"enhanced_chart_{uuid4().hex}.xlsx"
        filepath = f"/tmp/{filename}"

        workbook = xlsxwriter.Workbook(filepath)
        worksheet = workbook.add_worksheet("Chart Data")

        # Write data to Excel
        worksheet.write("A1", "Chart Data")

        # Headers
        headers = ["Month"] + [s["name"] for s in chart_data["series"]]
        worksheet.write_row("A2", headers)

        # Data rows
        for i, label in enumerate(chart_data["xAxis"]["labels"]):
            row_data = [label] + [series["data"][i] for series in chart_data["series"]]
            worksheet.write_row(f"A{i+3}", row_data)

        # Create the chart
        create_advanced_excel_chart(chart_data, worksheet, workbook)

        workbook.close()

        return FileResponse(
            filepath,
            filename="enhanced_chart.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except HTTPException:
        raise
    except Exception as e:
        print(f"Unexpected error: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Unexpected error: {str(e)}")