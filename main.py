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
    allow_origins=["*"],  # Or replace * with your frontend Replit domain for security
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ✅ Check if API key exists and initialize client
def get_openai_client():
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise HTTPException(
            status_code=500, 
            detail="OPENAI_API_KEY environment variable not set"
        )
    return OpenAI(api_key=api_key)

@app.get("/")
async def root():
    return {"message": "Chart to Excel API is running"}

@app.get("/health")
async def health_check():
    # Check if API key is available
    api_key = os.environ.get("OPENAI_API_KEY")
    return {
        "status": "healthy",
        "openai_key_configured": bool(api_key)
    }

@app.post("/generate-excel/")
async def generate_excel(file: UploadFile = File(...)):
    try:
        # Initialize OpenAI client with error handling
        client = get_openai_client()

        # Validate file type
        if not file.content_type or not file.content_type.startswith('image/'):
            raise HTTPException(status_code=400, detail="File must be an image")

        contents = await file.read()
        base64_image = base64.b64encode(contents).decode('utf-8')

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

        # ✅ Use modern OpenAI SDK with error handling
        try:
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
        except Exception as openai_error:
            raise HTTPException(
                status_code=500, 
                detail=f"OpenAI API error: {str(openai_error)}"
            )

        # Parse JSON response
        try:
            raw_response = response.choices[0].message.content
            print("🧠 Raw GPT response:", raw_response)

            try:
                import re

                # Remove ```json and ``` if present
                cleaned_response = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw_response.strip(), flags=re.DOTALL)

                chart_data = json.loads(cleaned_response)

            except json.JSONDecodeError as json_error:
                raise HTTPException(
                    status_code=500, 
                    detail=f"Failed to parse OpenAI response as JSON: {str(json_error)}. Raw content: {raw_response}"
                )

        except json.JSONDecodeError as json_error:
            raise HTTPException(
                status_code=500, 
                detail=f"Failed to parse OpenAI response as JSON: {str(json_error)}"
            )

        # Validate chart_data structure
        required_keys = ["title", "xAxis", "yAxis", "series"]
        if not all(key in chart_data for key in required_keys):
            raise HTTPException(
                status_code=500, 
                detail="Invalid chart data structure from OpenAI"
            )

        # ✅ Create Excel file
        filename = f"{uuid4().hex}.xlsx"
        filepath = f"/tmp/{filename}"

        try:
            workbook = xlsxwriter.Workbook(filepath)
            worksheet = workbook.add_worksheet("Chart Data")

            # Write headers
            worksheet.write('A1', chart_data["xAxis"]["title"])
            worksheet.write_row('A2', ["Category"] + [s["name"] for s in chart_data["series"]])

            # Write data
            for i, label in enumerate(chart_data["xAxis"]["labels"]):
                row = [label] + [s["data"][i] if i < len(s["data"]) else 0 for s in chart_data["series"]]
                worksheet.write_row(f'A{i+3}', row)

            # Create chart
            chart = workbook.add_chart({'type': 'column'})
            for i, s in enumerate(chart_data["series"]):
                chart.add_series({
                    'name': s["name"],
                    'categories': f"='Chart Data'!$A$3:$A${len(chart_data['xAxis']['labels']) + 2}",
                    'values': f"='Chart Data'!${chr(66 + i)}$3:${chr(66 + i)}${len(chart_data['xAxis']['labels']) + 2}"
                })

            chart.set_title({'name': chart_data["title"]})
            chart.set_x_axis({'name': chart_data["xAxis"]["title"]})
            chart.set_y_axis({'name': chart_data["yAxis"]["title"]})

            worksheet.insert_chart('E2', chart)
            workbook.close()

        except Exception as excel_error:
            raise HTTPException(
                status_code=500, 
                detail=f"Failed to create Excel file: {str(excel_error)}"
            )

        return FileResponse(
            filepath, 
            filename="chart.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except HTTPException:
        # Re-raise HTTP exceptions
        raise
    except Exception as e:
        # Catch all other exceptions
        raise HTTPException(status_code=500, detail=f"Unexpected error: {str(e)}")

