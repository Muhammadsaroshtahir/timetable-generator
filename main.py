
# uvicorn main:app --reload for runing
# after runing add /docs after link
# /docs
# /generate-core-timetable/
# /generate-elective-timetable/


from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse, JSONResponse
from Core import TimetableGenerator as CoreTimetable
from Electives import ElectivesManager as ElectiveGenerator  # Use the actual class name
import os
import shutil

app = FastAPI()

UPLOAD_DIR = "uploads"
OUTPUT_DIR = "output"

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

@app.get("/")
async def read_root():
    return {"message": "Welcome to the Timetable & Elective Generator API"}

# ---------------------------------------
# ---------- CORE TIMETABLE -------------
# ---------------------------------------
@app.post("/generate-timetable/")
async def generate_timetable(file: UploadFile = File(...)):
    try:
        temp_path = os.path.join(UPLOAD_DIR, f"core_{file.filename}")
        with open(temp_path, "wb") as f:
            shutil.copyfileobj(file.file, f)

        generator = CoreTimetable(temp_path)
        generator.run()

        output_file = os.path.join(OUTPUT_DIR, "final_output.xlsx")
        if os.path.exists(output_file):
            return FileResponse(output_file, filename="final_output.xlsx", media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            return JSONResponse(status_code=500, content={"error": "Output file not found after generation."})
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})

# ---------------------------------------
# -------- ELECTIVE GENERATION ---------
# ---------------------------------------
@app.post("/generate-electives/")
async def generate_electives():
    try:
        generator = ElectiveGenerator()
        generator.generate_electives()
        return {"message": "Electives generated successfully."}
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})

@app.post("/generate-elective-timetable/")
async def generate_elective_timetable():
    try:
        generator = ElectiveGenerator()
        generator.generate_timetable()
        return {"message": "Elective timetable generated successfully."}
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})
