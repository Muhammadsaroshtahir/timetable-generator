# Railway deployment ready FastAPI app
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import os
import shutil
from pathlib import Path

# Import your timetable generators
from Core import TimetableGenerator as CoreTimetable
from Electives import ElectivesManager as ElectiveGenerator

app = FastAPI(
    title="University Timetable Generator",
    description="Advanced timetable generation system for universities",
    version="1.0.0"
)

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Directory setup
UPLOAD_DIR = "uploads"
OUTPUT_DIR = "output"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

@app.get("/")
async def read_root():
    return {
        "message": "🎓 University Timetable Generator API",
        "status": "running",
        "endpoints": {
            "docs": "/docs",
            "health": "/health",
            "generate_core_timetable": "/generate-core-timetable/",
            "generate_electives": "/generate-electives/",
            "generate_elective_timetable": "/generate-elective-timetable/"
        },
        "version": "1.0.0"
    }

@app.get("/health")
async def health_check():
    """Health check endpoint for Railway"""
    return {
        "status": "healthy", 
        "message": "Timetable Generator API is running",
        "upload_dir_exists": os.path.exists(UPLOAD_DIR),
        "output_dir_exists": os.path.exists(OUTPUT_DIR)
    }

# ---------------------------------------
# ---------- CORE TIMETABLE -------------
# ---------------------------------------

@app.post("/generate-core-timetable/")
async def generate_core_timetable(files: list[UploadFile] = File(...)):
    """
    Generate core timetable from uploaded Excel files
    Supports the 12-file system (6 core + 6 cohort files)
    """
    try:
        # Validate files
        for file in files:
            if not file.filename.endswith(('.xlsx', '.xls')):
                raise HTTPException(
                    status_code=400, 
                    detail=f"Invalid file type: {file.filename}. Only Excel files (.xlsx, .xls) are allowed."
                )

        # Clear and recreate upload directory
        if os.path.exists(UPLOAD_DIR):
            shutil.rmtree(UPLOAD_DIR)
        os.makedirs(UPLOAD_DIR, exist_ok=True)
        
        uploaded_files = []
        
        # Save uploaded files
        for file in files:
            temp_path = os.path.join(UPLOAD_DIR, file.filename)
            with open(temp_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)
            uploaded_files.append(temp_path)
        
        # Initialize core timetable generator
        generator = CoreTimetable()
        
        # Run the generator
        success = generator.run(uploaded_files)
        
        if not success:
            raise HTTPException(
                status_code=500, 
                detail="Timetable generation failed. Check file formats and data."
            )
        
        # Find the generated output file
        output_files = [
            "Ultimate_12File_Timetable.xlsx",
            os.path.join(OUTPUT_DIR, "Ultimate_12File_Timetable.xlsx"),
            "Generated_Timetable.xlsx"
        ]
        
        output_file = None
        for file_path in output_files:
            if os.path.exists(file_path):
                output_file = file_path
                break
        
        if output_file:
            return FileResponse(
                output_file,
                filename="University_Timetable.xlsx",
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            # Return success info even if file not found for download
            return JSONResponse(
                content={
                    "message": "Timetable generated successfully",
                    "files_processed": len(uploaded_files),
                    "generator_stats": {
                        "core_placed": generator.stats.get('placed', 0),
                        "cohort_placed": generator.stats.get('cohort', 0),
                        "failed": generator.stats.get('failed', 0)
                    },
                    "note": "Timetable generated but file download not available"
                }
            )
            
    except HTTPException:
        raise
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={
                "error": f"Unexpected error: {str(e)}",
                "type": type(e).__name__,
                "message": "Please check your file formats and try again"
            }
        )

# ---------------------------------------
# -------- ELECTIVE GENERATION ---------
# ---------------------------------------

@app.post("/generate-electives/")
async def generate_electives(files: list[UploadFile] = File(...)):
    """
    Generate electives schedule from core files
    """
    try:
        # Validate files
        for file in files:
            if not file.filename.endswith(('.xlsx', '.xls')):
                raise HTTPException(
                    status_code=400,
                    detail=f"Invalid file type: {file.filename}. Only Excel files are allowed."
                )

        # Clear and setup directories
        if os.path.exists(UPLOAD_DIR):
            shutil.rmtree(UPLOAD_DIR)
        os.makedirs(UPLOAD_DIR, exist_ok=True)
        
        uploaded_files = []
        
        # Save uploaded files
        for file in files:
            temp_path = os.path.join(UPLOAD_DIR, file.filename)
            with open(temp_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)
            uploaded_files.append(temp_path)
        
        # Initialize electives generator
        generator = ElectiveGenerator()
        
        # Run electives system
        success = generator.run_electives_system(uploaded_files)
        
        if not success:
            raise HTTPException(
                status_code=500,
                detail="Electives generation failed. Check if files contain electives data."
            )
        
        return {
            "message": "✅ Electives generated successfully",
            "files_processed": len(uploaded_files),
            "statistics": {
                "total_electives": generator.stats['total_electives'],
                "sections_created": generator.stats['sections_created'],
                "cross_department": generator.stats['cross_dept_electives'],
                "students_simulated": len(generator.student_preferences)
            },
            "output_files": [
                "Electives_Timetable.xlsx",
                "Electives_Conflict_Report.xlsx", 
                "Electives_Capacity_Report.xlsx",
                "Electives_Choice_Analysis.xlsx"
            ]
        }
        
    except HTTPException:
        raise
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={
                "error": f"Error generating electives: {str(e)}",
                "type": type(e).__name__
            }
        )

@app.post("/generate-elective-timetable/")
async def generate_elective_timetable(files: list[UploadFile] = File(...)):
    """
    Generate and download elective timetable
    """
    try:
        # Process files similar to above
        if os.path.exists(UPLOAD_DIR):
            shutil.rmtree(UPLOAD_DIR)
        os.makedirs(UPLOAD_DIR, exist_ok=True)
        
        uploaded_files = []
        for file in files:
            if not file.filename.endswith(('.xlsx', '.xls')):
                raise HTTPException(status_code=400, detail=f"Invalid file: {file.filename}")
            
            temp_path = os.path.join(UPLOAD_DIR, file.filename)
            with open(temp_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)
            uploaded_files.append(temp_path)
        
        # Generate electives
        generator = ElectiveGenerator()
        success = generator.run_electives_system(uploaded_files)
        
        if not success:
            raise HTTPException(status_code=500, detail="Failed to generate elective timetable")
        
        # Look for output file
        elective_files = [
            "Electives_Timetable.xlsx",
            os.path.join(OUTPUT_DIR, "Electives_Timetable.xlsx")
        ]
        
        output_file = None
        for file_path in elective_files:
            if os.path.exists(file_path):
                output_file = file_path
                break
        
        if output_file:
            return FileResponse(
                output_file,
                filename="University_Electives_Timetable.xlsx",
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            return JSONResponse(
                content={
                    "message": "Elective timetable generated",
                    "stats": generator.stats,
                    "note": "File generated but download not available"
                }
            )
            
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"error": f"Error: {str(e)}"}
        )

# Additional utility endpoints
@app.get("/stats")
async def get_stats():
    """Get system statistics"""
    return {
        "upload_directory": UPLOAD_DIR,
        "output_directory": OUTPUT_DIR,
        "upload_dir_exists": os.path.exists(UPLOAD_DIR),
        "output_dir_exists": os.path.exists(OUTPUT_DIR),
        "available_endpoints": [
            "/",
            "/health", 
            "/stats",
            "/docs",
            "/generate-core-timetable/",
            "/generate-electives/",
            "/generate-elective-timetable/"
        ]
    }

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)