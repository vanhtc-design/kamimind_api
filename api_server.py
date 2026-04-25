import os
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import JSONResponse
from typing import List
from fastapi.middleware.cors import CORSMiddleware
import uvicorn
import tempfile
import json

# Import components
from flatten_ctdt import extract_ctdt_data
from rule_based_checker import RuleBasedChecker

app = FastAPI(
    title="KamiMind Quality Assurance API",
    description="Backend services for academic curriculum flattening and syllabus rule-based auditing.",
    version="2.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Shared database path
DB_PATH = "db_mon_hoc.json"

@app.get("/api/v1/health")
async def health_check():
    return {"status": "healthy", "db_exists": os.path.exists(DB_PATH)}

@app.post("/api/v1/flatten-ctdt", summary="Flatten Curriculum (Stage 1)")
async def flatten_ctdt(files: List[UploadFile] = File(...)):
    """
    Extracts structured data (Courses, Mapping Matrix, PLOs) from one or more Curriculum (.docx) files.
    Consolidates the results into a central database.
    """
    if not files:
        raise HTTPException(status_code=400, detail="No files uploaded.")
        
    try:
        all_courses = []
        all_matrix = []
        all_plos = []
        processed_files = []

        for file in files:
            if not file.filename.endswith('.docx'):
                continue
                
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_file:
                content = await file.read()
                temp_file.write(content)
                temp_path = temp_file.name

            try:
                courses, matrix, plos = extract_ctdt_data(temp_path)
                all_courses.extend(courses)
                all_matrix.extend(matrix)
                all_plos.extend(plos)
                processed_files.append(file.filename)
            finally:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
        
        if not processed_files:
            raise HTTPException(status_code=400, detail="No valid .docx files found.")

        # Consolidate and save to local DB
        db_data = {
            "courses": all_courses,
            "matrix": all_matrix,
            "plos": all_plos,
            "total_files": len(processed_files)
        }
        with open(DB_PATH, "w", encoding="utf-8") as f:
            json.dump(db_data, f, ensure_ascii=False, indent=2)

        return JSONResponse(content={
            "status": "success",
            "message": f"Successfully processed {len(processed_files)} documents.",
            "data": db_data
        })
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Processing error: {str(e)}")

@app.post("/api/v1/check-syllabus", summary="Rule-Based Syllabus Audit (Stage 2)")
async def check_syllabus(file: UploadFile = File(...)):
    """
    Audits a Syllabus (.docx) file against the Curriculum database and academic rules.
    """
    if not file.filename.endswith('.docx'):
        raise HTTPException(status_code=400, detail="Only .docx files are supported.")
        
    try:
        # Initialize Checker with the consolidated DB
        checker = RuleBasedChecker(DB_PATH)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_file:
            content = await file.read()
            temp_file.write(content)
            temp_path = temp_file.name

        try:
            errors = checker.check_syllabus(temp_path)
        finally:
            if os.path.exists(temp_path):
                os.remove(temp_path)

        return JSONResponse(content={
            "status": "success",
            "file_checked": file.filename,
            "total_errors": len(errors),
            "errors": errors,
            "is_valid": len(errors) == 0
        })
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Audit error: {str(e)}")

if __name__ == "__main__":
    uvicorn.run("api_server:app", host="0.0.0.0", port=8000, reload=True)
