from fastapi import FastAPI, Query, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from utils import list_tables, get_table_details, row_sum

app = FastAPI(
    title="Excel Processor API",
    description="API to interact with Excel data from capbudg.xls",
    version="1.0.0"
)

# Allow CORS (useful for Postman or web clients)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
def read_root():
    """
    Root endpoint to check if the API is running.
    """
    return {"message": "Welcome to the Excel Processor API!"}

@app.get("/list_tables")
def list_all_tables():
    """
    Lists all sheet names (tables) in the Excel file.
    """
    try:
        return {"tables": list_tables()}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    
@app.get("/get_table_details")
def get_details(table_name: str = Query(..., description="Name of the table (sheet)")):
    details = get_table_details(table_name)
    if details is None:
        raise HTTPException(status_code=404, detail="Table not found.")
    return {"table_name": table_name, "row_names": details}

@app.get("/row_sum")
def get_row_sum(
    table_name: str = Query(..., description="Name of the table (sheet)"),
    row_name: str = Query(..., description="Label of the row (first-column value)")
):
    """
    Calculates the sum of numeric values in a specific row.
    """
    result = row_sum(table_name, row_name)
    if result is None:
        raise HTTPException(status_code=404, detail="Table or row not found.")
    return {
        "table_name": table_name,
        "row_name": row_name,
        "sum": result
    }
