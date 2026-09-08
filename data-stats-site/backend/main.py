import io
import json
import math
from pathlib import Path

import pandas as pd
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

MAX_FILE_SIZE = 20 * 1024 * 1024  # 20MB
ALLOWED_EXTENSIONS = {"csv", "xlsx"}
CSV_ENCODINGS = ("utf-8-sig", "cp949", "euc-kr", "latin1")

FRONTEND_DIR = Path(__file__).resolve().parent.parent / "frontend"

app = FastAPI(title="데이터 속성 및 기초 통계량 분석")


class AnalysisError(Exception):
    def __init__(self, message: str):
        self.message = message


def format_bytes(n: float) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.2f} {unit}"
        n /= 1024
    return f"{n:.2f} TB"


def safe_num(value):
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(value, float) and math.isinf(value):
        return None
    return float(value)


def stringify(value):
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    return str(value)


def read_csv_safe(contents: bytes) -> pd.DataFrame:
    last_err = None
    for enc in CSV_ENCODINGS:
        try:
            return pd.read_csv(io.BytesIO(contents), encoding=enc)
        except UnicodeDecodeError as e:
            last_err = e
            continue
        except Exception as e:
            raise e
    raise last_err


def read_excel_safe(contents: bytes) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(contents), engine="openpyxl")


def build_analysis(df: pd.DataFrame) -> dict:
    n_rows, n_cols = df.shape

    columns_info = []
    for col in df.columns:
        non_null = int(df[col].notna().sum())
        null_count = int(n_rows - non_null)
        null_pct = round((null_count / n_rows) * 100, 2) if n_rows else 0.0
        columns_info.append(
            {
                "name": str(col),
                "dtype": str(df[col].dtype),
                "non_null": non_null,
                "null_count": null_count,
                "null_pct": null_pct,
            }
        )

    mem_bytes = int(df.memory_usage(deep=True).sum())

    numeric_df = df.select_dtypes(include="number")
    numeric_stats = []
    if not numeric_df.empty:
        desc = numeric_df.describe().round(2)
        for col in desc.columns:
            row = desc[col]
            numeric_stats.append(
                {
                    "name": str(col),
                    "count": safe_num(row.get("count")),
                    "mean": safe_num(row.get("mean")),
                    "std": safe_num(row.get("std")),
                    "min": safe_num(row.get("min")),
                    "p25": safe_num(row.get("25%")),
                    "p50": safe_num(row.get("50%")),
                    "p75": safe_num(row.get("75%")),
                    "max": safe_num(row.get("max")),
                }
            )

    categorical_df = df.select_dtypes(exclude="number")
    categorical_stats = []
    for col in categorical_df.columns:
        series = df[col].dropna()
        count = int(series.shape[0])
        if count == 0:
            categorical_stats.append(
                {"name": str(col), "count": 0, "unique": 0, "top": None, "freq": 0}
            )
            continue
        vc = series.value_counts()
        categorical_stats.append(
            {
                "name": str(col),
                "count": count,
                "unique": int(series.nunique()),
                "top": stringify(vc.index[0]),
                "freq": int(vc.iloc[0]),
            }
        )

    preview_json = df.head(10).to_json(orient="records", date_format="iso", force_ascii=False)
    preview_records = json.loads(preview_json)
    preview_columns = [str(c) for c in df.columns]

    return {
        "shape": {"rows": n_rows, "cols": n_cols},
        "memory_bytes": mem_bytes,
        "memory_readable": format_bytes(mem_bytes),
        "columns_info": columns_info,
        "numeric_stats": numeric_stats,
        "categorical_stats": categorical_stats,
        "preview_columns": preview_columns,
        "preview_rows": preview_records,
    }


@app.post("/api/analyze")
async def analyze(file: UploadFile = File(...)):
    filename = file.filename or ""
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""

    if ext not in ALLOWED_EXTENSIONS:
        raise HTTPException(
            status_code=400,
            detail="지원하지 않는 파일 형식입니다. CSV(.csv) 또는 Excel(.xlsx) 파일만 업로드해주세요.",
        )

    contents = await file.read()

    if len(contents) == 0:
        raise HTTPException(
            status_code=400, detail="빈 파일입니다. 데이터가 포함된 파일을 업로드해주세요."
        )

    if len(contents) > MAX_FILE_SIZE:
        raise HTTPException(
            status_code=400,
            detail=f"파일 크기가 너무 큽니다. 최대 {MAX_FILE_SIZE // (1024 * 1024)}MB까지 업로드할 수 있습니다.",
        )

    try:
        if ext == "csv":
            df = read_csv_safe(contents)
        else:
            df = read_excel_safe(contents)
    except pd.errors.EmptyDataError:
        raise HTTPException(
            status_code=400, detail="빈 파일입니다. 데이터가 포함된 파일을 업로드해주세요."
        )
    except Exception:
        raise HTTPException(
            status_code=400,
            detail="파일을 읽는 중 오류가 발생했습니다. 파일 형식이나 내용을 확인해주세요.",
        )

    if df.shape[1] == 0:
        raise HTTPException(status_code=400, detail="파일에서 열(컬럼)을 찾을 수 없습니다.")

    if df.shape[0] == 0:
        raise HTTPException(
            status_code=400, detail="파일에 데이터 행이 없습니다. 헤더 외에 실제 데이터가 필요합니다."
        )

    try:
        result = build_analysis(df)
    except Exception:
        raise HTTPException(
            status_code=500, detail="데이터를 분석하는 중 오류가 발생했습니다."
        )

    return JSONResponse(content=result)


app.mount("/static", StaticFiles(directory=str(FRONTEND_DIR)), name="static")


@app.get("/")
async def index():
    return FileResponse(str(FRONTEND_DIR / "index.html"))
