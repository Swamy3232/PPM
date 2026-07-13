import os
import tempfile
import uuid
from typing import List, Dict, Any, Optional

from fastapi import APIRouter, HTTPException, Depends
from fastapi.responses import FileResponse
from pydantic import BaseModel
from docx.shared import RGBColor, Pt
from sqlalchemy.orm import Session

from db import get_db
from models.model import DynamicTable
from services.dynamic_table_headers import compute_rows_for_header, set_cell_background

router = APIRouter(prefix="/dynamic-tables", tags=["Dynamic Tables"])


class DynamicTableItem(BaseModel):
    header_name: str
    columns: List[str]
    rows: List[Dict[str, Any]]


class GenerateWordPayload(BaseModel):
    title: Optional[str] = "Cost Breakdown"
    created_by: Optional[str] = None
    tables: List[DynamicTableItem]


@router.get("/{project_id}")
def get_saved_tables(project_id: int, db: Session = Depends(get_db)):
    """
    Returns previously saved (raw, editable) tables for a project, if any.
    Frontend uses this to pre-populate the Cost Estimation modal for editing.
    """
    rows = (
        db.query(DynamicTable)
        .filter(DynamicTable.project_id == project_id)
        .order_by(DynamicTable.id)
        .all()
    )
    return [
        {
            "header_name": r.header_name,
            "columns": r.columns,
            "rows": r.rows,
        }
        for r in rows
    ]


@router.post("/{project_id}/generate-word")
def save_and_generate_word_document(
    project_id: int, payload: GenerateWordPayload, db: Session = Depends(get_db)
):
    """
    Saves the raw (editable) table data for this project - replacing whatever
    was saved before for this project - then generates and returns a formatted .docx.
    """
    from docx import Document as DocxDocument

    if not payload.tables:
        raise HTTPException(status_code=400, detail="At least one table is required")

    # 1. Replace previously saved tables for this project with the current state
    db.query(DynamicTable).filter(DynamicTable.project_id == project_id).delete()
    for item in payload.tables:
        db.add(
            DynamicTable(
                project_id=project_id,
                header_name=item.header_name,
                columns=item.columns,
                rows=item.rows,
                created_by=payload.created_by,
            )
        )
    db.commit()

    # 2. Build the Word document from freshly computed rows (display-only, never persisted)
    doc = DocxDocument()

    title_heading = doc.add_heading(payload.title or "Cost Breakdown", level=1)
    title_heading.runs[0].font.color.rgb = RGBColor(0x2B, 0x57, 0x9A)

    if payload.created_by:
        meta = doc.add_paragraph()
        meta_run = meta.add_run(f"Prepared by: {payload.created_by}")
        meta_run.font.size = Pt(10)
        meta_run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    doc.add_paragraph()

    grand_total = 0.0

    for item in payload.tables:
        if not item.columns:
            raise HTTPException(
                status_code=400,
                detail=f"'{item.header_name}' requires at least one column",
            )

        rows, total_amount = compute_rows_for_header(
            item.header_name, item.rows, item.columns
        )
        if total_amount is not None:
            grand_total += total_amount

        section_heading = doc.add_heading(item.header_name, level=2)
        section_heading.runs[0].font.color.rgb = RGBColor(0x33, 0x33, 0x33)

        word_table = doc.add_table(rows=1, cols=len(item.columns))
        word_table.style = "Table Grid"

        header_cells = word_table.rows[0].cells
        for idx, col_name in enumerate(item.columns):
            header_cells[idx].text = str(col_name)
            set_cell_background(header_cells[idx], "EDF2F7")
            for p in header_cells[idx].paragraphs:
                for run in p.runs:
                    run.font.bold = True

        for row_data in rows:
            row_cells = word_table.add_row().cells
            first_val = str(row_data.get(item.columns[0], "") or "").lower()
            is_total_row = "total" in first_val

            for idx, col_name in enumerate(item.columns):
                cell_value = str(row_data.get(col_name, "") or "")
                row_cells[idx].text = cell_value
                if is_total_row:
                    set_cell_background(row_cells[idx], "F7FAFC")
                    for p in row_cells[idx].paragraphs:
                        for run in p.runs:
                            run.font.bold = True

        doc.add_paragraph()

    doc.add_paragraph()
    summary_para = doc.add_paragraph()
    summary_run = summary_para.add_run(
        f"The Amount for this project {payload.title or ''} is."
    )
    summary_run.font.size = Pt(11)

    total_para = doc.add_paragraph()
    total_label_run = total_para.add_run("Total Amount: ")
    total_label_run.font.bold = True
    total_label_run.font.size = Pt(12)
    total_value_run = total_para.add_run(f"{grand_total:.2f}")
    total_value_run.font.bold = True
    total_value_run.font.size = Pt(12)

    tmp_dir = tempfile.gettempdir()
    filename = f"cost_breakdown_{uuid.uuid4().hex[:8]}.docx"
    filepath = os.path.join(tmp_dir, filename)
    doc.save(filepath)

    return FileResponse(
        filepath,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=filename,
    )