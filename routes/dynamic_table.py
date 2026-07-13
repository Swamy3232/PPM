import os
import tempfile
import uuid
from typing import List, Dict, Any, Optional

from fastapi import APIRouter, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel
from docx.shared import RGBColor, Pt

from services.dynamic_table_headers import compute_rows_for_header, set_cell_background

router = APIRouter(prefix="/dynamic-tables", tags=["Dynamic Tables"])


# ---------- Schemas (no DB fields needed - nothing is persisted) ----------

class DynamicTableItem(BaseModel):
    header_name: str
    columns: List[str]
    rows: List[Dict[str, Any]]


class GenerateWordPayload(BaseModel):
    title: Optional[str] = "Cost Breakdown"
    created_by: Optional[str] = None
    tables: List[DynamicTableItem]


# ---------- Single endpoint: compute + generate + return file ----------

@router.post("/generate-word")
def generate_word_document(payload: GenerateWordPayload):
    """
    Stateless: takes header/column/row data straight from the frontend popup,
    applies the same calculation rules (Manpower formula, Amount auto-sum),
    builds a nicely formatted .docx in memory, and returns it for download.
    Nothing is saved to the database.
    """
    from docx import Document as DocxDocument

    if not payload.tables:
        raise HTTPException(status_code=400, detail="At least one table is required")

    doc = DocxDocument()

    # 1. Main document title, styled in blue
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

        # 2. Section subheading per table, dark gray
        section_heading = doc.add_heading(item.header_name, level=2)
        section_heading.runs[0].font.color.rgb = RGBColor(0x33, 0x33, 0x33)

        word_table = doc.add_table(rows=1, cols=len(item.columns))
        word_table.style = "Table Grid"

        # 3. Header row - light blue fill, bold text
        header_cells = word_table.rows[0].cells
        for idx, col_name in enumerate(item.columns):
            header_cells[idx].text = str(col_name)
            set_cell_background(header_cells[idx], "EDF2F7")
            for p in header_cells[idx].paragraphs:
                for run in p.runs:
                    run.font.bold = True

        # 4. Data rows - detect and highlight a "Total" row if present
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

        doc.add_paragraph()  # spacing between tables

    # 5. Grand total summary across ALL tables, shown once at the end
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