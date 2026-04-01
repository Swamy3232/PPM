from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session
from datetime import datetime

from db import get_db
from models.model import Remarks
# from pydantic_schema.response import schemas
# ✅ CORRECT
from pydantic_schema import schemas

router = APIRouter(prefix="/Remarkss", tags=["Remarkss"])


# ---------------- CREATE ----------------
@router.post("/", response_model=schemas.TransitionResponse)
def create_Remarks(data: schemas.RemarksCreate, db: Session = Depends(get_db)):
    new_item = Remarks(
        from_=data.from_,
        to=data.to,
        project_id=data.project_id,
        remarks_description=data.remarks_description,
        respond_to_remarks=data.respond_to_remarks,
        updated_at=datetime.now()
    )
    db.add(new_item)
    db.commit()
    db.refresh(new_item)
    return new_item


# ---------------- READ ALL ----------------
@router.get("/", response_model=list[schemas.RemarksResponse])
def get_all_Remarkss(db: Session = Depends(get_db)):
    return db.query(Remarks).all()


# ---------------- READ ONE ----------------
@router.get("/{id}", response_model=schemas.TransitionResponse)
def get_Remarks(id: int, db: Session = Depends(get_db)):
    item = db.query(Remarks).filter(Remarks.id == id).first()
    if not item:
        raise HTTPException(status_code=404, detail="Not found")
    return item


# ---------------- UPDATE ----------------
@router.put("/{id}", response_model=schemas.TransitionResponse)
def update_Remarks(id: int, data: schemas.RemarksUpdate, db: Session = Depends(get_db)):
    item = db.query(Remarks).filter(Remarks.id == id).first()
    if not item:
        raise HTTPException(status_code=404, detail="Not found")

    if data.from_ is not None:
        item.from_ = data.from_
    if data.to is not None:
        item.to = data.to
    if data.project_id is not None:
        item.project_id = data.project_id
    if data.remarks_description is not None:
        item.remarks_description = data.remarks_description
    if data.respond_to_remarks is not None:
        item.respond_to_remarks = data.respond_to_remarks

    item.updated_at = datetime.now()

    db.commit()
    db.refresh(item)
    return item


# ---------------- DELETE ----------------
@router.delete("/{id}")
def delete_Remarks(id: int, db: Session = Depends(get_db)):
    item = db.query(Remarks).filter(Remarks.id == id).first()
    if not item:
        raise HTTPException(status_code=404, detail="Not found")

    db.delete(item)
    db.commit()
    return {"message": "Deleted successfully"}