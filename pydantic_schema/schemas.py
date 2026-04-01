from pydantic import BaseModel
from typing import Optional
from datetime import datetime

class RemarksBase(BaseModel):
    from_: Optional[str]
    to: Optional[str]
    project_id: Optional[int]
    remarks_description: Optional[str]
    respond_to_remarks: Optional[str]

class RemarksCreate(RemarksBase):
    pass

class RemarksUpdate(RemarksBase):
    pass

class RemarksResponse(RemarksBase):
    id: int
    updated_at: Optional[datetime]

    class Config:
        from_attributes = True

class TransitionResponse(BaseModel):
    id: int
    from_: Optional[str]
    to: Optional[str]
    project_id: Optional[int]
    remarks_description: Optional[str]
    respond_to_remarks: Optional[str]
    updated_at: Optional[datetime]

    class Config:
        from_attributes = True