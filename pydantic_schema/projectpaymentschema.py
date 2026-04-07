from pydantic import BaseModel
from typing import Optional


class ProjectPaymentStageBase(BaseModel):
    name: Optional[str] = None
    project_no: Optional[str] = None
    value: Optional[str] = None   
    status: Optional[str] = None



class ProjectPaymentStageCreate(ProjectPaymentStageBase):
    name: str                 
    project_no: str
    value: str
    status: str


class ProjectPaymentStageUpdate(BaseModel):
    name: Optional[str] = None
    project_no: Optional[str] = None
    value: Optional[str] = None
    status: Optional[str] = None



class ProjectPaymentStageResponse(ProjectPaymentStageBase):
    id: int

    class Config:
        orm_mode = True   

class ProjectPaymentStageListResponse(BaseModel):
    data: list[ProjectPaymentStageResponse]