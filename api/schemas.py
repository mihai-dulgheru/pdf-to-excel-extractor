from pydantic import BaseModel


class ProcessRequest(BaseModel):
    percentage: float
