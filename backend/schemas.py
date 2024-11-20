from typing import Dict, List

from pydantic import BaseModel


class ListaOFs(BaseModel):
    resposta: List[Dict]
    quantidade: int


class Quantidade(BaseModel):
    quantidade: int
