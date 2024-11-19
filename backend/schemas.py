from pydantic import BaseModel


class Message(BaseModel):
    message: str


class Of(BaseModel):
    ChaveC: str
    Nome: str
    Email: str
    Matricula: int
    ServiceLine: str
    WorkItemForecast: str
    AcionamentoOF: str
    OF: int
    OFGrupoWorkitem: int
    ChavecVinculada: str
    PerfilOF: str
    GECAP: str
    Form: str
    ForecastUSTIBB: float
    ForecastReais: float
    USTIBBsOFATUAL: float
    DELTAUSTIBBs: float
    DELTAReaisForecast: float
    DPE: str
    GerenteEquipeBB: str
    RTBB: str
    FeriasAfastamentosObservacoes: str
    OnBoard: str
