from pydantic import BaseModel, ConfigDict, Field, ValidationError
from typing import List, Optional, Dict, Any


class RowConfig(BaseModel):
    mandatory: Optional[bool] = None
    Value: Optional[str] = None
    skipRowIfMissing: Optional[bool] = None
    targetColumn: str


class RowsConfig(BaseModel):
    model_config = ConfigDict(extra="allow")

    @classmethod
    def validate(cls, value: Any) -> "RowsConfig":
        if not isinstance(value, dict):
            raise ValueError(f"Value must be a dictionary, got {type(value)}")
        for key, val in value.items():
            if not isinstance(val, dict):
                raise ValueError(
                    f"Value for {key} must be a dictionary, got {type(val)}"
                )
            RowConfig(**val)  # Validate each dictionary against RowConfig
        return cls(**value)


class mal_config(BaseModel):
    startrow: Optional[int] = None
    readSheet: Optional[str] = None
    outPutRow: Optional[int] = None
    PrimarKeyColumn: Optional[str] = None
    OutSheet: Optional[str] = None
    Rows: List[RowsConfig]
