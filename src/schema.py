from __future__ import annotations
from typing import Literal, Union, Optional, Any
from pydantic import BaseModel as _BaseModel, Field, model_validator
from pydantic.json_schema import GenerateJsonSchema
from enum import Enum

class BaseModel(_BaseModel):
    model_config = {
        "extra": "forbid"
}

class Section01(BaseModel):
    title: str
    date_written: str
    date_due: str

class Personnel(BaseModel):
    name: str
    phone_number: str
    email_address: str
    active_employee: bool

class Section02(BaseModel):
    participants: list[Personnel]

class BaseStep(BaseModel):
    uuid: str
    description: str

# --- Step Variants ---
class StandardStep(BaseStep):
    step_type: Literal["standard"]

class DateTimeStep(BaseStep):
    step_type: Literal["datetime"]
    date_completed: str
    
class SubtitleStep(BaseModel):
    step_type: Literal["subtitle"]

Step = Union[StandardStep, DateTimeStep, SubtitleStep]
BaseStep.model_rebuild()

class Section03(BaseModel):
    steps: list[Step]


class Version(BaseModel):
    date: str
    author: str
    change_description: str
    version: str

class Section04(BaseModel):
    version_history: list[Version] 

class Contents(BaseModel):
    section_01: Optional[Section01] = None
    section_02: Optional[Section02] = None
    section_03: Optional[Section03] = None
    section_04: Optional[Section04] = None
    section_05: Optional[Section05] = None

class SchemaGenerator(GenerateJsonSchema):
    def generate(self, schema, mode="validation"):
        json_schema = super().generate(schema, mode=mode)
        if "title" in json_schema:
            del json_schema["title"]
        return json_schema

    def get_schema_from_definitions(self, json_ref):
        json_schema = super().get_schema_from_definitions(json_ref)
        if json_schema and "title" in json_schema:
            del json_schema["title"]
        return json_schema

    def field_title_should_be_set(self, schema) -> bool:
        return False

if __name__ == "__main__":
    import json
    print(json.dumps(Contents.model_json_schema(schema_generator=SchemaGenerator), indent=2))