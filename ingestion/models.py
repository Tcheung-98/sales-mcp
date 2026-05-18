from __future__ import annotations
from typing import Literal
from pydantic import BaseModel, PositiveInt

class Tags(BaseModel):
    industry: str
    sub_industry: str = ""
    product_line: str = ""
    deal_size: str = ""
    client_name: str = ""
    date: str = ""
    deck_type: str = ""
    status: str = ""

class TagSources(BaseModel):
    industry: Literal["folder_path", "metadata", "manual"] | None = None
    sub_industry: Literal["folder_path", "metadata", "manual"] | None = None
    product_line: Literal["folder_path", "metadata", "manual"] | None = None
    deal_size: Literal["folder_path", "metadata", "manual"] | None = None
    client_name: Literal["folder_path", "metadata", "manual"] | None = None
    date: Literal["folder_path", "metadata", "manual"] | None = None
    deck_type: Literal["folder_path", "metadata", "manual"] | None = None
    status: Literal["folder_path", "metadata", "manual"] | None = None

class SlideRow(BaseModel):
    deck_id: str
    source_path: str
    content_hash: str
    ingested_at: str
    slide_number: PositiveInt
    layout_name: str | None = None
    title: str | None = None
    body_text: list[str]
    speaker_notes: str
    tags: Tags
    tag_sources: TagSources | None = None

class FailedRecord(BaseModel):
    deck_id: str
    source_path: str
    error: str
    failed_at: str
    layer: Literal["bronze", "silver", "gold"] = "bronze"