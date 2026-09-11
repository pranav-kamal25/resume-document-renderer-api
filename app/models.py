from typing import List, Literal, Optional

from pydantic import BaseModel, Field


RenderStatus = Literal[
    "PASS",
    "UNDERFILL",
    "OVERDENSE",
    "OVERFLOW",
    "FAIL",
]

FitDirection = Literal[
    "NONE",
    "ADD_CONTENT",
    "REMOVE_CONTENT",
    "FIX_TECHNICAL",
]


class Header(BaseModel):
    name: str
    location: str
    email: str
    phone: str
    linkedin: str


class Education(BaseModel):
    school_line: str = Field(
        description=(
            "Full bold education line, e.g. school | degree | expected graduation"
        )
    )
    gpa_text: str = Field(
        description="Text after 'GPA: '"
    )
    coursework: List[str] = Field(
        default_factory=list
    )


class Experience(BaseModel):
    heading: str = Field(
        description="Copy-exact title | employer | dates heading"
    )
    bullets: List[str] = Field(
        default_factory=list
    )


class Project(BaseModel):
    heading: str
    bullets: List[str] = Field(
        default_factory=list
    )


class SkillGroup(BaseModel):
    label: str = Field(
        description=(
            "Label without trailing colon, e.g. Programming & Data"
        )
    )
    items: List[str] = Field(
        default_factory=list
    )


class ResumeContent(BaseModel):
    header: Header
    education: Education
    experiences: List[Experience]
    projects: List[Project] = Field(
        default_factory=list
    )
    certifications: List[str] = Field(
        default_factory=list
    )
    skills: List[SkillGroup] = Field(
        default_factory=list
    )


class RenderRequest(BaseModel):
    filename_stem: str = Field(
        default="resume",
        min_length=1,
        max_length=140,
    )
    content: ResumeContent
    include_docx: bool = True


class RenderQA(BaseModel):
    status: RenderStatus
    page_count: int
    max_text_y_pt: Optional[float] = None
    golden_max_text_y_pt: float
    bottom_delta_from_golden_pt: Optional[float] = None
    fit_direction: FitDirection = "NONE"
    warnings: List[str] = Field(
        default_factory=list
    )


class RenderResponse(BaseModel):
    render_id: str
    status: RenderStatus
    pdf_url: str
    docx_url: Optional[str] = None
    qa: RenderQA
    expires_in_minutes: int
