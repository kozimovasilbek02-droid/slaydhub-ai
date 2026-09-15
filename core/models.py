from typing import List, Optional, Literal
from pydantic import BaseModel, Field

class ColorRGB(BaseModel):
    r: int = Field(..., ge=0, le=255)
    g: int = Field(..., ge=0, le=255)
    b: int = Field(..., ge=0, le=255)

    @classmethod
    def from_hex(cls, hex_str: str) -> 'ColorRGB':
        hex_clean = hex_str.lstrip('#')
        if len(hex_clean) == 3:
            hex_clean = ''.join([c * 2 for c in hex_clean])
        if len(hex_clean) != 6:
            return cls(r=50, g=50, b=50)
        try:
            r = int(hex_clean[0:2], 16)
            g = int(hex_clean[2:4], 16)
            b = int(hex_clean[4:6], 16)
            return cls(r=r, g=g, b=b)
        except ValueError:
            return cls(r=50, g=50, b=50)

    def to_hex(self) -> str:
        return f'#{self.r:02x}{self.g:02x}{self.b:02x}'

class Position(BaseModel):
    left_pct: float
    top_pct: float
    width_pct: float
    height_pct: float

class TextRun(BaseModel):
    text: str
    font_size_pt: float = 14.0
    font_bold: bool = False
    font_italic: bool = False
    font_color: Optional[str] = '#333333'
    font_name: str = 'Segoe UI'

class TextBlock(BaseModel):
    runs: List[TextRun] = Field(default_factory=list)
    alignment: Literal['LEFT', 'CENTER', 'RIGHT', 'JUSTIFY'] = 'LEFT'
    bullet: bool = False

class TextBoxElement(BaseModel):
    id: str
    position: Position
    paragraphs: List[TextBlock] = Field(default_factory=list)
    background_color: Optional[str] = None
    border_color: Optional[str] = None
    border_width_pt: float = 0.0

class ShapeElement(BaseModel):
    id: str
    shape_type: Literal[
        'RECTANGLE', 'ROUNDED_RECTANGLE', 'OVAL', 'CIRCLE', 
        'CHEVRON', 'RIGHT_ARROW', 'LINE', 'DONUT', 'STAR', 'BADGE', 'CONNECTOR'
    ] = 'RECTANGLE'
    position: Position
    fill_color: Optional[str] = '#3B82F6'
    fill_opacity: float = 1.0
    line_color: Optional[str] = None
    line_width_pt: float = 0.0
    line_dashed: bool = False
    shadow: bool = False
    text: Optional[str] = None
    text_color: Optional[str] = '#FFFFFF'
    text_size_pt: float = 12.0
    text_bold: bool = False

class MilestoneNode(BaseModel):
    index: int
    title: str
    description: str
    step_label: str = '20XX'
    accent_color: str = '#3B82F6'
    position_type: Literal['TOP', 'BOTTOM', 'CENTER'] = 'CENTER'
    icon_type: Optional[str] = None

class SlideData(BaseModel):
    slide_index: int = 1
    title: Optional[str] = None
    subtitle: Optional[str] = None
    background_color: str = '#FFFFFF'
    theme_colors: List[str] = Field(default_factory=lambda: ['#8A56AC', '#00838F', '#EB4D55', '#FA8900', '#0288D1'])
    shapes: List[ShapeElement] = Field(default_factory=list)
    text_boxes: List[TextBoxElement] = Field(default_factory=list)
    milestones: List[MilestoneNode] = Field(default_factory=list)
    layout_type: Literal[
        'ROADMAP_TIMELINE', 'PROCESS_CHEVRON', 'FEATURE_CARDS', 
        'STATISTICS_GRID', 'GENERAL_INFOGRAPHIC', 'TITLE_CONTENT'
    ] = 'ROADMAP_TIMELINE'

class PresentationProject(BaseModel):
    presentation_title: str = 'Converted Presentation'
    aspect_ratio: Literal['16:9', '4:3'] = '16:9'
    slides: List[SlideData] = Field(default_factory=list)
