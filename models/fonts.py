from dataclasses import dataclass, field
from collections import defaultdict
from typing import Optional, DefaultDict

@dataclass
class FontInfo:
    """
    Represents font formatting information.
    
    Attributes:
        name: Font family name
        size: Font size in points
        bold: Whether the font is bold
        italic: Whether the font is italic
    """
    name: Optional[str] = None
    size: Optional[float] = None
    bold: bool = False
    italic: bool = False

    def __post_init__(self):
        """Convert Word's half-points to points and validate size."""
        if self.size is not None:
            self.size = self.size / 2
            if self.size > 100:  # Unreasonable size check
                self.size = None

    def __str__(self) -> str:
        """Return a human-readable string representation of the font info."""
        attributes = []
        if self.name:
            attributes.append(f"Font: {self.name}")
        if self.size is not None:
            attributes.append(f"Size: {self.size:.1f}pt")
        if self.bold:
            attributes.append("Bold")
        if self.italic:
            attributes.append("Italic")
        return ", ".join(attributes) if attributes else "Default"

@dataclass
class FontUsageSummary:
    """
    Tracks font usage throughout the document.
    
    Attributes:
        body_fonts: Font usage in body text
        header_fonts: Font usage in headers by level
        caption_fonts: Font usage in captions
        table_fonts: Font usage in tables
    """
    body_fonts: DefaultDict[str, int] = field(default_factory=lambda: defaultdict(int))
    header_fonts: DefaultDict[int, DefaultDict[str, int]] = field(
        default_factory=lambda: defaultdict(lambda: defaultdict(int)))
    caption_fonts: DefaultDict[str, int] = field(default_factory=lambda: defaultdict(int))
    table_fonts: DefaultDict[str, int] = field(default_factory=lambda: defaultdict(int))

    def add_body_font(self, font_info: str, char_count: int) -> None:
        """Add body text font usage."""
        self.body_fonts[font_info] += char_count

    def add_header_font(self, level: int, font_info: str, char_count: int) -> None:
        """Add header font usage for specific level."""
        self.header_fonts[level][font_info] += char_count

    def add_caption_font(self, font_info: str, char_count: int) -> None:
        """Add caption font usage."""
        self.caption_fonts[font_info] += char_count

    def add_table_font(self, font_info: str, char_count: int) -> None:
        """Add table font usage."""
        self.table_fonts[font_info] += char_count

    def get_formatted_summary(self) -> str:
        """Generate a formatted summary of font usage."""
        summary = []

        # Body text summary
        if self.body_fonts:
            summary.append("Body Text Fonts:")
            total_body_chars = sum(self.body_fonts.values())
            for font, count in sorted(self.body_fonts.items(), key=lambda x: x[1], reverse=True):
                percentage = (count / total_body_chars) * 100
                summary.append(f"  - {font}: {percentage:.1f}% ({count} characters)")

        # Header text summary by level
        if self.header_fonts:
            summary.append("\nHeader Text Fonts:")
            for level in sorted(self.header_fonts.keys()):
                summary.append(f"  Level {level}:")
                total_level_chars = sum(self.header_fonts[level].values())
                for font, count in sorted(self.header_fonts[level].items(), 
                                        key=lambda x: x[1], reverse=True):
                    percentage = (count / total_level_chars) * 100
                    summary.append(f"    - {font}: {percentage:.1f}% ({count} characters)")

        # Caption text summary
        if self.caption_fonts:
            summary.append("\nCaption Text Fonts:")
            total_caption_chars = sum(self.caption_fonts.values())
            for font, count in sorted(self.caption_fonts.items(), key=lambda x: x[1], reverse=True):
                percentage = (count / total_caption_chars) * 100
                summary.append(f"  - {font}: {percentage:.1f}% ({count} characters)")

        # Table text summary
        if self.table_fonts:
            summary.append("\nTable Text Fonts:")
            total_table_chars = sum(self.table_fonts.values())
            for font, count in sorted(self.table_fonts.items(), key=lambda x: x[1], reverse=True):
                percentage = (count / total_table_chars) * 100
                summary.append(f"  - {font}: {percentage:.1f}% ({count} characters)")

        return "\n".join(summary)