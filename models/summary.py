from dataclasses import dataclass, field
from collections import defaultdict
from typing import DefaultDict, Dict, Any

@dataclass
class AnalysisSummary:
    """
    Represents the overall analysis results.
    
    Attributes:
        pronouns: Count of pronouns found
        contractions: Count of contractions found
        terminology: Dictionary of terminology usage counts
        font_issues: Count of font-related issues
        caption_issues: Count of caption-related issues
        header_issues: Count of header-related issues
        table_style_issues: Count of table style issues
    """
    pronouns: int = 0
    contractions: int = 0
    terminology: DefaultDict[str, int] = field(default_factory=lambda: defaultdict(int))
    font_issues: int = 0
    caption_issues: int = 0
    header_issues: int = 0
    table_style_issues: int = 0

    def to_dict(self) -> Dict[str, Any]:
        """Convert the summary to a dictionary format."""
        return {
            "pronouns": self.pronouns,
            "contractions": self.contractions,
            "terminology": dict(self.terminology),
            "style_issues": {
                "font": self.font_issues,
                "captions": self.caption_issues,
                "headers": self.header_issues,
                "tables": self.table_style_issues
            }
        }