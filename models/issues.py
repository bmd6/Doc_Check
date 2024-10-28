from dataclasses import dataclass
from typing import Optional

@dataclass
class Issue:
    """
    Base class for document analysis issues.
    
    Attributes:
        type: Type of the issue (e.g., 'Pronoun', 'Contraction')
        term: The specific term that caused the issue
        page: Page number where the issue was found
        section: Document section where the issue was found
        context: Surrounding text context of the issue
        normalized_term: Normalized version of the term for comparison
    """
    type: str
    term: str
    page: int
    section: Optional[str]
    context: str
    normalized_term: str

@dataclass
class StyleIssue:
    """
    Represents style-related issues found in the document.
    
    Attributes:
        type: Type of style issue (e.g., 'Font', 'Spacing')
        element: Element where the issue was found
        expected: Expected style
        found: Actual style found
        page: Page number where the issue was found
        section: Document section where the issue was found
        context: Surrounding context of the issue
    """
    type: str
    element: str
    expected: str
    found: str
    page: int
    section: Optional[str]
    context: str
