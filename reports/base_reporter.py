# reports/base_reporter.py
from abc import ABC, abstractmethod
from typing import List
from models.issues import Issue, StyleIssue
from models.fonts import FontUsageSummary
from models.summary import AnalysisSummary
from analyzers.acronym_analyzer import AcronymAnalyzer

class BaseReporter(ABC):
    """Base class for all report generators."""
    
    def __init__(self, summary: AnalysisSummary, font_usage: FontUsageSummary, 
                 acronym_analyzer: AcronymAnalyzer):
        """
        Initialize reporter with analysis results.
        
        Args:
            summary: Overall analysis summary
            font_usage: Font usage statistics
            acronym_analyzer: Acronym analysis results
        """
        self.summary = summary
        self.font_usage = font_usage
        self.acronym_analyzer = acronym_analyzer  # Updated attribute name

    @abstractmethod
    def generate_report(self, content_issues: List[Issue], 
                       style_issues: List[StyleIssue], 
                       output_path: str) -> None:
        """
        Generate and save the report.
        
        Args:
            content_issues: List of content-related issues
            style_issues: List of style-related issues
            output_path: Path where report should be saved
        """
        pass