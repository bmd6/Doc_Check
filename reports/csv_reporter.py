# reports/csv_reporter.py
from typing import List
import csv
from .base_reporter import BaseReporter
from models.issues import Issue, StyleIssue

class CSVReporter(BaseReporter):
    """Generates reports in CSV format."""
    
    def generate_report(self, content_issues: List[Issue], 
                       style_issues: List[StyleIssue], 
                       output_path: str) -> None:
        """
        Generate a CSV format report.
        
        Args:
            content_issues: List of content-related issues
            style_issues: List of style-related issues
            output_path: Path where the CSV report should be saved
        """
        with open(output_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            
            # Write summary section
            writer.writerow(["Document Analysis Summary"])
            writer.writerow([])
            
            # Font Usage Summary
            writer.writerow(["Font Usage Analysis"])
            for line in self.font_usage.get_formatted_summary().split('\n'):
                writer.writerow([line])
            writer.writerow([])
            
            # Content Issues Summary
            writer.writerow(["Content Issues Summary"])
            writer.writerow(["Pronouns found", self.summary.pronouns])
            writer.writerow(["Contractions found", self.summary.contractions])
            writer.writerow([])
            
            # Terminology conflicts
            writer.writerow(["Terminology conflicts"])
            writer.writerow(["Term", "Occurrences"])
            for term, count in self.summary.terminology.items():
                writer.writerow([term, count])
            
            # Style Issues Summary
            writer.writerow([])
            writer.writerow(["Style Issues Summary"])
            writer.writerow(["Category", "Count"])
            writer.writerow(["Font inconsistencies", self.summary.font_issues])
            writer.writerow(["Caption style issues", self.summary.caption_issues])
            writer.writerow(["Header style issues", self.summary.header_issues])
            writer.writerow(["Table style issues", self.summary.table_style_issues])
            
            # Acronyms Summary
            writer.writerow([])
            writer.writerow(["Acronyms Summary"])
            writer.writerow(["Acronym", "Definition"])
            for acronym, info in sorted(self.acronym_analyzer.found_acronyms.items()):
                writer.writerow([acronym, info['definition'] or ""])
            
            # Detailed Findings
            if content_issues:
                writer.writerow([])
                writer.writerow(["Content Issues"])
                writer.writerow(["Type", "Term", "Page", "Section", "Context"])
                for issue in content_issues:
                    writer.writerow([
                        issue.type,
                        issue.term,
                        issue.page,
                        issue.section or "",
                        issue.context or ""
                    ])
            
            if style_issues:
                writer.writerow([])
                writer.writerow(["Style Issues"])
                writer.writerow(["Type", "Element", "Expected", "Found", 
                               "Page", "Section", "Context"])
                for issue in style_issues:
                    writer.writerow([
                        issue.type,
                        issue.element,
                        issue.expected,
                        issue.found,
                        issue.page,
                        issue.section or "",
                        issue.context or ""
                    ])