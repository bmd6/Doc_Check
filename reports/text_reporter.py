# reports/text_reporter.py
from typing import List
from .base_reporter import BaseReporter
from models.issues import Issue, StyleIssue

class TextReporter(BaseReporter):
    """Generates reports in plain text format."""
    
    def generate_report(self, content_issues: List[Issue], 
                       style_issues: List[StyleIssue], 
                       output_path: str) -> None:
        """
        Generate a plain text report.
        
        Args:
            content_issues: List of content-related issues
            style_issues: List of style-related issues
            output_path: Path where the text report should be saved
        """
        with open(output_path, 'w', encoding='utf-8') as f:
            # Document header
            f.write("Document Analysis Summary\n")
            f.write("=======================\n\n")
            
            # Font usage analysis
            f.write("Font Usage Analysis\n")
            f.write("------------------\n")
            f.write(self.font_usage.get_formatted_summary())
            f.write("\n\n")
            
            # Content Issues Summary
            f.write("Content Issues Summary\n")
            f.write("---------------------\n")
            f.write(f"- Pronouns found: {self.summary.pronouns}\n")
            f.write(f"- Contractions found: {self.summary.contractions}\n")
            f.write("- Terminology conflicts:\n")
            for term, count in self.summary.terminology.items():
                f.write(f"  - {term}: {count} occurrences\n")
            
            # Style Issues Summary
            f.write("\nStyle Issues Summary\n")
            f.write("-------------------\n")
            f.write(f"- Font inconsistencies: {self.summary.font_issues}\n")
            f.write(f"- Caption style issues: {self.summary.caption_issues}\n")
            f.write(f"- Header style issues: {self.summary.header_issues}\n")
            f.write(f"- Table style issues: {self.summary.table_style_issues}\n")
            
            # Acronyms Summary
            f.write("\nAcronyms Summary\n")
            f.write("----------------\n")
            for acronym, info in self.acronym_analyzer.found_acronyms.items():
                pages = ', '.join(map(str, sorted(info['pages'])))
                definition = info['definition'] or ""
                f.write(f"{acronym}: {definition} (Pages: {pages})\n")
            
            # Detailed Findings
            if content_issues or style_issues:
                f.write("\nDetailed Findings\n")
                f.write("=================\n")
            
            # Content Issues Details
            if content_issues:
                f.write("\nContent Issues:\n")
                for issue in content_issues:
                    f.write(f"\n{issue.type}:\n")
                    f.write(f"- Term: {issue.term}\n")
                    f.write(f"- Page: {issue.page}\n")
                    if issue.section:
                        f.write(f"- Section: {issue.section}\n")
                    if issue.context:
                        f.write(f"- Context: {issue.context}\n")
            
            # Style Issues Details
            if style_issues:
                f.write("\nStyle Issues:\n")
                for issue in style_issues:
                    f.write(f"\n{issue.type}:\n")
                    f.write(f"- Element: {issue.element}\n")
                    f.write(f"- Expected: {issue.expected}\n")
                    f.write(f"- Found: {issue.found}\n")
                    f.write(f"- Page: {issue.page}\n")
                    if issue.section:
                        f.write(f"- Section: {issue.section}\n")
                    if issue.context:
                        f.write(f"- Context: {issue.context}\n")