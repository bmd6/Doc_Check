# reports/markdown_reporter.py
from typing import List
from .base_reporter import BaseReporter
from models.issues import Issue, StyleIssue
from datetime import datetime

class MarkdownReporter(BaseReporter):
    """Generates reports in Markdown format."""
    
    def generate_report(self, content_issues: List[Issue], 
                       style_issues: List[StyleIssue], 
                       output_path: str) -> None:
        """
        Generate a Markdown format report.
        
        Args:
            content_issues: List of content-related issues
            style_issues: List of style-related issues
            output_path: Path where the Markdown report should be saved
        """
        with open(output_path, 'w', encoding='utf-8') as f:
            # Document header
            f.write("# Document Analysis Summary\n\n")
            
            # Font Usage Analysis
            f.write("## Font Usage Analysis\n")
            f.write("```\n")
            f.write(self.font_usage.get_formatted_summary())
            f.write("\n```\n\n")
            
            # Content Issues Summary
            f.write("## Content Issues Summary\n")
            f.write(f"* Pronouns found: {self.summary.pronouns}\n")
            f.write(f"* Contractions found: {self.summary.contractions}\n")
            f.write("* Terminology conflicts:\n")
            for term, count in self.summary.terminology.items():
                f.write(f"  * {term}: {count} occurrences\n")
            
            # Style Issues Summary
            f.write("\n## Style Issues Summary\n")
            f.write("| Category | Count |\n")
            f.write("|----------|-------|\n")
            f.write(f"| Font inconsistencies | {self.summary.font_issues} |\n")
            f.write(f"| Caption style issues | {self.summary.caption_issues} |\n")
            f.write(f"| Header style issues | {self.summary.header_issues} |\n")
            f.write(f"| Table style issues | {self.summary.table_style_issues} |\n")
            
            # Acronyms Summary
            f.write("\n## Acronyms Summary\n")
            f.write("| Acronym | Definition |\n")
            f.write("|---------|------------|\n")
            for acronym, info in sorted(self.acronym_analyzer.found_acronyms.items()):
                definition = info['definition'] or "Unknown"
                f.write(f"| {acronym} | {definition} |\n")
            
            # Detailed Findings
            if content_issues or style_issues:
                f.write("\n# Detailed Findings\n")
            
            # Content Issues Details
            if content_issues:
                f.write("\n## Content Issues\n")
                for issue in content_issues:
                    f.write(f"\n### {issue.type}\n")
                    f.write(f"* **Term**: {issue.term}\n")
                    f.write(f"* **Page**: {issue.page}\n")
                    if issue.section:
                        f.write(f"* **Section**: {issue.section}\n")
                    if issue.context:
                        f.write(f"* **Context**: {issue.context}\n")
            
            # Style Issues Details
            if style_issues:
                f.write("\n## Style Issues\n")
                for issue in style_issues:
                    f.write(f"\n### {issue.type}\n")
                    f.write(f"* **Element**: {issue.element}\n")
                    f.write(f"* **Expected**: {issue.expected}\n")
                    f.write(f"* **Found**: {issue.found}\n")
                    f.write(f"* **Page**: {issue.page}\n")
                    if issue.section:
                        f.write(f"* **Section**: {issue.section}\n")
                    if issue.context:
                        f.write(f"* **Context**: {issue.context}\n")
            
            # Footer
            f.write(f"\n---\n*Report generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}*\n")

    def get_timestamp(self) -> str:
        """Get formatted timestamp for the report."""
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")