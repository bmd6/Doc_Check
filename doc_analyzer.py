from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Set, Optional
import docx
import re
import json
import csv
import logging
from enum import Enum
import sys

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class OutputFormat(Enum):
    TXT = "txt"
    CSV = "csv"
    ORG = "org"
    MD = "md"

@dataclass
class Issue:
    type: str
    word: str
    page: int
    section: Optional[str]
    context: str

class DocumentAnalyzer:
    def __init__(self, config_path: str = "config.json"):
        """
        Initialize the document analyzer with configuration.
        
        Args:
            config_path: Path to JSON configuration file
        """
        self.config = self._load_config(config_path)
        self.pronouns = set([
            "he", "him", "his", "she", "her", "hers", "it", "its",
            "they", "them", "their", "theirs", "we", "us", "our", "ours",
            "i", "me", "my", "mine", "you", "your", "yours"
        ])
        self.contractions = set([
            "n't", "'ll", "'re", "'ve", "'m", "'d", "'s"
        ])
        
    def _load_config(self, config_path: str) -> dict:
        """Load configuration from JSON file."""
        try:
            with open(config_path, 'r') as f:
                return json.load(f)
        except FileNotFoundError:
            logger.warning(f"Config file not found at {config_path}. Using default settings.")
            return {
                "terminology_groups": {},
                "output_format": ["txt"],
                "include_context": True
            }
        except json.JSONDecodeError as e:
            logger.error(f"Error parsing config file: {e}")
            sys.exit(1)

    def analyze_document(self, doc_path: str) -> List[Issue]:
        """
        Analyze document for issues.
        
        Args:
            doc_path: Path to the document
            
        Returns:
            List of Issue objects
        """
        try:
            doc = docx.Document(doc_path)
        except Exception as e:
            logger.error(f"Error opening document: {e}")
            return []

        issues = []
        current_section = None
        page_count = 1  # Approximate page counting
        word_count = 0
        words_per_page = 500  # Approximate words per page

        for paragraph in doc.paragraphs:
            # Check if paragraph is a section header
            if paragraph.style.name.startswith('Heading'):
                current_section = paragraph.text

            words = paragraph.text.lower().split()
            word_count += len(words)
            
            # Update page count
            if word_count >= words_per_page:
                page_count += 1
                word_count = 0

            # Analyze paragraph
            issues.extend(self._analyze_pronouns(words, page_count, current_section))
            issues.extend(self._analyze_contractions(paragraph.text, page_count, current_section))
            issues.extend(self._analyze_terminology(words, page_count, current_section))

        return issues

    def _analyze_pronouns(self, words: List[str], page: int, section: Optional[str]) -> List[Issue]:
        """Identify pronouns in text."""
        issues = []
        for word in words:
            if word.lower() in self.pronouns:
                issues.append(Issue(
                    type="pronoun",
                    word=word,
                    page=page,
                    section=section,
                    context=self._get_context(words, word)
                ))
        return issues

    def _analyze_contractions(self, text: str, page: int, section: Optional[str]) -> List[Issue]:
        """Identify contractions in text."""
        issues = []
        for contraction in self.contractions:
            if contraction in text.lower():
                issues.append(Issue(
                    type="contraction",
                    word=contraction,
                    page=page,
                    section=section,
                    context=text
                ))
        return issues

    def _analyze_terminology(self, words: List[str], page: int, section: Optional[str]) -> List[Issue]:
        """Identify inconsistent terminology."""
        issues = []
        for term_group in self.config["terminology_groups"].values():
            found_terms = set(word.lower() for word in words if word.lower() in term_group)
            if len(found_terms) > 1:
                for term in found_terms:
                    issues.append(Issue(
                        type="terminology",
                        word=term,
                        page=page,
                        section=section,
                        context=self._get_context(words, term)
                    ))
        return issues

    def _get_context(self, words: List[str], target: str, context_size: int = 5) -> str:
        """Get surrounding context for a word."""
        if not self.config.get("include_context", True):
            return ""
            
        try:
            idx = words.index(target)
            start = max(0, idx - context_size)
            end = min(len(words), idx + context_size + 1)
            return " ".join(words[start:end])
        except ValueError:
            return ""

    def save_results(self, issues: List[Issue], output_path: str, format: OutputFormat):
        """Save analysis results to file."""
        try:
            if format == OutputFormat.TXT:
                self._save_txt(issues, output_path)
            elif format == OutputFormat.CSV:
                self._save_csv(issues, output_path)
            elif format == OutputFormat.ORG:
                self._save_org(issues, output_path)
            elif format == OutputFormat.MD:
                self._save_md(issues, output_path)
        except Exception as e:
            logger.error(f"Error saving results: {e}")

    def _save_txt(self, issues: List[Issue], output_path: str):
        with open(output_path, 'w') as f:
            for issue in issues:
                f.write(f"Type: {issue.type}\n")
                f.write(f"Word: {issue.word}\n")
                f.write(f"Page: {issue.page}\n")
                if issue.section:
                    f.write(f"Section: {issue.section}\n")
                if issue.context:
                    f.write(f"Context: {issue.context}\n")
                f.write("\n")

    def _save_csv(self, issues: List[Issue], output_path: str):
        with open(output_path, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(["Type", "Word", "Page", "Section", "Context"])
            for issue in issues:
                writer.writerow([
                    issue.type,
                    issue.word,
                    issue.page,
                    issue.section or "",
                    issue.context
                ])

    def _save_org(self, issues: List[Issue], output_path: str):
        with open(output_path, 'w') as f:
            f.write("* Document Analysis Results\n")
            for issue in issues:
                f.write(f"** {issue.type.capitalize()}\n")
                f.write(f"- Word: {issue.word}\n")
                f.write(f"- Page: {issue.page}\n")
                if issue.section:
                    f.write(f"- Section: {issue.section}\n")
                if issue.context:
                    f.write(f"- Context: {issue.context}\n")
                f.write("\n")

    def _save_md(self, issues: List[Issue], output_path: str):
        with open(output_path, 'w') as f:
            f.write("# Document Analysis Results\n\n")
            for issue in issues:
                f.write(f"## {issue.type.capitalize()}\n")
                f.write(f"- Word: {issue.word}\n")
                f.write(f"- Page: {issue.page}\n")
                if issue.section:
                    f.write(f"- Section: {issue.section}\n")
                if issue.context:
                    f.write(f"- Context: {issue.context}\n")
                f.write("\n")

def main():
    """Main function to run the document analyzer."""
    if len(sys.argv) < 2:
        logger.error("Please provide a document path")
        sys.exit(1)

    doc_path = sys.argv[1]
    config_path = sys.argv[2] if len(sys.argv) > 2 else "config.json"

    analyzer = DocumentAnalyzer(config_path)
    issues = analyzer.analyze_document(doc_path)

    # Print results to terminal
    for issue in issues:
        print(f"Found {issue.type}: {issue.word}")
        print(f"Page: {issue.page}")
        if issue.section:
            print(f"Section: {issue.section}")
        if issue.context:
            print(f"Context: {issue.context}")
        print()

    # Save results in configured formats
    for format_str in analyzer.config.get("output_format", ["txt"]):
        try:
            format_enum = OutputFormat(format_str.lower())
            output_path = f"analysis_results.{format_str}"
            analyzer.save_results(issues, output_path, format_enum)
            logger.info(f"Results saved to {output_path}")
        except ValueError:
            logger.warning(f"Unsupported output format: {format_str}")

if __name__ == "__main__":
    main()
