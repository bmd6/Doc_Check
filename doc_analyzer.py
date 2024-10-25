from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Set, Optional, Tuple
import docx
import re
import json
import csv
import logging
from enum import Enum
import sys
from itertools import tee

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
    term: str  # Changed from 'word' to 'term' to better reflect multi-word support
    page: int
    section: Optional[str]
    context: str
    normalized_term: str  # Added to store the normalized version

class DocumentAnalyzer:
    def __init__(self, config_path: str = "config.json"):
        """
        Initialize the document analyzer with configuration.
        
        Args:
            config_path: Path to JSON configuration file
        """
        self.config = self._load_config(config_path)
        self._prepare_terminology()
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

    def _prepare_terminology(self):
        """Prepare terminology for efficient matching."""
        self.term_groups = {}
        self.max_term_words = 1  # Track longest term for window size
        
        for group_name, terms in self.config["terminology_groups"].items():
            normalized_terms = {}
            for term in terms:
                # Generate variations of the term
                variations = self._generate_term_variations(term)
                for variation in variations:
                    normalized = self._normalize_term(variation)
                    normalized_terms[normalized] = {
                        'original': term,
                        'group': group_name
                    }
                    self.max_term_words = max(
                        self.max_term_words,
                        len(normalized.split())
                    )
            self.term_groups[group_name] = normalized_terms

    def _generate_term_variations(self, term: str) -> Set[str]:
        """Generate common variations of a term."""
        variations = {term}
        
        # Handle common separators and their variations
        separators = [' ', '-', '_', '&', 'and']
        words = re.split(r'[-_\s&]+', term)
        
        if len(words) > 1:
            # Generate variations with different separators
            for sep in separators:
                variations.add(sep.join(words))
            
            # Handle special case for '&' and 'and'
            if '&' in term:
                variations.add(term.replace('&', 'and'))
            if 'and' in term:
                variations.add(term.replace('and', '&'))
                
            # Handle optional words in parentheses
            # Example: "integration (and) test" -> ["integration test", "integration and test"]
            term_without_optional = re.sub(r'\s*\([^)]+\)\s*', ' ', term)
            if term_without_optional != term:
                variations.add(term_without_optional.strip())
                
        return variations

    def _normalize_term(self, term: str) -> str:
        """Normalize a term for consistent matching."""
        # Convert to lowercase and replace multiple spaces with single space
        normalized = ' '.join(term.lower().split())
        # Replace various separators with space
        normalized = re.sub(r'[-_&]', ' ', normalized)
        # Replace 'and' with space
        normalized = normalized.replace(' and ', ' ')
        return normalized

    def _sliding_window(self, sequence: List[str], window_size: int):
        """Create a sliding window iterator over a sequence."""
        iterators = tee(sequence, window_size)
        for i, iterator in enumerate(iterators):
            for _ in range(i):
                next(iterator, None)
        return zip(*iterators)

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
        page_count = 1
        word_count = 0
        words_per_page = 500

        for paragraph in doc.paragraphs:
            if paragraph.style.name.startswith('Heading'):
                current_section = paragraph.text

            # Split text into words while preserving original format
            original_words = paragraph.text.split()
            word_count += len(original_words)
            
            if word_count >= words_per_page:
                page_count += 1
                word_count = 0

            # Analyze for different types of issues
            issues.extend(self._analyze_pronouns(original_words, page_count, current_section))
            issues.extend(self._analyze_contractions(paragraph.text, page_count, current_section))
            issues.extend(self._analyze_terminology(paragraph.text, original_words, page_count, current_section))

        return issues

    def _analyze_pronouns(self, words: List[str], page: int, section: Optional[str]) -> List[Issue]:
        """Identify pronouns in text."""
        issues = []
        for word in words:
            if word.lower() in self.pronouns:
                issues.append(Issue(
                    type="pronoun",
                    term=word,
                    page=page,
                    section=section,
                    context=self._get_context(words, word),
                    normalized_term=word.lower()
                ))
        return issues

    def _analyze_contractions(self, text: str, page: int, section: Optional[str]) -> List[Issue]:
        """Identify contractions in text."""
        issues = []
        for contraction in self.contractions:
            if contraction in text.lower():
                issues.append(Issue(
                    type="contraction",
                    term=contraction,
                    page=page,
                    section=section,
                    context=text,
                    normalized_term=contraction.lower()
                ))
        return issues

    def _analyze_terminology(self, full_text: str, words: List[str], page: int, section: Optional[str]) -> List[Issue]:
        """Identify inconsistent terminology."""
        issues = []
        found_terms = {}  # Group -> Set of found terms

        # Look for terms of different lengths up to max_term_words
        for window_size in range(1, self.max_term_words + 1):
            for window in self._sliding_window(words, window_size):
                potential_term = ' '.join(window)
                normalized_term = self._normalize_term(potential_term)
                
                # Check each terminology group
                for group_name, terms in self.term_groups.items():
                    if normalized_term in terms:
                        if group_name not in found_terms:
                            found_terms[group_name] = set()
                        found_terms[group_name].add((
                            potential_term,
                            normalized_term,
                            terms[normalized_term]['original']
                        ))

        # Create issues for groups with multiple terms
        for group_name, terms in found_terms.items():
            if len(terms) > 1:
                for term, normalized_term, original in terms:
                    issues.append(Issue(
                        type="terminology",
                        term=term,
                        page=page,
                        section=section,
                        context=self._get_context(words, term),
                        normalized_term=normalized_term
                    ))

        return issues

    def _get_context(self, words: List[str], target: str, context_size: int = 5) -> str:
        """Get surrounding context for a term."""
        if not self.config.get("include_context", True):
            return ""
            
        try:
            # Handle multi-word terms
            target_words = target.split()
            for i in range(len(words) - len(target_words) + 1):
                if words[i:i+len(target_words)] == target_words:
                    start = max(0, i - context_size)
                    end = min(len(words), i + len(target_words) + context_size)
                    return " ".join(words[start:end])
            return ""
        except ValueError:
            return ""

    # [Previous save_results methods remain unchanged]
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
                f.write(f"Term: {issue.term}\n")
                f.write(f"Normalized Form: {issue.normalized_term}\n")
                f.write(f"Page: {issue.page}\n")
                if issue.section:
                    f.write(f"Section: {issue.section}\n")
                if issue.context:
                    f.write(f"Context: {issue.context}\n")
                f.write("\n")

    def _save_csv(self, issues: List[Issue], output_path: str):
        with open(output_path, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(["Type", "Term", "Normalized Form", "Page", "Section", "Context"])
            for issue in issues:
                writer.writerow([
                    issue.type,
                    issue.term,
                    issue.normalized_term,
                    issue.page,
                    issue.section or "",
                    issue.context
                ])

    def _save_org(self, issues: List[Issue], output_path: str):
        with open(output_path, 'w') as f:
            f.write("* Document Analysis Results\n")
            for issue in issues:
                f.write(f"** {issue.type.capitalize()}\n")
                f.write(f"- Term: {issue.term}\n")
                f.write(f"- Normalized Form: {issue.normalized_term}\n")
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
                f.write(f"- Term: {issue.term}\n")
                f.write(f"- Normalized Form: {issue.normalized_term}\n")
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
        print(f"Found {issue.type}: {issue.term}")
        print(f"Normalized Form: {issue.normalized_term}")
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
