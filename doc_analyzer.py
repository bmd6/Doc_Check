import sys
import json
import csv
import re
import logging
import time
from enum import Enum
from pathlib import Path
from collections import defaultdict
from dataclasses import dataclass, field
from typing import List, Set, Optional, Tuple, DefaultDict, Iterator
from datetime import datetime
from itertools import tee

import docx
from docx.oxml.shared import qn
from docx.oxml.ns import qn
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph

from acronym_finder import AcronymFinder  # Ensure this module is available


# =============================
# Logging Configuration
# =============================

# Ensure logs directory exists
logs_dir = Path("logs")
logs_dir.mkdir(exist_ok=True)

# Create a custom logger
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)  # Set to DEBUG to capture all levels of logs

# Create handlers
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)  # Set higher level for console

log_file_path = logs_dir / 'document_analyzer.log'
file_handler = logging.FileHandler(log_file_path, mode='a')
file_handler.setLevel(logging.DEBUG)  # Capture all logs in file

# Create formatters and add them to handlers
formatter = logging.Formatter(
    '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
console_handler.setFormatter(formatter)
file_handler.setFormatter(formatter)

# Add handlers to the logger
logger.addHandler(console_handler)
logger.addHandler(file_handler)


# =============================
# Enums and Data Classes
# =============================

class OutputFormat(Enum):
    TXT = "txt"
    CSV = "csv"
    ORG = "org"
    MD = "md"


@dataclass
class FontInfo:
    name: Optional[str] = None
    size: Optional[float] = None
    bold: bool = False
    italic: bool = False

    def __post_init__(self):
        # Convert half-points to points and handle common font sizes
        if self.size is not None:
            # Word stores font sizes in half-points, so divide by 2
            self.size = self.size / 2
            # Common font sizes are typically between 8 and 72 points
            if self.size > 100:  # If size is unreasonably large
                logger.debug(f"Font size {self.size}pt is unreasonably large. Marking as unknown.")
                self.size = None  # Mark as unknown/default

    def __str__(self):
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

    def __eq__(self, other):
        if other is None:
            return False
        # Compare only non-None attributes
        if self.name and other.name and self.name != other.name:
            return False
        if self.size and other.size and abs(self.size - other.size) > 0.1:  # Allow small float differences
            return False
        if self.bold != other.bold:
            return False
        if self.italic != other.italic:
            return False
        return True


@dataclass
class Issue:
    type: str
    term: str
    page: int
    section: Optional[str]
    context: str
    normalized_term: str


@dataclass
class StyleIssue:
    type: str
    element: str
    expected: str
    found: str
    page: int
    section: Optional[str]
    context: str


@dataclass
class FontUsageSummary:
    body_fonts: DefaultDict[str, int] = field(default_factory=lambda: defaultdict(int))
    header_fonts: DefaultDict[int, DefaultDict[str, int]] = field(
        default_factory=lambda: defaultdict(lambda: defaultdict(int)))
    caption_fonts: DefaultDict[str, int] = field(default_factory=lambda: defaultdict(int))
    table_fonts: DefaultDict[str, int] = field(default_factory=lambda: defaultdict(int))

    def add_body_font(self, font_info: str, char_count: int):
        self.body_fonts[font_info] += char_count

    def add_header_font(self, level: int, font_info: str, char_count: int):
        self.header_fonts[level][font_info] += char_count

    def add_caption_font(self, font_info: str, char_count: int):
        self.caption_fonts[font_info] += char_count

    def add_table_font(self, font_info: str, char_count: int):
        self.table_fonts[font_info] += char_count

    def get_formatted_summary(self) -> str:
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
                for font, count in sorted(self.header_fonts[level].items(), key=lambda x: x[1], reverse=True):
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


@dataclass
class AnalysisSummary:
    pronouns: int = 0
    contractions: int = 0
    terminology: DefaultDict[str, int] = field(default_factory=lambda: defaultdict(int))
    font_issues: int = 0
    caption_issues: int = 0
    header_issues: int = 0
    table_style_issues: int = 0

    def to_dict(self):
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


# =============================
# Helper Classes and Functions
# =============================

class ProgressTracker:
    """Helper class to track and display progress"""

    def __init__(self, total_steps: int, description: str):
        self.total = total_steps
        self.current = 0
        self.description = description
        self.start_time = time.time()
        self._print_progress()

    def update(self, steps: int = 1):
        self.current += steps
        self._print_progress()

    def _print_progress(self):
        percentage = (self.current / self.total) * 100 if self.total > 0 else 0
        elapsed_time = time.time() - self.start_time
        sys.stdout.write(f"\r{self.description}: [{self.current}/{self.total}] {percentage:.1f}% "
                         f"(Elapsed: {elapsed_time:.1f}s)")
        sys.stdout.flush()

    def complete(self):
        self.current = self.total
        self._print_progress()
        print()  # New line after completion


class PageCounter:
    """Helper class to handle page counting in Word documents"""

    def __init__(self, document):
        logger.info("Analyzing document structure...")
        self.doc = document
        self.current_page = 1
        self._initialize_page_breaks()
        logger.info("Document structure analysis complete.")

    def _initialize_page_breaks(self):
        """Initialize page break tracking"""
        self.explicit_breaks = set()
        current_pos = 0

        total_paragraphs = len(self.doc.paragraphs)
        progress = ProgressTracker(total_paragraphs, "Analyzing page breaks")

        for paragraph in self.doc.paragraphs:
            for run in paragraph.runs:
                current_pos += len(run.text)
                if '\f' in run.text or self._has_page_break(paragraph):
                    self.explicit_breaks.add(current_pos)
            progress.update()

        progress.complete()

        if not self.explicit_breaks:
            self.chars_per_page = 3500  # Default approximation
            logger.debug("No explicit page breaks found. Using default chars_per_page = 3500.")
        else:
            breaks_list = sorted(self.explicit_breaks)
            if len(breaks_list) > 1:
                page_lengths = [j - i for i, j in zip(breaks_list[:-1], breaks_list[1:])]
                self.chars_per_page = sum(page_lengths) / len(page_lengths)
                logger.debug(f"Calculated average chars_per_page = {self.chars_per_page}.")
            else:
                self.chars_per_page = breaks_list[0]
                logger.debug(f"Single page break found. chars_per_page = {self.chars_per_page}.")

    def _has_page_break(self, paragraph):
        try:
            return paragraph._p.get_or_add_pPr().pageBreakBefore_val is True
        except AttributeError:
            return False

    def get_page_number(self, current_pos):
        if not self.explicit_breaks:
            page = max(1, int(current_pos / self.chars_per_page) + 1)
            logger.debug(f"Calculated page number {page} based on chars_per_page.")
            return page

        page = 1
        for break_pos in sorted(self.explicit_breaks):
            if current_pos > break_pos:
                page += 1
            else:
                break
        logger.debug(f"Calculated page number {page} based on explicit breaks.")
        return page


def iter_block_items(parent) -> Iterator:
    """
    Yield each paragraph and table child within parent, in document order.
    Each returned value is an instance of either Paragraph or Table.
    """
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl
    from docx.text.paragraph import Paragraph
    from docx.table import Table

    for child in parent.element.body:
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


# =============================
# Main Analyzer Class
# =============================

@dataclass
class DocumentAnalyzer:
    config_path: str = "config.json"
    acronyms_csv: Optional[str] = None

    def __post_init__(self):
        """Initialize the document analyzer with configuration."""
        logger.info("Initializing document analyzer...")
        self.acronym_csv = self.acronyms_csv
        logger.info(f"Loading configuration from {self.config_path}")
        self.config = self._load_config(self.config_path)
        self._validate_config()
        logger.info("Preparing terminology analysis...")
        self._prepare_terminology()
        self.pronouns = set([
            "he", "him", "his", "she", "her", "hers", "it", "its",
            "they", "them", "their", "theirs", "we", "us", "our", "ours",
            "i", "me", "my", "mine", "you", "your", "yours"
        ])
        self.contractions = set([
            "n't", "'ll", "'re", "'ve", "'m", "'d", "'s"
        ])
        self.summary = AnalysisSummary()
        self.font_usage = FontUsageSummary()
        logger.info("Initialization complete.")

    def _load_config(self, config_path: str) -> dict:
        """Load configuration from JSON file."""
        config = {}
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            logger.info(f"Configuration loaded from {config_path}")
        except FileNotFoundError:
            logger.warning(f"Config file not found at {config_path}. Using default settings.")
            config = {}
        except json.JSONDecodeError as e:
            logger.error(f"Error parsing config file: {e}")
            sys.exit(1)
        return config

    def _validate_config(self):
        """Validate and set default configuration values."""
        # Default configuration
        default_config = {
            "terminology_groups": {},
            "output_format": ["txt"],
            "include_context": True
        }

        # Merge default config with loaded config
        merged_config = default_config.copy()
        if isinstance(self.config, dict):
            for key in default_config:
                if key not in self.config:
                    merged_config[key] = default_config[key]
                    logger.warning(f"Missing '{key}' in config. Using default value.")
                else:
                    merged_config[key] = self.config[key]
        else:
            logger.error("Configuration is not a valid dictionary.")
            sys.exit(1)

        # Validate types
        if not isinstance(merged_config["terminology_groups"], dict):
            logger.error("'terminology_groups' should be a dictionary.")
            sys.exit(1)
        if not isinstance(merged_config["output_format"], list):
            logger.error("'output_format' should be a list.")
            sys.exit(1)
        if not isinstance(merged_config["include_context"], bool):
            logger.error("'include_context' should be a boolean.")
            sys.exit(1)

        self.config = merged_config
        logger.info("Configuration validation complete.")

    def _prepare_terminology(self):
        """Prepare terminology for efficient matching."""
        self.term_groups = {}
        self.max_term_words = 1
        self.term_relationships = {}

        total_terms = sum(len(terms) for terms in self.config["terminology_groups"].values())
        progress = ProgressTracker(total_terms, "Preparing terminology")

        for group_name, terms in self.config["terminology_groups"].items():
            normalized_terms = {}
            all_variations = {}

            for term in terms:
                variations = self._generate_term_variations(term)
                for variation in variations:
                    normalized = self._normalize_term(variation)
                    normalized_terms[normalized] = {
                        'original': term,
                        'group': group_name
                    }
                    all_variations[normalized] = term
                    self.max_term_words = max(
                        self.max_term_words,
                        len(normalized.split())
                    )
                progress.update()

            # Establish term relationships for redundancy checks
            for term1 in all_variations:
                self.term_relationships[term1] = set()
                term1_words = set(term1.split())

                for term2 in all_variations:
                    if term1 != term2:
                        term2_words = set(term2.split())
                        if term1_words.issubset(term2_words) or term2_words.issubset(term1_words):
                            self.term_relationships[term1].add(term2)

            self.term_groups[group_name] = normalized_terms

        progress.complete()
        logger.info("Terminology preparation complete.")

    def _generate_term_variations(self, term: str) -> Set[str]:
        """Generate common variations of a term."""
        variations = {term}
        separators = [' ', '-', '_', '&', 'and']
        words = re.split(r'[-_\s&]+', term)

        if len(words) > 1:
            for sep in separators:
                variations.add(sep.join(words))

            if '&' in term:
                variations.add(term.replace('&', 'and'))
            if 'and' in term:
                variations.add(term.replace('and', '&'))

            term_without_optional = re.sub(r'\s*\([^)]+\)\s*', ' ', term)
            if term_without_optional != term:
                variations.add(term_without_optional.strip())

        return variations

    def _normalize_term(self, term: str) -> str:
        """Normalize a term for consistent matching."""
        normalized = ' '.join(term.lower().split())
        normalized = re.sub(r'[-_&]', ' ', normalized)
        normalized = normalized.replace(' and ', ' ')
        return normalized

    def _sliding_window(self, sequence: List[str], window_size: int):
        """Create a sliding window iterator over a sequence."""
        iterators = tee(sequence, window_size)
        for i, iterator in enumerate(iterators):
            for _ in range(i):
                next(iterator, None)
        return zip(*iterators)

    def _get_font_info(self, run) -> Optional[FontInfo]:
        """Extract font information from a run."""
        try:
            font = run._element.rPr.rFonts
            size_element = run._element.rPr.sz
            bold = run._element.rPr.b
            italic = run._element.rPr.i

            # Get font name
            font_name = None
            if font is not None:
                # Try different font attributes in order of preference
                font_name = (font.get(qn('w:ascii')) or
                             font.get(qn('w:hAnsi')) or
                             font.get(qn('w:cs')) or
                             font.get(qn('w:eastAsia')))

            # Get font size
            font_size = None
            if size_element is not None:
                try:
                    font_size = float(size_element.val)
                except (ValueError, TypeError):
                    font_size = None

            return FontInfo(
                name=font_name,
                size=font_size,
                bold=bold is not None,
                italic=italic is not None
            )
        except AttributeError as e:
            logger.debug(f"AttributeError in _get_font_info: {e}")
            return FontInfo()  # Return default font info instead of None

    def _get_paragraph_style_info(self, paragraph) -> str:
        """Get a string representation of paragraph style information."""
        total_runs = 0
        font_counts = {}  # Track frequency of each font configuration

        for run in paragraph.runs:
            if run.text.strip():  # Only consider runs with actual text
                total_runs += 1
                current_font = self._get_font_info(run)
                font_str = str(current_font)

                if font_str not in font_counts:
                    font_counts[font_str] = 0
                font_counts[font_str] += 1

        if not font_counts:  # No text runs found
            return "Default"

        # Find the most common font configuration
        most_common_font = max(font_counts.items(), key=lambda x: x[1])
        percentage = (most_common_font[1] / total_runs) * 100

        if percentage == 100:  # All runs use the same font
            return most_common_font[0]
        else:  # Mixed fonts
            mixed_fonts = ', '.join(f"'{font}'" for font in font_counts.keys())
            return f"Mixed ({mixed_fonts})"

    def analyze_document(self, doc_path: str) -> Tuple[List[Issue], List[StyleIssue]]:
        """Analyze document for issues."""
        logger.info(f"Starting analysis of document: {doc_path}")
        logger.info(f"Started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

        if not doc_path:
            logger.error("Document path is required.")
            raise ValueError("Document path is required.")

        try:
            logger.info("Loading document...")
            doc = docx.Document(doc_path)
            logger.info(f"Document loaded successfully. Found {len(doc.paragraphs)} paragraphs.")
        except Exception as e:
            logger.error(f"Error opening document: {e}")
            return [], []

        # Initialize AcronymFinder with the valid doc_path here
        try:
            self.acronym_finder = AcronymFinder(docx_path=doc_path, known_acronyms_csv=self.acronym_csv)
            logger.info("AcronymFinder initialized successfully.")
        except Exception as e:
            logger.error(f"Error initializing AcronymFinder: {e}")
            return [], []

        self.summary = AnalysisSummary()
        self.font_usage = FontUsageSummary()

        # Content analysis
        logger.info("Starting content analysis...")
        page_counter = PageCounter(doc)
        self.page_counter = page_counter  # Store for later use
        content_issues = []
        current_section = None
        current_pos = 0

        # Start acronym finder process
        try:
            self.acronym_finder.process_document()  # Process the entire document
            logger.info("AcronymFinder processing complete.")
        except Exception as e:
            logger.error(f"Error processing document with AcronymFinder: {e}")

        # Process found acronyms and add them to content issues
        for acronym, info in self.acronym_finder.found_acronyms.items():
            pages = sorted(info['pages'])
            issue = Issue(
                type="Acronym",
                term=acronym,
                page=pages[0] if pages else 1,  # use first page for summary
                section=None,
                context=f"Definition: {info['definition']}" if info['definition'] else "",
                normalized_term=acronym.lower()
            )
            content_issues.append(issue)
            self.summary.terminology[acronym.lower()] += 1
            logger.debug(f"Found acronym: {acronym} on page {pages}")

        progress = ProgressTracker(len(doc.paragraphs), "Analyzing paragraphs")

        for paragraph in doc.paragraphs:
            if paragraph.style.name.startswith('Heading'):
                current_section = paragraph.text

            current_pos += len(paragraph.text)
            page_number = page_counter.get_page_number(current_pos)

            original_words = paragraph.text.split()

            pronoun_issues = self._analyze_pronouns(original_words, page_number, current_section)
            contraction_issues = self._analyze_contractions(paragraph.text, page_number, current_section)
            terminology_issues = self._analyze_terminology(paragraph.text, original_words, page_number, current_section)

            self.summary.pronouns += len(pronoun_issues)
            self.summary.contractions += len(contraction_issues)
            for issue in terminology_issues:
                self.summary.terminology[issue.normalized_term] += 1

            content_issues.extend(pronoun_issues)
            content_issues.extend(contraction_issues)
            content_issues.extend(terminology_issues)

            progress.update()

        progress.complete()

        # Style analysis
        logger.info("Starting style analysis...")
        style_issues = self._analyze_styles(doc)

        logger.info(f"Analysis completed at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        return content_issues, style_issues

    def _analyze_pronouns(self, words: List[str], page: int, section: Optional[str]) -> List[Issue]:
        """Identify pronouns in text."""
        issues = []
        for word in words:
            if word.lower() in self.pronouns:
                issue = Issue(
                    type="Pronoun",
                    term=word,
                    page=page,
                    section=section,
                    context=self._get_context(words, word),
                    normalized_term=word.lower()
                )
                issues.append(issue)
                logger.debug(f"Pronoun found: {word} on page {page}")
        return issues

    def _analyze_contractions(self, text: str, page: int, section: Optional[str]) -> List[Issue]:
        """Identify contractions in text."""
        issues = []
        text_lower = text.lower()
        for contraction in self.contractions:
            if contraction in text_lower:
                issue = Issue(
                    type="Contraction",
                    term=contraction,
                    page=page,
                    section=section,
                    context=text,
                    normalized_term=contraction.lower()
                )
                issues.append(issue)
                logger.debug(f"Contraction found: {contraction} on page {page}")
        return issues

    def _analyze_terminology(self, full_text: str, words: List[str], page: int, section: Optional[str]) -> List[Issue]:
        """Identify inconsistent terminology."""
        found_terms = {}
        term_positions = {}

        for window_size in range(1, self.max_term_words + 1):
            for i, window in enumerate(self._sliding_window(words, window_size)):
                potential_term = ' '.join(window)
                normalized_term = self._normalize_term(potential_term)

                for group_name, terms in self.term_groups.items():
                    if normalized_term in terms:
                        if group_name not in found_terms:
                            found_terms[group_name] = set()
                        found_terms[group_name].add((
                            potential_term,
                            normalized_term,
                            terms[normalized_term]['original']
                        ))
                        term_positions[normalized_term] = (i, i + window_size)
                        logger.debug(f"Terminology term found: {potential_term} in group {group_name}")

        filtered_issues = []
        for group_name, terms in found_terms.items():
            if len(terms) > 1:
                sorted_terms = sorted(terms, key=lambda x: len(x[1].split()), reverse=True)
                covered_positions = set()

                for term, normalized_term, original in sorted_terms:
                    current_pos = term_positions[normalized_term]
                    is_redundant = False

                    for start, end in covered_positions:
                        if (current_pos[0] >= start and current_pos[0] < end) or \
                           (current_pos[1] > start and current_pos[1] <= end):
                            related_terms = self.term_relationships.get(normalized_term, set())
                            for other_term, _, _ in sorted_terms:
                                other_normalized = self._normalize_term(other_term)
                                if other_normalized in related_terms:
                                    is_redundant = True
                                    logger.debug(f"Redundant terminology term: {term}")
                                    break
                            if is_redundant:
                                break

                    if not is_redundant:
                        issue = Issue(
                            type="Terminology",
                            term=term,
                            page=page,
                            section=section,
                            context=self._get_context(words, term),
                            normalized_term=normalized_term
                        )
                        filtered_issues.append(issue)
                        covered_positions.add(current_pos)
                        logger.debug(f"Inconsistent terminology term added: {term} on page {page}")

        return filtered_issues

    def _get_context(self, words: List[str], target: str, context_size: int = 5) -> str:
        """Get surrounding context for a term."""
        if not self.config.get("include_context", True):
            return ""

        try:
            target_words = target.split()
            for i in range(len(words) - len(target_words) + 1):
                if words[i:i + len(target_words)] == target_words:
                    start = max(0, i - context_size)
                    end = min(len(words), i + len(target_words) + context_size)
                    context = " ".join(words[start:end])
                    logger.debug(f"Context for term '{target}': {context}")
                    return context
            return ""
        except ValueError:
            return ""

    def _analyze_styles(self, doc) -> List[StyleIssue]:
        """Analyze document styles for consistency."""
        logger.info("Analyzing document styles...")
        issues = []
        current_pos = 0

        # Use iter_block_items to process blocks in order
        blocks = list(iter_block_items(doc))
        total_blocks = len(blocks)
        progress = ProgressTracker(total_blocks, "Processing document blocks")

        previous_block = None

        for block in blocks:
            if isinstance(block, Paragraph):
                style_info = self._get_paragraph_style_info(block)
                char_count = len(block.text)

                if block.style.name.startswith('Heading'):
                    try:
                        header_level = int(block.style.name.split()[-1])
                    except (ValueError, IndexError):
                        header_level = 1  # Default to level 1 if not found
                    self.font_usage.add_header_font(header_level, style_info, char_count)
                elif "Caption" in block.style.name.lower():
                    self.font_usage.add_caption_font(style_info, char_count)
                elif not block.style.name.startswith('TOC'):
                    self.font_usage.add_body_font(style_info, char_count)

                previous_block = block

            elif isinstance(block, Table):
                # Check for table captions in the previous block
                has_caption = False
                if isinstance(previous_block, Paragraph):
                    if 'caption' in previous_block.style.name.lower():
                        has_caption = True

                if not has_caption:
                    # Check the next block for a caption
                    block_index = blocks.index(block)
                    if block_index + 1 < total_blocks:
                        next_block = blocks[block_index + 1]
                        if isinstance(next_block, Paragraph) and 'caption' in next_block.style.name.lower():
                            has_caption = True

                if not has_caption:
                    # Estimate page number based on current_pos
                    page_number = self.page_counter.get_page_number(current_pos)
                    issue = StyleIssue(
                        type="Missing Caption",
                        element="Table",
                        expected="Table should have a caption",
                        found="No caption found",
                        page=page_number,
                        section=None,
                        context=block.rows[0].cells[0].text[:50]  # Sample context
                    )
                    issues.append(issue)
                    self.summary.caption_issues += 1
                    logger.debug(f"Missing caption for table on page {page_number}")

                # Check if table headers repeat on each page
                headers_repeat = self._check_table_header_repeat(block)
                if not headers_repeat:
                    page_number = self.page_counter.get_page_number(current_pos)
                    issue = StyleIssue(
                        type="Table Header",
                        element="Table",
                        expected="Table headers should repeat on each page",
                        found="Table headers do not repeat",
                        page=page_number,
                        section=None,
                        context=block.rows[0].cells[0].text[:50]  # Sample context
                    )
                    issues.append(issue)
                    self.summary.table_style_issues += 1
                    logger.debug(f"Table headers do not repeat for table on page {page_number}")

                # Collect font usage in tables
                for row in block.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            style_info = self._get_paragraph_style_info(para)
                            char_count = len(para.text)
                            self.font_usage.add_table_font(style_info, char_count)

                previous_block = block  # Update previous_block to table

            progress.update()

        progress.complete()

        # Analyze figures
        if doc.inline_shapes:
            logger.info("Analyzing figure styles...")
            progress = ProgressTracker(len(doc.inline_shapes), "Processing figures")
            for shape in doc.inline_shapes:
                # Check for figure captions
                has_caption = False
                # Attempt to find the associated paragraph
                shape_paragraph = self._get_shape_paragraph(doc, shape)
                if shape_paragraph:
                    # Check the previous and next blocks for a caption
                    block_index = None
                    for idx, block in enumerate(blocks):
                        if isinstance(block, Paragraph) and block == shape_paragraph:
                            block_index = idx
                            break
                    if block_index is not None:
                        # Check previous block
                        if block_index > 0:
                            prev_block = blocks[block_index - 1]
                            if isinstance(prev_block, Paragraph) and 'caption' in prev_block.style.name.lower():
                                has_caption = True
                        # Check next block
                        if not has_caption and block_index + 1 < len(blocks):
                            next_block = blocks[block_index + 1]
                            if isinstance(next_block, Paragraph) and 'caption' in next_block.style.name.lower():
                                has_caption = True

                if not has_caption:
                    # Estimate page number based on shape position
                    page_number = 1
                    if shape_paragraph:
                        current_pos = sum(len(p.text) for p in doc.paragraphs[:doc.paragraphs.index(shape_paragraph)+1])
                        page_number = self.page_counter.get_page_number(current_pos)
                    issue = StyleIssue(
                        type="Missing Caption",
                        element="Figure",
                        expected="Figure should have a caption",
                        found="No caption found",
                        page=page_number,
                        section=None,
                        context=shape_paragraph.text[:50] if shape_paragraph else ""
                    )
                    issues.append(issue)
                    self.summary.caption_issues += 1
                    logger.debug(f"Missing caption for figure on page {page_number}")

                progress.update()

            progress.complete()

        logger.info("Style analysis complete.")
        return issues

    def _check_table_header_repeat(self, table: Table) -> bool:
        """Check if the table has header rows set to repeat on each page."""
        for row in table.rows:
            tr = row._tr  # Access the underlying CT_Row object
            trPr = tr.trPr  # Access the row properties (trPr)
            if trPr is not None:
                tbl_header = trPr.find(qn('w:tblHeader'))  # Find the tblHeader element
                if tbl_header is not None and tbl_header.val:
                    return True  # Header is set to repeat
        return False  # No header row is set to repeat

    def _get_shape_paragraph(self, doc, shape) -> Optional[Paragraph]:
        """Retrieve the paragraph object associated with the shape."""
        for paragraph in doc.paragraphs:
            if paragraph._element.contains(shape._inline):
                return paragraph
        return None

    def _save_results_with_summary(self, content_issues: List[Issue], style_issues: List[StyleIssue],
                                   output_path: str, format: OutputFormat):
        """Save analysis results with summary to file."""
        logger.info(f"Saving results to {output_path} in {format.value.upper()} format.")
        try:
            if format == OutputFormat.TXT:
                self._save_txt_with_summary(content_issues, style_issues, output_path)
            elif format == OutputFormat.CSV:
                self._save_csv_with_summary(content_issues, style_issues, output_path)
            elif format == OutputFormat.ORG:
                self._save_org_with_summary(content_issues, style_issues, output_path)
            elif format == OutputFormat.MD:
                self._save_md_with_summary(content_issues, style_issues, output_path)
            logger.info(f"Results successfully saved to {output_path}")
        except Exception as e:
            logger.error(f"Error saving results to {output_path}: {e}")
            print(f"Error saving results to {output_path}: {e}")

    def _save_txt_with_summary(self, content_issues: List[Issue], style_issues: List[StyleIssue], output_path: str):
        """Save results in plain text format with summary."""
        with open(output_path, 'w', encoding='utf-8') as f:
            # Write summary
            f.write("Document Analysis Summary\n")
            f.write("=======================\n\n")

            # Font Usage Summary
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

            # Acronym summary
            f.write("\nAcronyms Summary\n")
            f.write("----------------\n")
            for acronym, info in self.acronym_finder.found_acronyms.items():
                pages = ', '.join(map(str, sorted(info['pages'])))
                definition = info['definition'] or ""
                f.write(f"{acronym}: {definition} (Pages: {pages})\n")

            # Write detailed findings
            if content_issues or style_issues:
                f.write("\nDetailed Findings\n")
                f.write("=================\n")

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

    def _save_csv_with_summary(self, content_issues: List[Issue], style_issues: List[StyleIssue], output_path: str):
        """Save results in CSV format with summary."""
        with open(output_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)

            # Write summary
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
            writer.writerow(["Terminology conflicts"])
            writer.writerow(["Term", "Occurrences"])
            for term, count in self.summary.terminology.items():
                writer.writerow([term, count])

            # Style Issues Summary
            writer.writerow([])
            writer.writerow(["Style Issues Summary"])
            writer.writerow(["Font inconsistencies", self.summary.font_issues])
            writer.writerow(["Caption style issues", self.summary.caption_issues])
            writer.writerow(["Header style issues", self.summary.header_issues])
            writer.writerow(["Table style issues", self.summary.table_style_issues])

            # Acronyms Summary
            writer.writerow([])
            writer.writerow(["Acronyms Summary"])
            writer.writerow(["Acronym", "Definition", "Pages"])
            for acronym, info in self.acronym_finder.found_acronyms.items():
                pages = ', '.join(map(str, sorted(info['pages'])))
                definition = info['definition'] or ""
                writer.writerow([acronym, definition, pages])

            # Write detailed findings
            if content_issues or style_issues:
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
                    writer.writerow(["Type", "Element", "Expected", "Found", "Page", "Section", "Context"])
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

    def _save_org_with_summary(self, content_issues: List[Issue], style_issues: List[StyleIssue], output_path: str):
        """Save results in Org format with summary."""
        with open(output_path, 'w', encoding='utf-8') as f:
            # Write summary
            f.write("* Document Analysis Summary\n\n")

            # Font Usage Summary
            f.write("** Font Usage Analysis\n")
            for line in self.font_usage.get_formatted_summary().split('\n'):
                f.write(f"{line}\n")
            f.write("\n")

            # Content Issues Summary
            f.write("** Content Issues\n")
            f.write(f"- Pronouns found: {self.summary.pronouns}\n")
            f.write(f"- Contractions found: {self.summary.contractions}\n")
            f.write("- Terminology conflicts:\n")
            for term, count in self.summary.terminology.items():
                f.write(f"  - {term}: {count} occurrences\n")

            # Style Issues Summary
            f.write("\n** Style Issues\n")
            f.write(f"- Font inconsistencies: {self.summary.font_issues}\n")
            f.write(f"- Caption style issues: {self.summary.caption_issues}\n")
            f.write(f"- Header style issues: {self.summary.header_issues}\n")
            f.write(f"- Table style issues: {self.summary.table_style_issues}\n")

            # Acronyms Summary
            f.write("\n** Acronyms Summary\n")
            for acronym, info in self.acronym_finder.found_acronyms.items():
                pages = ', '.join(map(str, sorted(info['pages'])))
                definition = info['definition'] or ""
                f.write(f"- {acronym}: {definition} (Pages: {pages})\n")

            # Write detailed findings
            if content_issues or style_issues:
                f.write("\n* Detailed Findings\n")

            if content_issues:
                f.write("\n** Content Issues\n")
                for issue in content_issues:
                    f.write(f"\n*** {issue.type}\n")
                    f.write(f"- **Term**: {issue.term}\n")
                    f.write(f"- **Page**: {issue.page}\n")
                    if issue.section:
                        f.write(f"- **Section**: {issue.section}\n")
                    if issue.context:
                        f.write(f"- **Context**: {issue.context}\n")

            if style_issues:
                f.write("\n** Style Issues\n")
                for issue in style_issues:
                    f.write(f"\n*** {issue.type}\n")
                    f.write(f"- **Element**: {issue.element}\n")
                    f.write(f"- **Expected**: {issue.expected}\n")
                    f.write(f"- **Found**: {issue.found}\n")
                    f.write(f"- **Page**: {issue.page}\n")
                    if issue.section:
                        f.write(f"- **Section**: {issue.section}\n")
                    if issue.context:
                        f.write(f"- **Context**: {issue.context}\n")

    def _save_md_with_summary(self, content_issues: List[Issue], style_issues: List[StyleIssue], output_path: str):
        """Save results in Markdown format with summary."""
        with open(output_path, 'w', encoding='utf-8') as f:
            # Write summary
            f.write("# Document Analysis Summary\n\n")

            # Font Usage Summary
            f.write("## Font Usage Analysis\n")
            f.write(self.font_usage.get_formatted_summary())
            f.write("\n\n")

            # Content Issues Summary
            f.write("## Content Issues\n")
            f.write(f"- Pronouns found: {self.summary.pronouns}\n")
            f.write(f"- Contractions found: {self.summary.contractions}\n")
            f.write("- Terminology conflicts:\n")
            for term, count in self.summary.terminology.items():
                f.write(f"  - {term}: {count} occurrences\n")

            # Style Issues Summary
            f.write("\n## Style Issues\n")
            f.write(f"- Font inconsistencies: {self.summary.font_issues}\n")
            f.write(f"- Caption style issues: {self.summary.caption_issues}\n")
            f.write(f"- Header style issues: {self.summary.header_issues}\n")
            f.write(f"- Table style issues: {self.summary.table_style_issues}\n")

            # Acronyms Summary
            f.write("\n## Acronyms Summary\n")
            for acronym, info in self.acronym_finder.found_acronyms.items():
                pages = ', '.join(map(str, sorted(info['pages'])))
                definition = info['definition'] or ""
                f.write(f"- **{acronym}**: {definition} (Pages: {pages})\n")

            # Write detailed findings
            if content_issues or style_issues:
                f.write("\n# Detailed Findings\n\n")

            if content_issues:
                f.write("## Content Issues\n\n")
                for issue in content_issues:
                    f.write(f"### {issue.type}\n")
                    f.write(f"- **Term**: {issue.term}\n")
                    f.write(f"- **Page**: {issue.page}\n")
                    if issue.section:
                        f.write(f"- **Section**: {issue.section}\n")
                    if issue.context:
                        f.write(f"- **Context**: {issue.context}\n")
                    f.write("\n")

            if style_issues:
                f.write("## Style Issues\n\n")
                for issue in style_issues:
                    f.write(f"### {issue.type}\n")
                    f.write(f"- **Element**: {issue.element}\n")
                    f.write(f"- **Expected**: {issue.expected}\n")
                    f.write(f"- **Found**: {issue.found}\n")
                    f.write(f"- **Page**: {issue.page}\n")
                    if issue.section:
                        f.write(f"- **Section**: {issue.section}\n")
                    if issue.context:
                        f.write(f"- **Context**: {issue.context}\n")
                    f.write("\n")


# =============================
# Main Function
# =============================

def main():
    """Main function to run the document analyzer."""
    if len(sys.argv) < 2:
        logger.error("No document path provided.")
        print("Usage: python doc_analyzer.py <document_path> [config_path] [acronyms_csv]")
        sys.exit(1)

    doc_path = sys.argv[1]
    config_path = sys.argv[2] if len(sys.argv) > 2 else "config.json"
    acronyms_csv = sys.argv[3] if len(sys.argv) > 3 else None

    logger.info("\nDocument Analysis Tool")
    logger.info("=" * 50)
    print("\nDocument Analysis Tool")
    print("=" * 50)

    analyzer = DocumentAnalyzer(config_path=config_path, acronyms_csv=acronyms_csv)
    content_issues, style_issues = analyzer.analyze_document(doc_path)

    # Print summary to terminal
    print("\nAnalysis Summary")
    print("-" * 30)

    # Print Font Usage Summary
    print("\nFont Usage Analysis")
    print(analyzer.font_usage.get_formatted_summary())

    # Print Content Issues Summary
    print("\nContent Issues:")
    print(f"- Pronouns found: {analyzer.summary.pronouns}")
    print(f"- Contractions found: {analyzer.summary.contractions}")
    print("- Terminology conflicts:")
    for term, count in analyzer.summary.terminology.items():
        print(f"  - {term}: {count} occurrences")

    # Print Style Issues Summary
    print("\nStyle Issues:")
    print(f"- Font inconsistencies: {analyzer.summary.font_issues}")
    print(f"- Caption style issues: {analyzer.summary.caption_issues}")
    print(f"- Header style issues: {analyzer.summary.header_issues}")
    print(f"- Table style issues: {analyzer.summary.table_style_issues}")

    # Save results in configured formats
    logger.info("Saving results...")
    print("\nSaving results...")
    for format_str in analyzer.config.get("output_format", ["txt"]):
        try:
            format_enum = OutputFormat(format_str.lower())
            output_path = f"analysis_results.{format_str}"
            analyzer._save_results_with_summary(content_issues, style_issues, output_path, format_enum)
        except ValueError:
            logger.warning(f"Unsupported output format: {format_str}")
            print(f"Unsupported output format: {format_str}")

    logger.info("Analysis complete.")
    print("\nAnalysis complete.")
    logger.info("=" * 50)


if __name__ == "__main__":
    main()
