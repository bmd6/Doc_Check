# analyzers/acronym_analyzer.py
import re
import pandas as pd
from typing import Dict, Optional, Set, List
from pathlib import Path
from datetime import datetime

from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import _Cell, Table
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

from utils.logger import setup_logger
from utils.progress import ProgressTracker

logger = setup_logger(__name__)

class AcronymAnalyzer:
    """Analyzes documents for acronyms and their definitions."""
    
    def __init__(self, docx_path: str, known_acronyms_csv: Optional[str] = None,
                 excluded_acronyms_csv: Optional[str] = None):
        """
        Initialize the acronym analyzer.
        
        Args:
            docx_path: Path to the Word document
            known_acronyms_csv: Optional path to CSV file with known acronyms
            excluded_acronyms_csv: Optional path to CSV file with acronyms to exclude
        """
        self.docx_path = Path(docx_path)
        if not self.docx_path.exists():
            raise FileNotFoundError(f"Document not found: {docx_path}")
            
        # Initialize document tracking attributes
        self.current_page = 1
        self.chars_on_page = 0
        self.chars_per_page = 1800  # Approximate characters per page
        
        # Load acronyms and exclusions
        self.known_acronyms = self._load_known_acronyms(known_acronyms_csv) if known_acronyms_csv else {}
        self.excluded_acronyms = self._load_excluded_acronyms(excluded_acronyms_csv) if excluded_acronyms_csv else set()
        self.found_acronyms: Dict[str, Dict] = {}
        
        # Initialize found_acronyms with known acronyms
        for acronym, definition in self.known_acronyms.items():
            if acronym not in self.excluded_acronyms:
                self.found_acronyms[acronym] = {
                    'definition': definition,
                    'pages': set()
                }
        
        # Set up acronym patterns
        self.acronym_patterns = [
            r'^[A-Z0-9][A-Z0-9]{1,5}$',  # Basic acronyms (2-6 chars)
            r'^[A-Z][&][A-Z]$',           # Special case for X&Y format
            r'^[A-Z]+/[A-Z]+$',           # Slash-separated acronyms
            r'^[A-Z0-9]+-[A-Z0-9]+$'      # Hyphenated acronyms
        ]
        
        # Set up default exclusions
        self.default_exclusions = {'I', 'A', 'OK', 'ID', 'NO', 'AM', 'PM', 'THE'}
        
        logger.info(f"Acronym analyzer initialized for {docx_path}")
        logger.debug(f"Known acronyms: {len(self.known_acronyms)}")
        logger.debug(f"Excluded acronyms: {len(self.excluded_acronyms)}")

    def _load_known_acronyms(self, csv_path: str) -> Dict[str, str]:
        """
        Load known acronyms from CSV file.
        
        Args:
            csv_path: Path to CSV file containing acronyms and definitions
            
        Returns:
            Dictionary mapping acronyms to their definitions
        """
        try:
            df = pd.read_csv(csv_path)
            if 'Acronym' in df.columns and 'Definition' in df.columns:
                acronyms = dict(zip(df['Acronym'], df['Definition']))
                logger.info(f"Loaded {len(acronyms)} known acronyms from {csv_path}")
                return acronyms
            else:
                logger.error(f"CSV file {csv_path} missing required columns (Acronym, Definition)")
                return {}
        except Exception as e:
            logger.error(f"Error loading known acronyms: {e}")
            return {}

    def _load_excluded_acronyms(self, csv_path: str) -> Set[str]:
        """
        Load acronyms to exclude from CSV file.
        
        Args:
            csv_path: Path to CSV file containing acronyms to exclude
            
        Returns:
            Set of acronyms to exclude from analysis
        """
        try:
            df = pd.read_csv(csv_path)
            if 'Exclusion' in df.columns:
                exclusions = set(df['Exclusion'].dropna())
                logger.info(f"Loaded {len(exclusions)} excluded acronyms from {csv_path}")
                return exclusions
            else:
                logger.error(f"CSV file {csv_path} missing required column (Exclusion)")
                return set()
        except Exception as e:
            logger.error(f"Error loading excluded acronyms: {e}")
            return set()

    def _is_potential_acronym(self, word: str) -> bool:
        """
        Check if a word could be an acronym based on patterns.
        
        Args:
            word: Word to check
            
        Returns:
            Boolean indicating if the word might be an acronym
        """
        # Check exclusions first
        if word in self.default_exclusions or word in self.excluded_acronyms:
            return False
            
        return any(bool(re.match(pattern, word)) for pattern in self.acronym_patterns)

    def _find_potential_definition(self, text: str, acronym: str) -> Optional[str]:
        """
        Look for potential definition of an acronym in surrounding text.
        
        Args:
            text: Text to search for definition
            acronym: Acronym to find definition for
            
        Returns:
            Potential definition if found, None otherwise
        """
        escaped_acronym = re.escape(acronym)
        
        patterns = [
            rf'{escaped_acronym}\s*\(([\w\s,/-]+)\)',              # Acronym (Definition)
            rf'\(([\w\s,/-]+)\)\s*{escaped_acronym}',              # (Definition) Acronym
            rf'{escaped_acronym}:\s*([\w\s,/-]+)',                 # Acronym: Definition
            rf'{escaped_acronym}\s*-\s*([\w\s,/-]+)',             # Acronym - Definition
            rf'{escaped_acronym}\s+stands\s+for\s+([\w\s,/-]+)',   # Acronym stands for Definition
            rf'{escaped_acronym}\s+means\s+([\w\s,/-]+)'          # Acronym means Definition
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1).strip()
        return None

    def reset_page_tracking(self):
        """Reset page tracking counters."""
        self.current_page = 1
        self.chars_on_page = 0

    def update_page_tracking(self, text_length: int):
        """
        Update page tracking based on text length.
        
        Args:
            text_length: Length of text being processed
        """
        self.chars_on_page += text_length
        if self.chars_on_page > self.chars_per_page:
            self.current_page += 1
            self.chars_on_page = text_length

    def process_text(self, text: str, page_number: int) -> None:
        """
        Process text to find acronyms and their definitions.
        
        Args:
            text: Text to analyze
            page_number: Current page number (must be int)
        """
        try:
            # Ensure page_number is an integer
            page_number = int(page_number)
            words = re.findall(r'\b[\w/&-]+\b', text)
            
            for word in words:
                if self._is_potential_acronym(word):
                    if word not in self.found_acronyms:
                        # Look for definition in surrounding text
                        definition = self._find_potential_definition(text, word)
                        # If no definition found in text, use known acronym definition
                        if not definition:
                            definition = self.known_acronyms.get(word)
                        
                        self.found_acronyms[word] = {
                            'definition': definition,
                            'pages': set()
                        }
                    
                    self.found_acronyms[word]['pages'].add(page_number)
        except Exception as e:
            logger.error(f"Error processing text for acronyms: {e}")
            logger.debug(f"Page number type: {type(page_number)}, value: {page_number}")
            raise

    def process_paragraph(self, paragraph: Paragraph, page_number: int) -> None:
        """
        Process a paragraph to find acronyms.
        
        Args:
            paragraph: Paragraph to analyze
            page_number: Current page number (must be int)
        """
        try:
            page_number = int(page_number)
            self.process_text(paragraph.text, page_number)
        except Exception as e:
            logger.error(f"Error processing paragraph: {e}")
            raise

    def process_table(self, table: Table, page_number: int) -> None:
        """
        Process a table to find acronyms.
        
        Args:
            table: Table to analyze
            page_number: Current page number (must be int)
        """
        try:
            page_number = int(page_number)
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        self.process_text(paragraph.text, page_number)
        except Exception as e:
            logger.error(f"Error processing table: {e}")
            raise

    def create_acronym_report(self) -> Document:
        """
        Create a document containing the acronym reference table.
        
        Returns:
            Word document containing formatted acronym report
        """
        doc = Document()
        
        # Add title
        title = doc.add_paragraph("Acronyms Found")
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title.style = 'Heading 1'
        
        # Create table with only acronym and definition columns
        table = doc.add_table(rows=1, cols=2)
        table.style = 'Table Grid'
        table.autofit = True
        
        # Set column widths
        for column in table.columns:
            for cell in column.cells:
                cell.width = Inches(3)
        
        # Set headers
        headers = table.rows[0].cells
        headers[0].text = 'Acronym'
        headers[1].text = 'Definition'
        
        # Populate table, sorted alphabetically by acronym
        for acronym, info in sorted(self.found_acronyms.items()):
            row = table.add_row().cells
            row[0].text = acronym
            row[1].text = info['definition'] or 'Unknown'
        
        return doc

    def analyze_document(self) -> Dict[str, Dict]:
        """
        Analyze the entire document for acronyms.
        
        Returns:
            Dictionary containing found acronyms and their information
        """
        try:
            doc = Document(self.docx_path)
            self.reset_page_tracking()
            
            # Count total elements for progress tracking
            total_elements = len(doc.paragraphs) + sum(1 for _ in doc.tables)
            progress = ProgressTracker(total_elements, "Analyzing acronyms")
            
            # Process paragraphs
            for paragraph in doc.paragraphs:
                self.update_page_tracking(len(paragraph.text))
                self.process_paragraph(paragraph, self.current_page)
                progress.update()
            
            # Process tables
            for table in doc.tables:
                table_text = "".join(
                    cell.text for row in table.rows for cell in row.cells
                )
                self.update_page_tracking(len(table_text))
                self.process_table(table, self.current_page)
                progress.update()
            
            progress.complete()
            
            # Create and save acronym report
            report = self.create_acronym_report()
            output_path = self.docx_path.with_name(f"{self.docx_path.stem}_acronyms.docx")
            report.save(output_path)
            logger.info(f"Saved acronym report to {output_path}")
            
            return self.found_acronyms
            
        except Exception as e:
            logger.error(f"Error analyzing document: {e}")
            raise

    def get_statistics(self) -> Dict[str, int]:
        """
        Get statistics about found acronyms.
        
        Returns:
            Dictionary containing acronym statistics
        """
        total_acronyms = len(self.found_acronyms)
        defined_acronyms = sum(
            1 for info in self.found_acronyms.values() 
            if info['definition']
        )
        undefined_acronyms = total_acronyms - defined_acronyms
        
        return {
            'total_acronyms': total_acronyms,
            'defined_acronyms': defined_acronyms,
            'undefined_acronyms': undefined_acronyms
        }