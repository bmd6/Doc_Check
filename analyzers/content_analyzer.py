# analyzers/content_analyzer.py
import re
from typing import List, Set, Dict, Tuple, DefaultDict
from collections import defaultdict
from itertools import tee

from docx.document import Document
from docx.text.paragraph import Paragraph

from models.issues import Issue
from models.summary import AnalysisSummary
from models.fonts import FontUsageSummary
from utils.logger import setup_logger
from utils.progress import ProgressTracker

logger = setup_logger(__name__)

class ContentAnalyzer:
    """
    Analyzes document content for writing style and terminology consistency.
    
    This analyzer focuses on:
    - Pronoun usage
    - Contractions
    - Terminology consistency
    - Word choice patterns
    """
    
    def __init__(self, config: Dict):
        """
        Initialize the content analyzer.
        
        Args:
            config: Configuration dictionary containing terminology groups and settings
        """
        self.config = config
        self.summary = AnalysisSummary()
        self.font_usage = FontUsageSummary()
        
        # Initialize sets of words to check
        self.pronouns = {
            "he", "him", "his", "she", "her", "hers", "it", "its",
            "they", "them", "their", "theirs", "we", "us", "our", "ours",
            "i", "me", "my", "mine", "you", "your", "yours"
        }
        
        self.contractions = {
            "n't", "'ll", "'re", "'ve", "'m", "'d", "'s"
        }
        
        # Prepare terminology matching
        self._prepare_terminology()
        logger.info("Content analyzer initialized successfully")

    def _prepare_terminology(self) -> None:
        """Prepare terminology patterns for efficient matching."""
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
        logger.info("Terminology preparation complete")

    def _generate_term_variations(self, term: str) -> Set[str]:
        """
        Generate common variations of a term.
        
        Args:
            term: Base term to generate variations for
            
        Returns:
            Set of term variations
        """
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

            # Handle optional parts in parentheses
            term_without_optional = re.sub(r'\s*\([^)]+\)\s*', ' ', term)
            if term_without_optional != term:
                variations.add(term_without_optional.strip())

        return variations

    def _normalize_term(self, term: str) -> str:
        """
        Normalize a term for consistent matching.
        
        Args:
            term: Term to normalize
            
        Returns:
            Normalized term
        """
        normalized = ' '.join(term.lower().split())
        normalized = re.sub(r'[-_&]', ' ', normalized)
        normalized = normalized.replace(' and ', ' ')
        return normalized

    def _sliding_window(self, sequence: List[str], window_size: int):
        """
        Create a sliding window iterator over a sequence.
        
        Args:
            sequence: List of items to create windows from
            window_size: Size of each window
            
        Returns:
            Iterator of windows
        """
        iterators = tee(sequence, window_size)
        for i, iterator in enumerate(iterators):
            for _ in range(i):
                next(iterator, None)
        return zip(*iterators)

    def _get_context(self, words: List[str], target: str, context_size: int = 5) -> str:
        """
        Get surrounding context for a term.
        
        Args:
            words: List of words from the text
            target: Target term to get context for
            context_size: Number of words before and after target to include
            
        Returns:
            Context string
        """
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

    def _analyze_pronouns(self, words: List[str], page: int, section: str) -> List[Issue]:
        """
        Identify pronouns in text.
        
        Args:
            words: List of words to analyze
            page: Current page number
            section: Current document section
            
        Returns:
            List of pronoun issues found
        """
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

    def _analyze_contractions(self, text: str, page: int, section: str) -> List[Issue]:
        """
        Identify contractions in text.
        
        Args:
            text: Text to analyze
            page: Current page number
            section: Current document section
            
        Returns:
            List of contraction issues found
        """
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

    def _analyze_terminology(self, full_text: str, words: List[str], 
                           page: int, section: str) -> List[Issue]:
        """
        Analyze text for terminology consistency issues.
        
        Args:
            full_text: Complete text to analyze
            words: List of words from the text
            page: Current page number
            section: Current document section
            
        Returns:
            List of terminology issues found
        """
        found_terms = {}
        term_positions = {}

        # Check each possible phrase length
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

        # Filter out redundant terms
        filtered_issues = []
        for group_name, terms in found_terms.items():
            if len(terms) > 1:
                sorted_terms = sorted(terms, key=lambda x: len(x[1].split()), reverse=True)
                covered_positions = set()

                for term, normalized_term, original in sorted_terms:
                    current_pos = term_positions[normalized_term]
                    is_redundant = False

                    # Check for overlapping terms
                    for start, end in covered_positions:
                        if (current_pos[0] >= start and current_pos[0] < end) or \
                           (current_pos[1] > start and current_pos[1] <= end):
                            related_terms = self.term_relationships.get(normalized_term, set())
                            for other_term, _, _ in sorted_terms:
                                other_normalized = self._normalize_term(other_term)
                                if other_normalized in related_terms:
                                    is_redundant = True
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

        return filtered_issues

    def analyze_paragraph(self, paragraph: Paragraph, page: int, 
                        current_section: str) -> List[Issue]:
        """
        Analyze a single paragraph for content issues.
        
        Args:
            paragraph: Paragraph to analyze
            page: Current page number
            current_section: Current document section
            
        Returns:
            List of content issues found
        """
        issues = []
        text = paragraph.text
        words = text.split()

        # Analyze different aspects
        pronoun_issues = self._analyze_pronouns(words, page, current_section)
        contraction_issues = self._analyze_contractions(text, page, current_section)
        terminology_issues = self._analyze_terminology(text, words, page, current_section)

        # Update summary
        self.summary.pronouns += len(pronoun_issues)
        self.summary.contractions += len(contraction_issues)
        for issue in terminology_issues:
            self.summary.terminology[issue.normalized_term] += 1

        # Combine all issues
        issues.extend(pronoun_issues)
        issues.extend(contraction_issues)
        issues.extend(terminology_issues)

        return issues