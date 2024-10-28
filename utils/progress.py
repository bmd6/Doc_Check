import sys
import time
from typing import Optional

class ProgressTracker:
    """
    Helper class to track and display progress.
    
    Attributes:
        total: Total number of steps
        current: Current step
        description: Description of the progress operation
        start_time: Start time of the operation
    """
    
    def __init__(self, total_steps: int, description: str):
        """
        Initialize progress tracker.
        
        Args:
            total_steps: Total number of steps to track
            description: Description of the operation
        """
        self.total = total_steps
        self.current = 0
        self.description = description
        self.start_time = time.time()
        self._print_progress()

    def update(self, steps: int = 1) -> None:
        """
        Update progress by specified number of steps.
        
        Args:
            steps: Number of steps to increment (default: 1)
        """
        self.current += steps
        self._print_progress()

    def _print_progress(self) -> None:
        """Print current progress to console."""
        percentage = (self.current / self.total) * 100 if self.total > 0 else 0
        elapsed_time = time.time() - self.start_time
        sys.stdout.write(
            f"\r{self.description}: [{self.current}/{self.total}] "
            f"{percentage:.1f}% (Elapsed: {elapsed_time:.1f}s)"
        )
        sys.stdout.flush()

    def complete(self) -> None:
        """Mark progress as complete."""
        self.current = self.total
        self._print_progress()
        print()