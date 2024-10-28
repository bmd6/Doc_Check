import json
import logging
from typing import Dict, Any
from pathlib import Path

logger = logging.getLogger(__name__)

class ConfigLoader:
    """Handles loading and validation of configuration files."""
    
    @staticmethod
    def load_config(config_path: str) -> Dict[str, Any]:
        """
        Load and validate configuration from JSON file.
        
        Args:
            config_path: Path to configuration file
            
        Returns:
            Validated configuration dictionary
            
        Raises:
            FileNotFoundError: If config file doesn't exist
            json.JSONDecodeError: If config file is invalid JSON
            ValueError: If config validation fails
        """
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            logger.info(f"Configuration loaded from {config_path}")
            return ConfigLoader._validate_config(config)
        except FileNotFoundError:
            logger.warning(f"Config file not found at {config_path}. Using default settings.")
            return ConfigLoader._validate_config({})
        except json.JSONDecodeError as e:
            logger.error(f"Error parsing config file: {e}")
            raise

    @staticmethod
    def _validate_config(config: Dict[str, Any]) -> Dict[str, Any]:
        """
        Validate and set default configuration values.
        
        Args:
            config: Configuration dictionary to validate
            
        Returns:
            Validated configuration with defaults applied
            
        Raises:
            ValueError: If configuration is invalid
        """
        default_config = {
            "terminology_groups": {},
            "output_format": ["txt"],
            "include_context": True
        }

        merged_config = default_config.copy()
        if isinstance(config, dict):
            for key in default_config:
                if key in config:
                    merged_config[key] = config[key]
                else:
                    logger.warning(f"Missing '{key}' in config. Using default value.")

        # Validate types
        if not isinstance(merged_config["terminology_groups"], dict):
            raise ValueError("'terminology_groups' should be a dictionary.")
        if not isinstance(merged_config["output_format"], list):
            raise ValueError("'output_format' should be a list.")
        if not isinstance(merged_config["include_context"], bool):
            raise ValueError("'include_context' should be a boolean.")

        return merged_config
