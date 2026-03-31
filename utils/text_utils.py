"""
Shared Text and HTML Utilities
"""
import re

def clean_html(html_string: str) -> str:
    """Strip all HTML tags and normalize whitespace."""
    if not html_string:
        return ""
    # Strip HTML tags
    clean = re.sub('<[^<]+?>', '', html_string)
    # Normalize whitespace (replace tabs/newlines with spaces)
    clean = re.sub(r'\s+', ' ', clean).strip()
    return clean

def normalize_filename(name: str) -> str:
    """Remove special characters and replace spaces with underscores for safe filenames."""
    if not name:
        return "document"
    # Allow alphanumeric, spaces, and hyphens; replace everything else
    clean = re.sub(r'[^\w\s-]', '', name).strip()
    return clean.replace(" ", "_")

def get_first_sentence(text: str, max_chars: int = 60) -> str:
    """Extract the first sentence or first N characters from text."""
    clean = clean_html(text)
    if not clean:
        return "New Chat"
    
    # Split by common sentence terminators but keep the terminator if it's there
    sentences = re.split(r'([.!?])', clean)
    if len(sentences) > 1:
        # Re-join first sentence + its terminator
        first = (sentences[0] + sentences[1]).strip()
    else:
        first = sentences[0].strip()
        
    return first[:max_chars]
