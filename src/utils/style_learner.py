"""
Email Style Learning Module

Analyzes user's sent emails to learn their preferred formatting style.
No caching - learns dynamically from recent sent items.

Version: 1.0.2
"""

import re
from collections import Counter
from typing import Dict, Optional
import win32com.client
from config import OUTLOOK_FOLDER_SENT, STYLE_LEARNING_EMAIL_COUNT


def extract_html_styles(html_body: str) -> Dict[str, str]:
    """
    Extract font styles from HTML body inline styles (body content only, excludes signature).
    
    Finds the actual signature (table with logo/contact info) and extracts styles
    only from content BEFORE the signature table.
    
    Args:
        html_body: HTML content of email
        
    Returns:
        Dictionary with style properties (font-family, font-size, color)
    """
    styles = {}
    
    # Extract only the body content (before signature)
    # The signature is typically in a table or starts with specific contact info
    body_content = html_body
    
    # Try multiple signature detection methods:
    # 1. Look for signature table (most Outlook signatures are in tables)
    sig_table_match = re.search(r'<table[^>]*>.*?Vincent PAPUCHON', html_body, re.IGNORECASE | re.DOTALL)
    if sig_table_match:
        body_content = html_body[:sig_table_match.start()]
    else:
        # 2. Look for signature text patterns
        sig_text_patterns = [
            r'<p[^>]*>\s*<span[^>]*>\s*Vincent PAPUCHON',
            r'<span[^>]*>Vincent PAPUCHON</span>',
            r'Analyste Etude Système',
        ]
        for pattern in sig_text_patterns:
            sig_match = re.search(pattern, html_body, re.IGNORECASE)
            if sig_match:
                body_content = html_body[:sig_match.start()]
                break
    
    # Extract ALL properties from inline styles in body content only
    # Need to handle quotes inside style values: style='font-family:"Calibri"'
    # Match either single or double quotes, but capture everything until matching quote
    inline_styles = []
    for match in re.finditer(r'style=(["\'])(.+?)\1', body_content, re.IGNORECASE):
        inline_styles.append(match.group(2))
    
    # Collect all font-family, font-size, and color from inline styles
    font_families = []
    font_sizes = []
    colors = []
    
    for inline_style in inline_styles:
        # font-family can contain quotes: "Calibri",sans-serif
        # Capture everything until semicolon or end, including quoted values
        font_family_match = re.search(r'font-family:\s*([^;]+)', inline_style, re.IGNORECASE)
        if font_family_match:
            family = font_family_match.group(1).strip()
            # Clean up: remove quotes and keep first font only
            family = family.replace('"', '').replace("'", "")
            # Take first font name before comma
            if ',' in family:
                family = family.split(',')[0].strip()
            font_families.append(family)
        
        font_size_match = re.search(r'font-size:\s*([^;]+)', inline_style, re.IGNORECASE)
        if font_size_match:
            font_sizes.append(font_size_match.group(1).strip())
        
        # color can be named (green, red) or hex (#004080)
        color_match = re.search(r'color:\s*([^;]+)', inline_style, re.IGNORECASE)
        if color_match:
            color = color_match.group(1).strip()
            # Convert named colors to hex if needed (keep as is for now)
            colors.append(color)
    
    # Use FIRST value found (user content is typically first, before empty paragraphs)
    # This avoids learning from empty formatting spans added by Outlook
    if font_families:
        styles['font-family'] = font_families[0]
    
    if font_sizes:
        styles['font-size'] = font_sizes[0]
    
    if colors:
        styles['color'] = colors[0]
    
    return styles


def learn_user_email_style(num_emails: int = None) -> Optional[Dict[str, str]]:
    """
    Analyze recent sent emails to learn user's preferred style.
    
    Learns from body content only (excludes signature).
    No caching - always analyzes sent items dynamically.
    
    Args:
        num_emails: Number of recent sent emails to analyze (default: from config.STYLE_LEARNING_EMAIL_COUNT)
        
    Returns:
        Dictionary with learned style, or None if analysis failed
    """
    if num_emails is None:
        num_emails = STYLE_LEARNING_EMAIL_COUNT
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        sent_folder = namespace.GetDefaultFolder(OUTLOOK_FOLDER_SENT)
        
        # Get recent sent items
        items = sent_folder.Items
        items.Sort("[SentOn]", True)  # Sort by sent date, descending
        
        # Collect styles from recent emails
        font_families = []
        font_sizes = []
        colors = []
        
        count = 0
        for item in items:
            if count >= num_emails:
                break
            
            try:
                # Only analyze emails with HTML body
                if hasattr(item, 'HTMLBody') and item.HTMLBody:
                    styles = extract_html_styles(item.HTMLBody)
                    
                    if 'font-family' in styles:
                        font_families.append(styles['font-family'])
                    if 'font-size' in styles:
                        font_sizes.append(styles['font-size'])
                    if 'color' in styles:
                        colors.append(styles['color'])
                    
                    count += 1
            except Exception:
                # Skip problematic items
                continue
        
        # No emails to analyze
        if count == 0:
            return None
        
        # Find most common values
        learned_style = {}
        
        if font_families:
            learned_style['font-family'] = Counter(font_families).most_common(1)[0][0]
        
        if font_sizes:
            learned_style['font-size'] = Counter(font_sizes).most_common(1)[0][0]
        
        if colors:
            learned_style['color'] = Counter(colors).most_common(1)[0][0]
        
        return learned_style if learned_style else None
        
    except Exception as e:
        print(f"Error learning email style: {e}")
        return None


def apply_style_to_html(body: str, style: Dict[str, str]) -> str:
    """
    Apply learned style to HTML body using inline styles on span elements.
    
    Outlook ignores CSS in <head> during email send, so we MUST use inline styles.
    Wrap content in <span> with explicit inline font-size, font-family, color.
    
    Args:
        body: Plain text or HTML body
        style: Dictionary with style properties
        
    Returns:
        HTML body with applied inline style in span
    """
    if not style:
        return body
        
    # Build COMPLETE style string including font-size (inline styles are respected)
    style_parts = []
    if 'font-family' in style:
        style_parts.append(f"font-family: {style['font-family']}")
    if 'font-size' in style:
        style_parts.append(f"font-size: {style['font-size']}")
    if 'color' in style:
        style_parts.append(f"color: {style['color']}")
    
    style_str = "; ".join(style_parts)
    
    # Convert plain text to HTML with style if needed
    if not body.strip().startswith('<'):
        # Plain text - wrap in SPAN with inline style
        body = body.replace('\n', '<br>\n')
        return f'<span style="{style_str}">{body}</span>'
    else:
        # Already HTML - add style to body tag or wrap in span
        if '<body' in body.lower():
            # Add style to existing body tag
            body = re.sub(
                r'<body([^>]*)>',
                f'<body\\1 style="{style_str}">',
                body,
                flags=re.IGNORECASE
            )
            return body
        else:
            # Wrap in span with inline style
            return f'<span style="{style_str}">{body}</span>'


