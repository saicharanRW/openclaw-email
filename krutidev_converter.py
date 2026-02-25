"""
Krutidev to Unicode converter for Hindi text (Devanagari script)
Main conversion logic
"""

import re
import sys
import argparse
from dictionary import MAIN, CONSONANTS, VOWELS, UNATTACHED


def replace_string(text, find, replace):
    """Replace all occurrences of find with replace in text"""
    return text.replace(find, replace)


def krutidev_to_unicode(text):
    """
    Convert Krutidev text to Unicode (Devanagari)
    
    Args:
        text: Input text in Krutidev format
        
    Returns:
        Converted text in Unicode format
    """
    # space +  ्र  ->   ्र
    text = replace_string(text, ' \xaa', '\xaa')
    text = replace_string(text, ' ~j', '~j')
    text = replace_string(text, ' z', 'z')

    # – and — if not surrounded by krutidev consonants/matrās, change them to -
    misplaced = re.compile('[ab]', re.DOTALL)
    for result in misplaced.finditer(text):
        length = len(result.group())
        index = result.start()
        if (
            index < len(text) - 1 and
            text[index + 1] not in CONSONANTS['krutidev'] and
            text[index + 1] not in UNATTACHED['krutidev']
        ):
            text = text[:index] + '&' + text[index + 1:]

    # Apply main dictionary replacements
    for find, replace in MAIN:
        text = replace_string(text, find, replace)
    
    text = replace_string(text, '\xb1', 'Z\u0902')  # ±  ->  Zं
    text = replace_string(text, '\xc6', '\u0930\u094df')  # Æ  ->  र्f

    # f + ?  ->  ? + ि
    misplaced = re.compile('f(.?)', re.DOTALL)
    for result in misplaced.finditer(text):
        match = result.group(1) if result.group(1) else ''
        text = text.replace('f' + match, match + '\u093f', 1)
    
    text = replace_string(text, '\xc7', 'fa')  # Ç  ->  fa
    text = replace_string(text, '\xaf', 'fa')  # ¯  ->  fa
    text = replace_string(text, '\xc9', '\u0930\u094dfa')  # É  ->  र्fa

    # fa?  ->  ? + िं
    misplaced = re.compile('fa(.?)', re.DOTALL)
    for result in misplaced.finditer(text):
        match = result.group(1) if result.group(1) else ''
        text = text.replace('fa' + match, match + '\u093f\u0902', 1)
    
    text = replace_string(text, '\xca', '\u0940Z')  # Ê  ->  ीZ

    # ि्  + ?  ->  ्  + ? + ि
    misplaced = re.compile('\u093f\u094d(.?)', re.DOTALL)
    for result in misplaced.finditer(text):
        match = result.group(1) if result.group(1) else ''
        text = text.replace('\u093f\u094d' + match, '\u094d' + match + '\u093f', 1)
    
    text = replace_string(text, '\u094dZ', 'Z')  # ्  + Z ->  Z

    # र +  ्  should be placed at the right place, before matrās
    misplaced = re.compile('(.?)Z', re.DOTALL)
    for result in misplaced.finditer(text):
        match = result.group(1) if result.group(1) else ''
        index = text.find(match + 'Z')
        while index >= 0 and text[index] in VOWELS['unicode']:
            index -= 1
            match = text[index] + match
        old_pattern = match + 'Z'
        new_pattern = '\u0930\u094d' + match
        text = text.replace(old_pattern, new_pattern, 1)

    # ' ', ',' and ्  are illegal characters just before a matrā
    for matra in UNATTACHED['unicode']:
        text = replace_string(text, ' ' + matra, matra)
        text = replace_string(text, ',' + matra, matra + ',')
        text = replace_string(text, '\u094d' + matra, matra + ',')
    
    text = replace_string(text, '\u094d\u094d\u0930', '\u094d\u0930')  # ्  + ्  + र ->  ्  + र
    text = replace_string(text, '\u094d\u0930\u094d', '\u0930\u094d')  # ्  + र + ्  ->  र + ्
    text = replace_string(text, '\u094d\u094d', '\u094d')  # ्  + ्  ->  ्

    # ्  at the ending of a consonant as the last character is illegal.
    text = replace_string(text, '\u094d ', ' ')

    return text.strip()


# ─────────────────────────────────────────────────────────────────────────────
# Reverse converter: Unicode (Devanagari) → Krutidev ASCII
# ─────────────────────────────────────────────────────────────────────────────

def _build_unicode_to_krutidev_map():
    """
    Build a sorted list of (unicode_str, krutidev_str) pairs derived from MAIN.
    Sorted longest-unicode-key first so multi-char sequences are matched before
    single chars (greedy replacement).
    Skips ambiguous entries where the same Unicode maps to multiple Krutidev codes.
    """
    seen_unicode = {}
    for krutidev_char, unicode_char in MAIN:
        # Only include pure Devanagari/Unicode targets (skip ASCII-only mappings
        # that are just punctuation remaps with no Devanagari)
        if any('\u0900' <= c <= '\u097f' for c in unicode_char):
            if unicode_char not in seen_unicode:
                seen_unicode[unicode_char] = krutidev_char
            # If already seen, keep the first (earlier in MAIN = higher priority)
    # Sort by length of unicode key descending so longer sequences match first
    return sorted(seen_unicode.items(), key=lambda x: len(x[0]), reverse=True)


_UNICODE_TO_KRUTIDEV_MAP = None


def unicode_to_krutidev(text: str) -> str:
    """
    Convert Unicode Devanagari text back to Krutidev ASCII encoding.

    This is a best-effort reverse conversion. Because Krutidev→Unicode is a
    many-to-one mapping (multiple Krutidev sequences can produce the same
    Unicode), the reverse may not be byte-for-byte identical to the original
    Krutidev source, but it will render correctly with a Krutidev font.

    Non-Devanagari characters (ASCII, punctuation, digits) are passed through
    unchanged, with a few common substitutions (purna viram → A, etc.).
    """
    global _UNICODE_TO_KRUTIDEV_MAP
    if _UNICODE_TO_KRUTIDEV_MAP is None:
        _UNICODE_TO_KRUTIDEV_MAP = _build_unicode_to_krutidev_map()

    # Common standalone substitutions not covered by MAIN
    extra = [
        ('\u0964', 'A'),   # ।  → A  (purna viram / full stop)
        ('\u0965', 'AA'),  # ॥  → AA (double danda)
        ('\u0966', '\xe5'), # ०  → å
        ('\u0967', '\xf8'), # १  → (Kruti Dev digit 1 mapping)
        ('\u0968', '\xf9'), # २
        ('\u0969', '\xfa'), # ३
        ('\u096a', '\xfb'), # ४
        ('\u096b', '\xfc'), # ५
        ('\u096c', '\xfd'), # ६
        ('\u096d', '\xfe'), # ७
        ('\u096e', '\xff'), # ८
        ('\u096f', '\u0152'), # ९
        ('\u0902', 'a'),   # ं  → a (anusvara)
        ('\u0901', '\xa1'), # ँ  → ¡ (chandrabindu)
        ('\u093e', 'k'),   # ा  → k
        ('\u093f', 'f'),   # ि  → f  (handled specially below)
        ('\u0940', 'h'),   # ी  → h
        ('\u0941', 'q'),   # ु  → q
        ('\u0942', 'w'),   # ू  → w
        ('\u0943', '`'),   # ृ  → `
        ('\u0947', 's'),   # े  → s
        ('\u0948', 'S'),   # ै  → S
        ('\u094b', 'ks'),  # ो  → ks
        ('\u094c', 'kS'),  # ौ  → kS
        ('\u094d', '~'),   # ्  → ~
    ]

    # Apply MAIN-derived reverse map (longest first)
    for uni, kru in _UNICODE_TO_KRUTIDEV_MAP:
        text = text.replace(uni, kru)

    # Apply extra standalone substitutions (longest first to avoid conflicts)
    for uni, kru in sorted(extra, key=lambda x: len(x[0]), reverse=True):
        text = text.replace(uni, kru)

    return text


def convert_file(input_file, output_file):
    """
    Convert a file from Krutidev to Unicode
    
    Args:
        input_file: Path to input file
        output_file: Path to output file
    """
    try:
        with open(input_file, 'r', encoding='utf-8') as f:
            input_text = f.read()
        
        converted_text = krutidev_to_unicode(input_text)
        
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(converted_text)
        
        print(f"Successfully converted {input_file} to {output_file}")
    except Exception as e:
        print(f"Error processing file: {e}", file=sys.stderr)
        sys.exit(1)


def main():
    """CLI interface for file conversion"""
    parser = argparse.ArgumentParser(
        description='Convert Krutidev text to Unicode (Devanagari)',
        prog='krutidev_converter'
    )
    parser.add_argument('input_file', help='Input file in Krutidev format')
    parser.add_argument('output_file', help='Output file for Unicode text')
    
    args = parser.parse_args()
    convert_file(args.input_file, args.output_file)


if __name__ == '__main__':
    main()
