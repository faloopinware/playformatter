import streamlit as st
import re
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
from io import BytesIO
import PyPDF2

# Page configuration
st.set_page_config(
    page_title="The TWNC FaloopinFormatter",
    page_icon="\U0001f3ad",
    layout="wide"
)

st.title("\U0001f3ad The TWNC FaloopinFormatter")
st.markdown("Upload your script and we'll reformat it to professional theatrical standards")

# Sidebar
with st.sidebar:
    st.header("\U0001f4cb How It Works")
    st.markdown("""
    **3 Simple Steps:**

    **1\ufe0f\u20e3 Upload Your Script**
    - PDF, DOCX, or TXT file
    - We'll read and analyze it

    **2\ufe0f\u20e3 Enter Details**
    - Provide title, author, and play details
    - Edit character list and scene/time info

    **3\ufe0f\u20e3 Download**
    - Get your professionally formatted .DOCX file
    - Formatted per the standard stage play format

    ---

    **What We Fix:**
    - Courier 12pt font throughout
    - Proper title page layout
    - Cast of Characters page
    - Character names indented & ALL CAPS
    - Dialogue full-width (left to right margin)
    - Stage directions indented in parentheses/italics
    - Act/Scene headings properly placed
    - SETTING/AT RISE two-column layout
    - Page numbers (Act-Scene-Page)
    """)


# ---------------------------------------------------------------------------
# TEXT EXTRACTION
# ---------------------------------------------------------------------------

def _reconstruct_lines_from_pdf(raw_text):
    """Reconstruct proper lines from fragmented PDF extraction.

    PyPDF2 often inserts newlines between every word. This function
    rebuilds paragraphs by treating blank-only lines as separators.
    Multiple consecutive blank lines indicate paragraph breaks; single
    blank lines separate words within the same line.
    """
    tokens = raw_text.split('\n')
    paragraphs = []
    current_words = []

    for token in tokens:
        stripped = token.strip()
        if stripped == '':
            # Blank line - if we have accumulated words, that might be a paragraph break
            # But a single blank between words is normal. We look at patterns:
            # consecutive blanks = paragraph break
            if not current_words:
                # Already at a break, just note it
                if paragraphs and paragraphs[-1] != '':
                    paragraphs.append('')  # mark paragraph break
            else:
                # Check if next token is also blank (lookahead handled below)
                current_words.append(None)  # placeholder for potential break
        elif stripped in (',', '.', '!', '?', ':', ';', ')', '"', "'", '\u2019', '\u201d'):
            # Punctuation that should attach to previous word
            if current_words:
                # Remove trailing break placeholders
                while current_words and current_words[-1] is None:
                    current_words.pop()
                if current_words:
                    current_words[-1] = (current_words[-1] or '') + stripped
                else:
                    current_words.append(stripped)
            else:
                if paragraphs:
                    paragraphs[-1] = paragraphs[-1] + stripped
                else:
                    current_words.append(stripped)
        else:
            # Real word - check how many break placeholders preceded it
            break_count = 0
            while current_words and current_words[-1] is None:
                current_words.pop()
                break_count += 1

            if break_count >= 2 and current_words:
                # Paragraph break: flush current words
                paragraphs.append(' '.join(w for w in current_words if w))
                current_words = [stripped]
            elif break_count >= 1 and current_words:
                # Could be paragraph break or just word spacing
                # Heuristic: if prev word ends with sentence-ending punct, it's a paragraph break
                last_word = current_words[-1] if current_words else ''
                if (last_word and last_word[-1] in ('.', '!', '?', ':', '"', '\u201d') and
                    (stripped[0].isupper() or stripped.startswith('('))):
                    # Likely a new paragraph
                    paragraphs.append(' '.join(w for w in current_words if w))
                    current_words = [stripped]
                else:
                    current_words.append(stripped)
            else:
                current_words.append(stripped)

    # Flush remaining
    if current_words:
        remaining = [w for w in current_words if w]
        if remaining:
            paragraphs.append(' '.join(remaining))

    # Clean up: remove empty strings from the list but keep them as paragraph breaks
    result = []
    for p in paragraphs:
        cleaned = p.strip()
        if cleaned:
            # Fix spacing around punctuation that may have extra spaces
            cleaned = re.sub(r'\s+([,\.!?;:\)\u2019\u201d])', r'\1', cleaned)
            cleaned = re.sub(r'([\(\u2018\u201c])\s+', r'\1', cleaned)
            result.append(cleaned)
        else:
            result.append('')

    return '\n'.join(result)


def extract_text_from_pdf(pdf_file):
    """Extract text from PDF and reconstruct proper lines."""
    try:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        all_text = []
        for page in pdf_reader.pages:
            page_text = page.extract_text()
            if page_text:
                all_text.append(page_text)
        raw = "\n".join(all_text)
        return _reconstruct_lines_from_pdf(raw)
    except Exception as e:
        st.error(f"Error reading PDF: {str(e)}")
        return None


def extract_text_from_docx(docx_file):
    """Extract text from DOCX file."""
    try:
        doc = Document(docx_file)
        text = ""
        for para in doc.paragraphs:
            text += para.text + "\n"
        return text
    except Exception as e:
        st.error(f"Error reading DOCX: {str(e)}")
        return None


def extract_text_from_txt(txt_file):
    """Extract text from TXT file."""
    try:
        return txt_file.read().decode('utf-8')
    except Exception as e:
        st.error(f"Error reading TXT: {str(e)}")
        return None


# ---------------------------------------------------------------------------
# SCRIPT PARSER
# ---------------------------------------------------------------------------

def _is_character_name(line, known_characters):
    """Decide whether *line* is a character name heading."""
    stripped = line.strip()
    if not stripped:
        return False

    # Already known
    if stripped in known_characters or stripped.rstrip('.') in known_characters:
        return True

    # Combined names like "RENA and RACHEL", "ANGELA AND KATHRYN"
    if ' AND ' in stripped.upper() or ' and ' in stripped:
        parts = re.split(r'\s+[Aa][Nn][Dd]\s+', stripped)
        if all(p.strip() in known_characters for p in parts):
            return True

    # Continuation markers: "DONALD (Cont.)"
    cont_match = re.match(r'^([A-Z][A-Z ]+?)\s*\(Cont\.?\)', stripped)
    if cont_match and cont_match.group(1).strip() in known_characters:
        return True

    # Heuristic: ALL CAPS, 1-4 words, not starting with '('
    if stripped.startswith('('):
        return False
    if not re.match(r'^[A-Z][A-Z\s.\'-]+$', stripped):
        return False
    word_count = len(stripped.split())
    if word_count > 5:
        return False
    # Exclude common non-character all-caps lines
    skip = {'ACT', 'SCENE', 'PROLOGUE', 'EPILOGUE', 'SETTING', 'AT RISE',
            'BLACKOUT', 'CURTAIN', 'END', 'THE END', 'LIGHTS',
            'TIME', 'CHARACTERS', 'CAST OF CHARACTERS',
            'LOGLINE', 'SYNOPSIS', 'ACT ONE', 'ACT TWO', 'ACT THREE',
            'ACT I', 'ACT II', 'ACT III'}
    if stripped.rstrip(':') in skip:
        return False
    if re.match(r'^(ACT|SCENE|SETTING|AT RISE|BLACKOUT|CURTAIN|END OF)', stripped, re.IGNORECASE):
        return False
    return True


def _discover_characters(lines):
    """Scan the raw lines to discover character names that appear as headings."""
    candidates = {}
    for i, line in enumerate(lines):
        stripped = line.strip()
        if not stripped:
            continue
        # Must be ALL CAPS, short, not a known structural keyword
        if not re.match(r'^[A-Z][A-Z\s.\'-]+$', stripped):
            continue
        if stripped.startswith('('):
            continue
        word_count = len(stripped.split())
        if word_count > 4:
            continue
        skip = {'ACT', 'SCENE', 'PROLOGUE', 'EPILOGUE', 'SETTING', 'AT RISE',
                'BLACKOUT', 'CURTAIN', 'END', 'THE END', 'LIGHTS',
                'TIME', 'CHARACTERS', 'CAST OF CHARACTERS',
                'LOGLINE', 'SYNOPSIS', 'ACT ONE', 'ACT TWO', 'ACT THREE',
                'ACT I', 'ACT II', 'ACT III', 'BY', 'COPYRIGHT'}
        if stripped.rstrip(':') in skip or stripped.rstrip('.') in skip:
            continue
        if re.match(r'^(ACT|SCENE|SETTING|AT RISE|BLACKOUT|CURTAIN|END OF)', stripped, re.IGNORECASE):
            continue
        # Must be followed (within 2 lines) by non-empty text that is NOT all-caps
        for j in range(i + 1, min(i + 3, len(lines))):
            nxt = lines[j].strip()
            if nxt:
                if not re.match(r'^[A-Z][A-Z\s.\'-]+$', nxt) or nxt.startswith('('):
                    candidates[stripped] = candidates.get(stripped, 0) + 1
                break
    # Keep names that appear at least once (lax for short plays)
    return {name for name, count in candidates.items() if count >= 1}


def parse_script(text):
    """Parse raw script text into structured elements.

    Returns (metadata_dict, elements_list).
    metadata_dict has keys: title, author, description, characters, scene, time, copyright, contact
    elements_list has dicts with 'type' and 'text' (and sometimes 'subtype').
    """
    raw_lines = text.split('\n')

    # --- Phase 1: discover metadata from first ~40 lines ---
    metadata = {
        'title': '',
        'author': '',
        'description': '',
        'characters': [],   # list of (name, description) tuples
        'scene': '',
        'time': '',
        'copyright': '',
        'contact': '',
    }

    # Find where the actual script body starts (first ACT/SCENE/AT RISE after metadata)
    body_start = 0
    skip_lines = set()  # lines consumed by lookahead

    # First pass: locate key metadata sections
    for i, line in enumerate(raw_lines):
        if i in skip_lines:
            continue
        stripped = line.strip()
        if not stripped:
            continue

        # Detect CHARACTERS: section
        if re.match(r'^CHARACTERS?\s*:?\s*$', stripped, re.IGNORECASE) or stripped.upper() in ('CHARACTERS', 'CHARACTERS:', 'CAST OF CHARACTERS'):
            j = i + 1
            while j < len(raw_lines):
                cline = raw_lines[j].strip()
                if not cline:
                    j += 1
                    continue
                # Stop at structural keywords
                if re.match(r'^(ACT|TIME\s*:|SETTING\s*:|AT RISE|PROLOGUE)', cline, re.IGNORECASE):
                    break
                # Character line: NAME – description or NAME, description
                char_match = re.match(r'^([A-Z][A-Z\s.\'-]+?)\s*[\u2013\u2014\-:,]\s*(.+)', cline)
                if char_match:
                    name = char_match.group(1).strip()
                    desc = char_match.group(2).strip()
                    # Gather continuation lines for this character's description
                    j += 1
                    while j < len(raw_lines):
                        cont = raw_lines[j].strip()
                        if not cont:
                            j += 1
                            continue
                        # If next line starts with an uppercase name + dash, it's a new character
                        if re.match(r'^[A-Z][A-Z\s.\'-]+?\s*[\u2013\u2014\-:,]', cont):
                            break
                        if re.match(r'^(ACT|TIME\s*:|SETTING\s*:|AT RISE)', cont, re.IGNORECASE):
                            break
                        desc += ' ' + cont
                        j += 1
                    metadata['characters'].append((name, desc))
                else:
                    j += 1
            continue

        # Detect TIME:
        time_match = re.match(r'^TIME\s*:\s*(.*)', stripped, re.IGNORECASE)
        if time_match:
            val = time_match.group(1).strip()
            # If TIME: and SETTING: got merged on same line
            setting_in_time = re.match(r'^(.+?)\s+SETTING\s*:\s*(.*)', val, re.IGNORECASE)
            if setting_in_time:
                metadata['time'] = setting_in_time.group(1).strip()
                metadata['scene'] = setting_in_time.group(2).strip()
            elif val:
                metadata['time'] = val
            else:
                # TIME: with empty value -- look at next non-empty line
                for j in range(i + 1, min(i + 3, len(raw_lines))):
                    nxt = raw_lines[j].strip()
                    if nxt:
                        # Check if next line has SETTING: merged in
                        merged = re.match(r'^(.+?)\s+SETTING\s*:\s*(.*)', nxt, re.IGNORECASE)
                        if merged:
                            metadata['time'] = merged.group(1).strip()
                            scene_val = merged.group(2).strip()
                            if not scene_val:
                                # Setting value might be on yet another line
                                for k in range(j + 1, min(j + 3, len(raw_lines))):
                                    nxt2 = raw_lines[k].strip()
                                    if nxt2:
                                        scene_val = nxt2
                                        skip_lines.add(k)
                                        break
                            metadata['scene'] = scene_val
                        else:
                            metadata['time'] = nxt
                        skip_lines.add(j)
                        break
            continue

        # Detect Setting as metadata
        setting_meta_match = re.match(r'^Setting\s*:\s*(.*)', stripped, re.IGNORECASE)
        if setting_meta_match and i < 40:
            val = setting_meta_match.group(1).strip()
            if not val:
                for j in range(i + 1, min(i + 3, len(raw_lines))):
                    nxt = raw_lines[j].strip()
                    if nxt:
                        val = nxt
                        skip_lines.add(j)
                        break
            metadata['scene'] = val
            continue

        # Detect ACT / AT RISE as start of body
        # Handle merged lines like "ACT ONE AT RISE:"
        if re.match(r'^ACT\s', stripped, re.IGNORECASE):
            body_start = i
            break
        if re.match(r'^AT RISE', stripped, re.IGNORECASE):
            body_start = i
            break

    # Extract title, author, description from header lines
    header_lines = []
    for line in raw_lines[:min(body_start if body_start > 0 else 20, 40)]:
        stripped = line.strip()
        if not stripped:
            continue
        # Skip already-parsed metadata
        if re.match(r'^(CHARACTERS?\s*:?|TIME\s*:|Setting\s*:|LOGLINE|SYNOPSIS)', stripped, re.IGNORECASE):
            continue
        # Skip contact/copyright info
        if re.match(r'^(\d{3}-\d{3}|Phone:|Fax:|E-mail:|www\.|http|Copyright|\u00a9)', stripped, re.IGNORECASE):
            if 'copyright' in stripped.lower() or '\u00a9' in stripped:
                metadata['copyright'] = stripped
            else:
                metadata['contact'] = (metadata['contact'] + '\n' + stripped).strip()
            continue
        # Skip lines that contain email-like patterns as part of contact/footer
        if re.match(r'^.*@.*\.\w+', stripped) and i < 15:
            metadata['contact'] = (metadata['contact'] + '\n' + stripped).strip()
            continue
        # Skip character entries
        skip_line = False
        for cname, _ in metadata['characters']:
            if cname in stripped:
                skip_line = True
                break
        if skip_line:
            continue
        # Skip long lines (synopsis, logline content)
        if len(stripped) > 120:
            continue
        # Skip page-number-like lines
        if re.match(r'^-?\s*\d+\s*-?$', stripped):
            continue
        # Skip footer/header lines from PDF (but not titles -- those are ALL CAPS single words)
        if re.match(r'^Melissa Jordan Grey', stripped, re.IGNORECASE):
            continue
        if re.match(r'^Bitter\s*cake\s+\S', stripped, re.IGNORECASE):
            # "Bitter cake melissa@..." is a footer, but "BITTERCAKE" alone is a title
            continue
        header_lines.append(stripped)

    # Parse title/author/description from header
    i_h = 0
    while i_h < len(header_lines):
        hl = header_lines[i_h]

        # Title: first line (strip off draft annotations)
        if not metadata['title']:
            title_text = re.sub(r'\s*-\s*[Dd]raft\s*\d*', '', hl).strip()
            # If the line contains "By Author" merged, split it
            by_match = re.match(r'^(.+?)\s+By\s+(.+)', title_text)
            if by_match:
                metadata['title'] = by_match.group(1).strip()
                metadata['author'] = by_match.group(2).strip()
                i_h += 1
                continue
            metadata['title'] = title_text
            i_h += 1
            continue

        # Description line: "A Play in Two Acts", "A 10-minute Play"
        if re.match(r'^[Aa]\s+(\d+.minute|play|drama|comedy|tragicomedy)', hl, re.IGNORECASE):
            metadata['description'] = hl
            i_h += 1
            continue

        # "By" line
        if re.match(r'^[Bb]y\s*$', hl):
            # Author is on next line(s)
            i_h += 1
            author_parts = []
            while i_h < len(header_lines):
                nxt = header_lines[i_h]
                if re.match(r'^(a\s+\d+|a play|copyright|\d{3}-)', nxt, re.IGNORECASE):
                    break
                author_parts.append(nxt)
                i_h += 1
                if len(author_parts) >= 2:
                    break
            metadata['author'] = ' '.join(author_parts)
            continue

        if hl.lower().startswith('by '):
            metadata['author'] = hl[3:].strip()
            i_h += 1
            continue

        i_h += 1

    # --- Phase 2: parse the script body ---
    if body_start == 0:
        # Fallback: scan for first character name or structural element
        for idx, line in enumerate(raw_lines):
            stripped = line.strip()
            if re.match(r'^(ACT\s|AT RISE|SETTING\s*:)', stripped, re.IGNORECASE):
                body_start = idx
                break

    body_lines = raw_lines[body_start:]

    # Discover character names from the body
    known_characters = _discover_characters(body_lines)
    # Also add characters from metadata
    for cname, _ in metadata['characters']:
        known_characters.add(cname.upper())

    elements = []

    i = 0
    while i < len(body_lines):
        line = body_lines[i].strip()

        if not line:
            i += 1
            continue

        # Skip page numbers (standalone digits, -N- patterns, I-1-1 patterns)
        if re.match(r'^-?\s*\d+\s*-?$', line) or re.match(r'^[IV]+-\d+-\d+$', line):
            i += 1
            continue

        # Skip footer lines (author/copyright footers in PDFs)
        if re.match(r'^(Melissa Jordan Grey|melissa@|www\.|Copyright|\u00a9)', line, re.IGNORECASE):
            i += 1
            continue
        if re.match(r'^Bitter\s*cake\s+\S', line, re.IGNORECASE):
            i += 1
            continue

        # Handle merged "ACT ONE AT RISE:" lines
        merged_act = re.match(r'^(ACT\s+(?:ONE|TWO|THREE|FOUR|FIVE|I{1,3}V?|V?I{0,3}))\s+(AT RISE)\s*:\s*(.*)', line, re.IGNORECASE)
        if merged_act:
            elements.append({'type': 'act_heading', 'text': merged_act.group(1).upper()})
            desc = merged_act.group(3).strip()
            i += 1
            while i < len(body_lines):
                nxt = body_lines[i].strip()
                if not nxt:
                    break
                if _is_character_name(nxt, known_characters):
                    break
                if re.match(r'^(ACT\s|SCENE\s|SETTING)', nxt, re.IGNORECASE):
                    break
                desc += ' ' + nxt
                i += 1
            elements.append({'type': 'setting', 'subtype': 'AT RISE', 'text': desc.strip()})
            continue

        # ACT heading (standalone)
        act_match = re.match(r'^(ACT\s+(?:ONE|TWO|THREE|FOUR|FIVE|I{1,3}V?|V?I{0,3}))\s*$', line, re.IGNORECASE)
        if act_match:
            elements.append({'type': 'act_heading', 'text': line.upper()})
            i += 1
            continue

        # SCENE heading
        scene_match = re.match(r'^(Scene\s+\d+|Scene\s+\w+)\s*$', line, re.IGNORECASE)
        if scene_match:
            elements.append({'type': 'scene_heading', 'text': line})
            i += 1
            continue

        # SETTING: line
        setting_match = re.match(r'^(SETTING)\s*:\s*(.*)', line, re.IGNORECASE)
        if setting_match:
            # Gather continuation lines (setting can span multiple lines)
            desc = setting_match.group(2).strip()
            i += 1
            while i < len(body_lines):
                nxt = body_lines[i].strip()
                if not nxt or re.match(r'^(AT RISE|ACT\s|SCENE\s)', nxt, re.IGNORECASE):
                    break
                if _is_character_name(nxt, known_characters):
                    break
                desc += ' ' + nxt
                i += 1
            elements.append({'type': 'setting', 'subtype': 'SETTING', 'text': desc.strip()})
            continue

        # AT RISE: line
        atrise_match = re.match(r'^(AT RISE)\s*:\s*(.*)', line, re.IGNORECASE)
        if atrise_match:
            desc = atrise_match.group(2).strip()
            i += 1
            while i < len(body_lines):
                nxt = body_lines[i].strip()
                if not nxt:
                    break
                if _is_character_name(nxt, known_characters):
                    break
                if re.match(r'^(ACT\s|SCENE\s|SETTING)', nxt, re.IGNORECASE):
                    break
                desc += ' ' + nxt
                i += 1
            elements.append({'type': 'setting', 'subtype': 'AT RISE', 'text': desc.strip()})
            continue

        # BLACKOUT / CURTAIN / END OF SCENE / END OF ACT / THE END
        if re.match(r'^\(?(BLACKOUT|CURTAIN|END OF SCENE|END OF ACT|THE END)\)?\.?\s*$', line, re.IGNORECASE):
            clean = line.strip('() .')
            elements.append({'type': 'end_designation', 'text': clean.upper()})
            i += 1
            continue

        # "Lights fade." type ending
        if re.match(r'^Lights?\s+(fade|dim|down|out)', line, re.IGNORECASE):
            elements.append({'type': 'stage_direction', 'text': line})
            i += 1
            continue

        # THE END
        if re.match(r'^THE END\.?\s*$', line, re.IGNORECASE):
            elements.append({'type': 'end_designation', 'text': 'THE END'})
            i += 1
            continue

        # Character name line
        if _is_character_name(line, known_characters):
            elements.append({'type': 'character', 'text': line.upper()})
            i += 1
            continue

        # Standalone stage direction in parentheses
        if line.startswith('(') and line.endswith(')'):
            elements.append({'type': 'stage_direction', 'text': line})
            i += 1
            continue

        # Multi-line parenthetical stage direction
        if line.startswith('(') and ')' not in line:
            full = line
            i += 1
            while i < len(body_lines):
                nxt = body_lines[i].strip()
                if not nxt:
                    i += 1
                    continue
                full += ' ' + nxt
                i += 1
                if ')' in nxt:
                    break
            elements.append({'type': 'stage_direction', 'text': full})
            continue

        # Standalone "Beat." or "Beat" as stage direction
        if re.match(r'^Beat\.?\s*$', line, re.IGNORECASE):
            elements.append({'type': 'stage_direction', 'text': '(Beat.)'})
            i += 1
            continue

        # Unparenthesized stage directions: only if they match strong patterns
        if (len(elements) > 0 and _looks_like_stage_direction(line) and
            len(line) < 100):
            elements.append({'type': 'stage_direction', 'text': f'({line})'})
            i += 1
            continue

        # Default: dialogue
        # Gather continuation lines (dialogue that wraps across multiple raw lines)
        dialogue = line
        i += 1
        while i < len(body_lines):
            nxt = body_lines[i].strip()
            if not nxt:
                break
            if _is_character_name(nxt, known_characters):
                break
            if re.match(r'^(ACT\s|SCENE\s|SETTING|AT RISE)', nxt, re.IGNORECASE):
                break
            if re.match(r'^\(?(BLACKOUT|CURTAIN|END OF SCENE|END OF ACT|THE END)\)?', nxt, re.IGNORECASE):
                break
            if re.match(r'^Beat\.?\s*$', nxt, re.IGNORECASE):
                break
            if nxt.startswith('(') and (nxt.endswith(')') or ')' not in nxt):
                break
            # If the next line is a known stage direction indicator, stop
            if _looks_like_stage_direction(nxt) and len(nxt) < 60:
                break
            dialogue += ' ' + nxt
            i += 1

        elements.append({'type': 'dialogue', 'text': dialogue})

    return metadata, elements


def _looks_like_stage_direction(line):
    """Heuristic: does this line look like an unparenthesized stage direction?"""
    # Already in parens
    if line.startswith('('):
        return True
    # Common stage direction patterns
    sd_patterns = [
        r'^(Yells?|Shouts?|Whispers?|Laughs?|Laughing|Smil(?:es?|ing)|Crying)',
        r'^(Pauses?|Pausing|Looking|Shrugs?|Shrugging|Squints?)',
        r'^(They both|She |He |ANGELA |KATHRYN |Both )',
        r'^(Neither speaks|Looks? (?:at|down|up)|Grabs? |Starts? |Stops? )',
        r'^(to (?:herself|himself|ANGELA|KATHRYN|BARETT|field|phone|DERMOTT))',
        r'^(into phone|off)',
        r'^(waving|clapping|trotting|jogging)',
        r'^(whistle is blown)',
        r'(?:looks? at|cringe|sit again|stand|enter|exit|cross)',
    ]
    for pat in sd_patterns:
        if re.match(pat, line, re.IGNORECASE):
            return True
    return False


# ---------------------------------------------------------------------------
# DOCX GENERATION -- Standard Stage Play Format
# ---------------------------------------------------------------------------

COURIER_FONT = 'Courier New'
FONT_SIZE = Pt(12)

# Margins per CPF guide
TITLE_PAGE_TOP = Inches(3.5)
TITLE_PAGE_LEFT = Inches(4)
TEXT_PAGE_TOP = Inches(1)
TEXT_PAGE_LEFT = Inches(1.5)
TEXT_PAGE_RIGHT = Inches(1)
TEXT_PAGE_BOTTOM = Inches(1)

# Indentation (from left edge of page, but since left margin is 1.5",
# we subtract that to get the paragraph indent)
# Character names: 4" from left edge -> 4 - 1.5 = 2.5" indent
CHAR_NAME_INDENT = Inches(2.5)
# Stage directions: 2.75" from left edge -> 2.75 - 1.5 = 1.25" indent
STAGE_DIR_INDENT = Inches(1.25)
# Stage direction right indent to keep ~2.5" wide: right margin from page edge = 1",
# so usable width = 8.5 - 1.5 - 1 = 6". Stage dir starts at 1.25" indent,
# so to be ~2.5" wide, right indent = 6 - 1.25 - 2.5 = 2.25"
STAGE_DIR_RIGHT_INDENT = Inches(2.25)
# Act/Scene headings: 4" from left edge -> 2.5" indent (same as char names)
HEADING_INDENT = Inches(2.5)
# SETTING/AT RISE labels at left margin, description at 4" from left -> 2.5" tab
SETTING_TAB = Inches(2.5)


def _set_run_font(run, bold=False, italic=False, underline=False):
    """Apply Courier 12pt to a run."""
    run.font.name = COURIER_FONT
    run.font.size = FONT_SIZE
    run.bold = bold
    run.italic = italic
    run.underline = underline


def _add_empty_paragraphs(doc, count):
    """Add empty paragraphs for spacing."""
    for _ in range(count):
        p = doc.add_paragraph()
        run = p.add_run()
        _set_run_font(run)


def _add_page_number_field(paragraph):
    """Insert a PAGE field code into a paragraph."""
    run = paragraph.add_run()
    fldChar1 = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>')
    run._r.append(fldChar1)

    run2 = paragraph.add_run()
    instrText = parse_xml(f'<w:instrText {nsdecls("w")} xml:space="preserve"> PAGE </w:instrText>')
    run2._r.append(instrText)

    run3 = paragraph.add_run()
    fldChar2 = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>')
    run3._r.append(fldChar2)


def create_formatted_docx(title, author, description, characters, scene, time_str,
                          copyright_info, contact_info, elements):
    """Create a Word document formatted per the standard stage play format."""

    doc = Document()

    # Set default style to Courier 12pt
    style = doc.styles['Normal']
    style.font.name = COURIER_FONT
    style.font.size = FONT_SIZE
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(0)
    style.paragraph_format.line_spacing = 1.0

    # ===================================================================
    # TITLE PAGE
    # ===================================================================
    section = doc.sections[0]
    section.top_margin = TITLE_PAGE_TOP
    section.left_margin = TITLE_PAGE_LEFT
    section.right_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.different_first_page_header_footer = True

    # Title in ALL CAPS
    p = doc.add_paragraph()
    run = p.add_run(title.upper())
    _set_run_font(run, underline=True)

    # Underscore line (same width as title) -- blank line then description
    _add_empty_paragraphs(doc, 1)

    # Description line (e.g., "A Play in Two Acts")
    if description:
        p = doc.add_paragraph()
        run = p.add_run(description)
        _set_run_font(run)

        _add_empty_paragraphs(doc, 1)

    # "by"
    p = doc.add_paragraph()
    run = p.add_run("by")
    _set_run_font(run)

    _add_empty_paragraphs(doc, 1)

    # Author name
    p = doc.add_paragraph()
    run = p.add_run(author)
    _set_run_font(run)

    # Copyright (lower left) and Contact (lower right) -- use a table at the bottom
    # We'll add enough spacing to push content toward the bottom
    _add_empty_paragraphs(doc, 10)

    if copyright_info or contact_info:
        # Create a two-column invisible table for bottom info
        table = doc.add_table(rows=1, cols=2)
        table.autofit = True
        # Left cell: copyright
        left_cell = table.cell(0, 0)
        left_cell.text = ''
        p = left_cell.paragraphs[0]
        run = p.add_run(copyright_info if copyright_info else '')
        _set_run_font(run)
        # Right cell: contact
        right_cell = table.cell(0, 1)
        right_cell.text = ''
        p = right_cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run(contact_info if contact_info else '')
        _set_run_font(run)
        # Remove table borders
        for cell in [left_cell, right_cell]:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcBorders = parse_xml(
                f'<w:tcBorders {nsdecls("w")}>'
                '  <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                '  <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                '  <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                '  <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                '</w:tcBorders>'
            )
            tcPr.append(tcBorders)

    # ===================================================================
    # CAST OF CHARACTERS PAGE
    # ===================================================================
    doc.add_page_break()

    # New section with text-page margins
    new_section = doc.add_section()
    new_section.top_margin = TEXT_PAGE_TOP
    new_section.left_margin = TEXT_PAGE_LEFT
    new_section.right_margin = TEXT_PAGE_RIGHT
    new_section.bottom_margin = TEXT_PAGE_BOTTOM
    new_section.different_first_page_header_footer = True

    # "Cast of Characters" centered and underlined
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Cast of Characters")
    _set_run_font(run, underline=True)

    _add_empty_paragraphs(doc, 1)

    # Character entries
    if characters:
        for char_name, char_desc in characters:
            p = doc.add_paragraph()
            # Name (underlined) followed by colon, then tab, then description
            run = p.add_run(f"{char_name}:")
            _set_run_font(run, underline=True)
            run = p.add_run(f"\t{char_desc}")
            _set_run_font(run)
            # Add tab stop at the setting position
            tab_stops = p.paragraph_format.tab_stops
            tab_stops.add_tab_stop(SETTING_TAB)
            _add_empty_paragraphs(doc, 1)

    # Scene and Time below character list
    if scene:
        _add_empty_paragraphs(doc, 2)
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run("Scene")
        _set_run_font(run, underline=True)
        _add_empty_paragraphs(doc, 1)
        p = doc.add_paragraph()
        run = p.add_run(scene)
        _set_run_font(run)

    if time_str:
        _add_empty_paragraphs(doc, 1)
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run("Time")
        _set_run_font(run, underline=True)
        _add_empty_paragraphs(doc, 1)
        p = doc.add_paragraph()
        run = p.add_run(time_str)
        _set_run_font(run)

    # ===================================================================
    # SCRIPT BODY PAGES
    # ===================================================================
    doc.add_page_break()

    # Add page number in header for body section
    header = new_section.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _add_page_number_field(hp)
    for run in hp.runs:
        _set_run_font(run)

    # Process elements
    for elem in elements:
        etype = elem['type']
        text = elem['text']

        if etype == 'act_heading':
            # New page for each act
            doc.add_page_break()
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = HEADING_INDENT
            p.space_before = Pt(0)
            run = p.add_run(text.upper())
            _set_run_font(run, underline=True)

        elif etype == 'scene_heading':
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = HEADING_INDENT
            p.space_before = Pt(12)
            run = p.add_run(text)
            _set_run_font(run, underline=True)

        elif etype == 'setting':
            # Two-column: label at left margin, description indented
            p = doc.add_paragraph()
            p.space_before = Pt(12)
            label = elem.get('subtype', 'SETTING')
            run = p.add_run(f"{label.upper()}:")
            _set_run_font(run)
            run = p.add_run(f"\t{text}")
            _set_run_font(run)
            tab_stops = p.paragraph_format.tab_stops
            tab_stops.add_tab_stop(SETTING_TAB)
            # Hanging indent so wrapped lines align with the tab
            p.paragraph_format.left_indent = SETTING_TAB
            p.paragraph_format.first_line_indent = -SETTING_TAB

        elif etype == 'character':
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = CHAR_NAME_INDENT
            p.space_before = Pt(12)
            p.space_after = Pt(0)
            run = p.add_run(text.upper())
            _set_run_font(run)

        elif etype == 'dialogue':
            p = doc.add_paragraph()
            p.space_before = Pt(0)
            p.space_after = Pt(0)
            # Dialogue may contain inline stage directions in parentheses or italics
            # Split on parenthetical segments
            pattern = r'(\([^)]+\))'
            segments = re.split(pattern, text)
            for segment in segments:
                if not segment:
                    continue
                if segment.startswith('(') and segment.endswith(')'):
                    # Inline stage direction
                    run = p.add_run(segment)
                    _set_run_font(run, italic=True)
                else:
                    # Check for *(italicized text)* patterns too
                    italic_pattern = r'(\*[^*]+\*)'
                    sub_segments = re.split(italic_pattern, segment)
                    for sub in sub_segments:
                        if not sub:
                            continue
                        if sub.startswith('*') and sub.endswith('*'):
                            run = p.add_run(sub.strip('*'))
                            _set_run_font(run, italic=True)
                        else:
                            run = p.add_run(sub)
                            _set_run_font(run)

        elif etype == 'stage_direction':
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = STAGE_DIR_INDENT
            p.paragraph_format.right_indent = STAGE_DIR_RIGHT_INDENT
            p.space_before = Pt(0)
            p.space_after = Pt(0)
            # Ensure text is in parentheses
            display_text = text.strip()
            if not display_text.startswith('('):
                display_text = f'({display_text})'
            if not display_text.endswith(')'):
                display_text = f'{display_text})'
            run = p.add_run(display_text)
            _set_run_font(run, italic=False)

        elif etype == 'end_designation':
            # Indented, ALL CAPS, in parentheses
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = CHAR_NAME_INDENT
            p.space_before = Pt(12)
            display = text.upper()
            if not display.startswith('('):
                display = f'({display})'
            if not display.endswith(')'):
                display = f'{display})'
            run = p.add_run(display)
            _set_run_font(run)

    # Save
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


# ---------------------------------------------------------------------------
# STREAMLIT UI
# ---------------------------------------------------------------------------

st.header("Reformat Your Script in 3 Easy Steps")

# STEP 1: Upload
st.subheader("Step 1: Upload Your Script File")

uploaded_file = st.file_uploader(
    "Choose your script file (PDF, DOCX, or TXT)",
    type=['pdf', 'docx', 'txt'],
    help="Upload a PDF, Word document, or text file"
)

if uploaded_file:
    st.success(f"File uploaded: {uploaded_file.name}")

    # Extract text
    with st.spinner("Reading your script..."):
        if uploaded_file.name.endswith('.pdf'):
            script_text = extract_text_from_pdf(uploaded_file)
        elif uploaded_file.name.endswith('.docx'):
            script_text = extract_text_from_docx(uploaded_file)
        elif uploaded_file.name.endswith('.txt'):
            script_text = extract_text_from_txt(uploaded_file)
        else:
            st.error("Unsupported file type")
            script_text = None

    if script_text:
        st.success("Script loaded successfully!")

        with st.expander("View Original Text (First 1000 characters)"):
            st.text(script_text[:1000] + "..." if len(script_text) > 1000 else script_text)

        # Parse
        with st.spinner("Analyzing script structure..."):
            metadata, elements = parse_script(script_text)

        char_count = sum(1 for e in elements if e['type'] == 'character')
        dialogue_count = sum(1 for e in elements if e['type'] == 'dialogue')
        sd_count = sum(1 for e in elements if e['type'] == 'stage_direction')
        st.info(f"Detected {len(elements)} elements: {char_count} character cues, {dialogue_count} dialogue blocks, {sd_count} stage directions")

        # STEP 2: Enter details
        st.divider()
        st.subheader("Step 2: Verify Title, Author, and Script Details")

        col1, col2 = st.columns(2)
        with col1:
            title = st.text_input("Play Title*", value=metadata['title'] or "", placeholder="e.g., HAMLET")
        with col2:
            author = st.text_input("Author*", value=metadata['author'] or "", placeholder="e.g., William Shakespeare")

        description = st.text_input(
            "Description (e.g., 'A Play in Two Acts', 'A 10-minute Play')",
            value=metadata['description'] or "",
            placeholder="A Play in Two Acts"
        )

        # Characters
        st.markdown("**Cast of Characters**")
        st.caption("Edit the character list below. One per line: NAME - Description")

        default_chars = ""
        if metadata['characters']:
            default_chars = "\n".join(f"{n} - {d}" for n, d in metadata['characters'])

        chars_text = st.text_area("Characters", value=default_chars, height=150,
                                   placeholder="HAMLET - Prince of Denmark\nOPHELIA - Daughter of Polonius")

        # Parse characters from text area
        parsed_characters = []
        if chars_text.strip():
            for cline in chars_text.strip().split('\n'):
                cline = cline.strip()
                if not cline:
                    continue
                match = re.match(r'^(.+?)\s*[\u2013\u2014\-:]\s*(.+)', cline)
                if match:
                    parsed_characters.append((match.group(1).strip(), match.group(2).strip()))
                else:
                    parsed_characters.append((cline, ''))

        col3, col4 = st.columns(2)
        with col3:
            scene_val = st.text_input("Scene / Setting", value=metadata['scene'] or "",
                                       placeholder="A living room in Brooklyn")
        with col4:
            time_val = st.text_input("Time", value=metadata['time'] or "",
                                      placeholder="The present.")

        col5, col6 = st.columns(2)
        with col5:
            copyright_val = st.text_input("Copyright (optional)", value=metadata['copyright'] or "",
                                           placeholder="Copyright \u00a9 2025 by Author Name")
        with col6:
            contact_val = st.text_input("Contact Info (optional)", value=metadata['contact'] or "",
                                         placeholder="email@example.com")

        # STEP 3: Download
        st.divider()
        st.subheader("Step 3: Download Your Formatted Script")

        if not title or not author:
            st.warning("Please enter both Title and Author above to continue")
        else:
            filename = st.text_input(
                "Filename (optional - will use title if blank)",
                value="",
                placeholder=title.replace(' ', '_').lower() if title else "my_play"
            )
            if not filename:
                filename = title.replace(' ', '_').lower() if title else "formatted_script"

            if st.button("Format & Download as .DOCX", type="primary", use_container_width=True):
                with st.spinner("Creating professionally formatted document..."):
                    formatted_doc = create_formatted_docx(
                        title=title,
                        author=author,
                        description=description,
                        characters=parsed_characters,
                        scene=scene_val,
                        time_str=time_val,
                        copyright_info=copyright_val,
                        contact_info=contact_val,
                        elements=elements
                    )

                st.success("Document created!")

                st.download_button(
                    label="Save Word Document",
                    data=formatted_doc,
                    file_name=f"{filename}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

            st.divider()
            st.success("""
            **Your script will be formatted with:**
            - Professional title page (title, underscore, description, author)
            - Cast of Characters page with Scene and Time
            - Courier 12pt font throughout
            - Character names in ALL CAPS, indented 4" from left edge
            - Dialogue runs full width (left margin to right margin)
            - Stage directions indented 2.75" from left edge, in parentheses
            - SETTING/AT RISE in two-column layout
            - Act/Scene headings indented and underscored
            - BLACKOUT/END designations properly placed
            - Page numbers in upper right corner
            """)

else:
    st.info("Upload a script file above to get started")

    st.markdown("""
    ### Supported Formats:
    - **PDF** (.pdf) - Most common
    - **Word** (.docx) - Microsoft Word documents
    - **Text** (.txt) - Plain text files

    ### What We'll Fix:
    1. Add professional title page per standard format
    2. Add Cast of Characters page
    3. Character names in ALL CAPS, indented
    4. Dialogue full-width, left-aligned
    5. Stage directions indented in parentheses
    6. Courier 12pt font throughout
    7. Proper margins (1.5" left, 1" top/right/bottom)
    8. Page numbers in upper right corner
    """)

# Footer
st.divider()
st.markdown("""
<div style='text-align: center; color: #666; font-size: 12px; margin-top: 40px;'>
    <p>The TWNC FaloopinFormatter | Professional Script Reformatting</p>
    <p>Formatted per the Standard Stage Play Format (CPF)</p>
</div>
""", unsafe_allow_html=True)
