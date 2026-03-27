"""
coa_mapper.py
=============
Phase 1 rules-first COA mapper for the owner financial extractor.

Takes source account labels from the EXR Rolling IS and maps them to the
standard P-Builder Chart of Accounts using a five-step pipeline:

  Step 1 — Exact approved match   (confidence 1.00)
            Verbatim string match against approved_mappings.csv.

  Step 2 — Normalized exact match (confidence 0.95)
            Strips GL account codes like (4000) and lowercases before
            comparing. Handles minor formatting differences between EXR
            file versions.

  Step 3 — Alias match            (confidence 0.85)
            Matches against known alternate label names in
            alias_mappings.csv.  Add rows there to cover new variations
            without touching this file.

  Step 4 — Fuzzy match            (confidence 0.50–0.84)
            difflib SequenceMatcher on normalized strings.
            Always flagged for review regardless of score.

  Step 5 — No match               (confidence 0.00)
            Needs a manual entry in approved_mappings.csv or
            alias_mappings.csv.

Each result dict contains:
  source_label    — the original label as extracted from Rolling IS
  coa             — suggested P-Builder COA line item
  coa2            — suggested rollup category (Net Rental Income, etc.)
  account_type    — Income / Expense / EXR_Rollup
  confidence      — float 0.0 – 1.0
  match_method    — one of the METHOD_* constants below
  review_required — True if the mapping needs human confirmation
  notes           — plain-English explanation for the audit trail

HOW TO MAINTAIN THIS SYSTEM:
  - To add or correct a known mapping:  edit approved_mappings.csv
  - To add an alternate spelling:       edit alias_mappings.csv
  - No code changes are needed for routine updates.
"""

import csv
import os
import re
from difflib import SequenceMatcher


# ---------------------------------------------------------------------------
# PATHS — CSV files live in the same directory as this module
# ---------------------------------------------------------------------------

_HERE = os.path.dirname(os.path.abspath(__file__))
APPROVED_MAPPINGS_FILE = os.path.join(_HERE, "approved_mappings_exr.csv")
ALIAS_MAPPINGS_FILE    = os.path.join(_HERE, "alias_mappings_exr.csv")


# ---------------------------------------------------------------------------
# CONFIDENCE THRESHOLDS
# ---------------------------------------------------------------------------

# Mappings at or above this score are auto-accepted (review_required = False)
# Steps 1, 2, 3 all meet or exceed this threshold.
CONFIDENCE_AUTO_ACCEPT = 0.85

# Fuzzy matches below this raw difflib score are treated as no-match
CONFIDENCE_FUZZY_MIN = 0.50


# ---------------------------------------------------------------------------
# MATCH METHOD LABELS
# Stored in each result for full auditability in the COA Mapping review tab.
# ---------------------------------------------------------------------------

METHOD_EXACT      = "exact_approved"   # verbatim match in approved CSV
METHOD_NORMALIZED = "normalized"       # match after stripping GL codes + lowercasing
METHOD_ALIAS      = "alias"            # match via alias_mappings.csv
METHOD_FUZZY      = "fuzzy"            # difflib similarity
METHOD_NONE       = "no_match"         # nothing found


# ---------------------------------------------------------------------------
# LABEL NORMALIZATION
# ---------------------------------------------------------------------------

def normalize_label(text):
    """
    Prepare a label string for comparison by removing superficial differences:
      - Lowercase
      - Strip leading/trailing whitespace
      - Remove GL account codes in parentheses: (4000), (5100), (5100/5090)
      - Collapse internal whitespace to a single space

    Examples:
      'Rental Income (4000)'         -> 'rental income'
      'Management Fee - ESMI (5100)' -> 'management fee - esmi'
      '  Late Fees  '                -> 'late fees'
      'Payroll Tax (5090)'           -> 'payroll tax'
    """
    if text is None:
        return ""
    s = str(text).strip().lower()
    # Remove parenthetical GL codes: numbers and slashes only, e.g. (4000) or (5100/5090)
    s = re.sub(r'\s*\([0-9][0-9/\s]*\)', '', s)
    # Collapse whitespace
    s = re.sub(r'\s+', ' ', s).strip()
    return s


# ---------------------------------------------------------------------------
# LOADING THE MAPPING TABLES
# ---------------------------------------------------------------------------

def load_approved_mappings(filepath):
    """
    Read approved_mappings.csv and build two lookup dicts.

    Returns a tuple (exact, normalized) where:
      exact      — {original_source_label: entry}
      normalized — {normalize_label(source_label): entry}
                   First entry wins when two labels normalize identically.

    Each entry is a dict:
      {source_label, coa, coa2, account_type, notes}
    """
    exact      = {}
    normalized = {}

    if not os.path.exists(filepath):
        return exact, normalized

    with open(filepath, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            label = row['source_label'].strip()
            entry = {
                'source_label': label,
                'coa':          row['coa'].strip(),
                'coa2':         row['coa2'].strip(),
                'account_type': row['account_type'].strip(),
                'notes':        row.get('notes', '').strip(),
            }
            exact[label] = entry
            norm_key = normalize_label(label)
            if norm_key and norm_key not in normalized:
                normalized[norm_key] = entry

    return exact, normalized


def load_alias_mappings(filepath, approved_exact):
    """
    Read alias_mappings.csv and build a lookup dict.

    Each alias row points to a canonical_label that must exist in
    approved_exact.  Returns:
      {normalize_label(alias): entry_from_approved_exact}

    The entry is a shallow copy with an updated notes field explaining
    that an alias was used.
    """
    aliases = {}

    if not os.path.exists(filepath):
        return aliases

    with open(filepath, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            alias     = row['alias'].strip()
            canonical = row['canonical_label'].strip()
            if canonical not in approved_exact:
                continue   # broken alias — skip silently
            entry = dict(approved_exact[canonical])   # shallow copy
            # Prepend alias explanation to the notes field
            alias_note = f"Alias match: '{alias}' -> '{canonical}'"
            original   = entry.get('notes', '')
            entry['notes'] = (alias_note + ' | ' + original) if original else alias_note
            aliases[normalize_label(alias)] = entry

    return aliases


# ---------------------------------------------------------------------------
# INTERNAL HELPERS
# ---------------------------------------------------------------------------

def _make_result(source_label, entry, confidence, method):
    """
    Build the standard result dict from a matched mapping entry.

    review_required is True when:
      - confidence is below CONFIDENCE_AUTO_ACCEPT, OR
      - account_type is EXR_Rollup (source subtotals the analyst must
        decide whether to exclude to avoid double-counting in the model)
    """
    is_rollup = entry.get('account_type', '').upper() == 'EXR_ROLLUP'
    review    = (confidence < CONFIDENCE_AUTO_ACCEPT) or is_rollup

    notes = entry.get('notes', '')
    if is_rollup and 'do not aggregate' not in notes.lower():
        suffix = 'EXR-calculated subtotal — verify no double-count in model'
        notes  = (notes + ' | ' + suffix) if notes else suffix

    return {
        'source_label':    source_label,
        'coa':             entry.get('coa', ''),
        'coa2':            entry.get('coa2', ''),
        'account_type':    entry.get('account_type', ''),
        'confidence':      round(confidence, 4),
        'match_method':    method,
        'review_required': review,
        'notes':           notes,
    }


def _no_match_result(source_label):
    """Return a result dict indicating no mapping was found."""
    return {
        'source_label':    source_label,
        'coa':             '',
        'coa2':            '',
        'account_type':    '',
        'confidence':      0.0,
        'match_method':    METHOD_NONE,
        'review_required': True,
        'notes':           (
            'No mapping found — add a row to approved_mappings.csv '
            'or alias_mappings.csv to resolve this account.'
        ),
    }


# ---------------------------------------------------------------------------
# FIVE-STEP MAPPING PIPELINE
# ---------------------------------------------------------------------------

def map_label(source_label, approved_exact, approved_normalized, aliases):
    """
    Run a single source label through the five-step pipeline.

    Returns a result dict — see _make_result() for field definitions.
    """
    if not source_label or not str(source_label).strip():
        return _no_match_result(source_label or '')

    label_str = str(source_label).strip()

    # -----------------------------------------------------------------------
    # Step 1: Exact approved match
    # Fastest path — verbatim string comparison against approved_mappings.csv.
    # Handles the vast majority of known EXR account labels.
    # -----------------------------------------------------------------------
    if label_str in approved_exact:
        return _make_result(label_str, approved_exact[label_str],
                            1.00, METHOD_EXACT)

    # -----------------------------------------------------------------------
    # Step 2: Normalized exact match
    # Strips GL codes like (4000) and lowercases.
    # Example: 'rental income (4000)' matches approved label 'Rental Income (4000)'
    # even if EXR changes its GL code display in a future file version.
    # -----------------------------------------------------------------------
    norm = normalize_label(label_str)
    if norm in approved_normalized:
        entry  = approved_normalized[norm]
        result = _make_result(label_str, entry, 0.95, METHOD_NORMALIZED)
        prefix = f"Normalized match for '{entry['source_label']}'"
        result['notes'] = (prefix + ' | ' + result['notes']
                           if result['notes'] else prefix)
        return result

    # -----------------------------------------------------------------------
    # Step 3: Alias match
    # Checks against alternate label names maintained in alias_mappings.csv.
    # Edit that file to add new variations; no code changes needed.
    # -----------------------------------------------------------------------
    if norm in aliases:
        return _make_result(label_str, aliases[norm], 0.85, METHOD_ALIAS)

    # -----------------------------------------------------------------------
    # Step 4: Fuzzy match
    # Uses difflib SequenceMatcher on normalized strings.
    # Confidence = raw_score * 0.90 so even a perfect fuzzy match (1.0)
    # yields 0.90, which is below CONFIDENCE_AUTO_ACCEPT only if the label
    # is essentially identical to a known one. ALL fuzzy matches are
    # force-flagged for review because false positives in financial
    # account mapping are worse than reviewing an extra row.
    # -----------------------------------------------------------------------
    best_score = 0.0
    best_entry = None

    for key, entry in approved_normalized.items():
        score = SequenceMatcher(None, norm, key).ratio()
        if score > best_score:
            best_score = score
            best_entry = entry

    if best_score >= CONFIDENCE_FUZZY_MIN and best_entry is not None:
        confidence   = round(best_score * 0.90, 4)
        result       = _make_result(label_str, best_entry, confidence, METHOD_FUZZY)
        result['review_required'] = True   # always review fuzzy matches
        result['notes'] = (
            f"Fuzzy match ({best_score:.0%} similarity) against "
            f"'{best_entry['source_label']}' — confirm this is correct"
        )
        return result

    # -----------------------------------------------------------------------
    # Step 5: No match
    # -----------------------------------------------------------------------
    return _no_match_result(label_str)


# ---------------------------------------------------------------------------
# PUBLIC MAPPER CLASS
# ---------------------------------------------------------------------------

class COAMapper:
    """
    Load the mapping tables once and map many labels efficiently.

    Typical usage:
        mapper = COAMapper()
        result  = mapper.map("Rental Income (4000)")
        lookup  = mapper.map_unique_from_rows(rolling_is_rows)

    The mapper caches results within the instance to avoid remapping
    the same label multiple times in one run.
    """

    def __init__(self,
                 approved_file=APPROVED_MAPPINGS_FILE,
                 alias_file=ALIAS_MAPPINGS_FILE):
        self.approved_exact, self.approved_normalized = load_approved_mappings(approved_file)
        self.aliases = load_alias_mappings(alias_file, self.approved_exact)
        self._cache  = {}

    def map(self, source_label):
        """
        Map a single source label.
        Results are cached within this instance for the current run.
        """
        key = str(source_label).strip() if source_label else ''
        if key not in self._cache:
            self._cache[key] = map_label(
                key,
                self.approved_exact,
                self.approved_normalized,
                self.aliases,
            )
        return self._cache[key]

    def map_unique_from_rows(self, rows):
        """
        Map every unique label in a list of Rolling IS row dicts
        (each dict must have a 'label' key).

        Preserves the order labels first appear in the data.
        Returns a dict: {source_label: result}
        """
        seen    = set()
        results = {}
        for row_data in rows:
            lbl = row_data.get('label', '')
            if lbl and lbl not in seen:
                seen.add(lbl)
                results[lbl] = self.map(lbl)
        return results

    @property
    def loaded(self):
        """True if the approved mappings file was found and loaded."""
        return bool(self.approved_exact)

    def print_summary(self):
        """Print a brief summary of what was loaded."""
        print(f"    Approved mappings : {len(self.approved_exact)}")
        print(f"    Alias mappings    : {len(self.aliases)}")
