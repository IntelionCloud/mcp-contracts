"""
Contract structure model — parse MD into addressable clauses.
Allows finding clauses by reference (e.g. "4.2") and extracting summaries.
"""
import re


def _text_after_accept(text: str) -> str:
    """Return text as it would appear after accepting all changes."""
    result = re.sub(r'\{--.*?--\}', '', text)
    result = re.sub(r'\{\+\+(.*?)\+\+\}', r'\1', result)
    result = re.sub(r'\[\^\d+\]', '', result)
    return result.strip()


def _is_uppercase_text(text: str) -> bool:
    clean = re.sub(r'^\d+\.\s*', '', text.strip())
    clean = re.sub(r'\{\+\+.*?\+\+\}|\{--.*?--\}', '', clean)
    alpha = [c for c in clean if c.isalpha()]
    return len(alpha) > 2 and all(c.isupper() for c in alpha)


def _text_ends_with_colon(text: str) -> bool:
    clean = re.sub(r'\{\+\+.*?\+\+\}|\{--.*?--\}|\[\^\d+\]', '', text)
    return clean.rstrip().endswith(':')


class Clause:
    """A single clause/sub-clause with its text and metadata."""
    __slots__ = ('ref', 'level', 'text', 'children', 'line_num')

    def __init__(self, ref: str, level: int, text: str, line_num: int = 0):
        self.ref = ref        # e.g. "4.2" or "4.2.1"
        self.level = level    # 0=chapter, 1=clause, 2=subclause
        self.text = text
        self.children = []
        self.line_num = line_num

    def clean_text(self) -> str:
        return _text_after_accept(self.text)

    def __repr__(self):
        preview = self.clean_text()[:50]
        return f"Clause({self.ref}: {preview}...)"


def parse_contract(md_text: str) -> list[Clause]:
    """Parse MD text into a flat list of Clause objects with correct refs.

    Returns a flat list where each clause has a ref like "1.", "1.1.", "1.1.1."
    """
    lines = md_text.split('\n')
    clauses = []

    chapter_num = 0
    clause_num = 0
    subclause_num = 0
    next_chapter = 1
    last_clause_number = 0
    parent_ends_with_colon = False
    in_subclause_block = False

    for i, line in enumerate(lines):
        stripped = line.strip()
        if not stripped or stripped == '---':
            continue
        if stripped.startswith('|') or stripped.startswith('[^') or stripped.startswith('**'):
            continue

        # Clean ## artifacts
        stripped = re.sub(r'^(\d+\.)\s*##\s*', r'\1 ', stripped)

        num_match = re.match(r'^(\d+)\.\s+(.+)$', stripped)
        if not num_match:
            if len(stripped) > 30:
                in_subclause_block = False
            parent_ends_with_colon = False
            continue

        number = int(num_match.group(1))
        text = num_match.group(2)

        # Chapter?
        is_chapter = False
        if _is_uppercase_text(stripped):
            is_chapter = True
        elif (number == next_chapter
              and len(line) - len(line.lstrip()) == 0
              and len(text) < 80
              and not _text_ends_with_colon(text)):
            is_chapter = True

        if is_chapter:
            chapter_num += 1
            next_chapter = chapter_num + 1
            clause_num = 0
            subclause_num = 0
            last_clause_number = 0
            in_subclause_block = False
            parent_ends_with_colon = False
            clauses.append(Clause(
                ref=f"{chapter_num}",
                level=0,
                text=text,
                line_num=i + 1,
            ))
            continue

        # Sub-clause detection
        next_clause = last_clause_number + 1

        if parent_ends_with_colon and number != next_clause:
            in_subclause_block = True
            subclause_num += 1
            parent_ends_with_colon = _text_ends_with_colon(text)
            clauses.append(Clause(
                ref=f"{chapter_num}.{clause_num}.{subclause_num}",
                level=2,
                text=text,
                line_num=i + 1,
            ))
            continue

        if in_subclause_block:
            if number == next_clause:
                in_subclause_block = False
                clause_num += 1
                subclause_num = 0
                last_clause_number = number
                parent_ends_with_colon = _text_ends_with_colon(text)
                clauses.append(Clause(
                    ref=f"{chapter_num}.{clause_num}",
                    level=1,
                    text=text,
                    line_num=i + 1,
                ))
                continue
            subclause_num += 1
            parent_ends_with_colon = _text_ends_with_colon(text)
            clauses.append(Clause(
                ref=f"{chapter_num}.{clause_num}.{subclause_num}",
                level=2,
                text=text,
                line_num=i + 1,
            ))
            continue

        # Regular clause
        clause_num += 1
        subclause_num = 0
        last_clause_number = number
        parent_ends_with_colon = _text_ends_with_colon(text)
        clauses.append(Clause(
            ref=f"{chapter_num}.{clause_num}",
            level=1,
            text=text,
            line_num=i + 1,
        ))

    return clauses


def find_clause(clauses: list[Clause], query: str) -> list[Clause]:
    """Find clauses by ref prefix or keyword search.

    Examples:
        find_clause(clauses, "4.2")      → clause 4.2 and its sub-clauses
        find_clause(clauses, "оплата")   → clauses containing "оплата"
        find_clause(clauses, "7")        → chapter 7 and all its clauses
    """
    results = []

    # Try as reference first (e.g. "4.2", "4.2.1", "7")
    ref_match = re.match(r'^[\d.]+$', query.strip().rstrip('.'))
    if ref_match:
        ref = query.strip().rstrip('.')
        for c in clauses:
            if c.ref == ref or c.ref.startswith(ref + '.'):
                results.append(c)
        if results:
            return results

    # Keyword search (case-insensitive)
    q = query.lower()
    for c in clauses:
        if q in c.clean_text().lower():
            results.append(c)

    return results


def contract_summary(clauses: list[Clause], md_text: str) -> str:
    """Generate a compact structural summary of a contract.

    Extracts: parties, subject, price, deadlines, chapters outline.
    """
    clean = _text_after_accept(md_text)
    lines = []

    # Title
    first_line = ''
    for l in md_text.split('\n'):
        if l.strip():
            first_line = l.strip()
            break
    lines.append(f"**{first_line}**\n")

    # Parties — look for "именуемое в дальнейшем"
    for m in re.finditer(r'«([^»]+)»[^«]*именуем\w+ в дальнейшем\s*[«"]([^»"]+)[»"]', clean):
        lines.append(f"- {m.group(2)}: {m.group(1)}")

    # Price — look for "Цена Договора" or stoimost
    price_match = re.search(
        r'в размере\s+([\d\s]+)\s*\(([^)]+)\)\s*рублей', clean
    )
    if price_match:
        lines.append(f"- Цена: {price_match.group(1).strip()} ({price_match.group(2)}) руб.")

    # Deadline — look for calendar/working days
    deadline_match = re.search(
        r'Срок выполнения работ?:\s*(\d+\s*(?:календарных|рабочих)\s*дней[^.]*)', clean
    )
    if deadline_match:
        lines.append(f"- Срок: {deadline_match.group(1)}")

    # Chapters outline
    lines.append("\n**Структура:**")
    for c in clauses:
        if c.level == 0:
            lines.append(f"  {c.ref}. {c.clean_text()}")
        elif c.level == 1:
            # Count sub-clauses
            subs = [x for x in clauses if x.ref.startswith(c.ref + '.')]
            sub_info = f" ({len(subs)} подп.)" if subs else ""
            preview = c.clean_text()[:60]
            lines.append(f"    {c.ref}. {preview}...{sub_info}" if len(c.clean_text()) > 60
                         else f"    {c.ref}. {preview}{sub_info}")

    return '\n'.join(lines)


def validate_references(clauses: list[Clause], md_text: str) -> list[dict]:
    """Check internal references (п. X.Y) for correctness.

    Returns list of {ref, line, context, issue}.
    """
    clean = _text_after_accept(md_text)
    valid_refs = {c.ref for c in clauses}
    # Also add with trailing dot
    valid_refs_dot = {r + '.' for r in valid_refs}

    issues = []
    # Find all "п. X.Y." references
    for m in re.finditer(r'п\.\s*([\d]+\.[\d]+\.?)', clean):
        ref = m.group(1).rstrip('.')
        if ref not in valid_refs:
            # Get context
            start = max(0, m.start() - 30)
            end = min(len(clean), m.end() + 30)
            context = clean[start:end].replace('\n', ' ')
            issues.append({
                'ref': f'п. {ref}',
                'context': f'...{context}...',
                'issue': f'Reference п. {ref} not found in contract structure',
            })

    return issues
