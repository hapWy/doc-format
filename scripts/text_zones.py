"""Conservative Markdown exclusion zones shared by the text tools. Stdlib only."""
import re


def protected_zones(text):
    protected = bytearray(len(text))

    def mark(start, end):
        protected[start:end] = b'\x01' * (end - start)

    offset, fence, quoted, html = 0, None, False, None
    lines = text.splitlines(keepends=True)
    frontmatter = bool(lines and lines[0].strip() == '---')
    for number, line in enumerate(lines, 1):
        stripped = line.rstrip('\r\n')
        match = re.match(r'^ {0,3}(`{3,}|~{3,})(.*)$', stripped)
        blank = not stripped.strip()
        if frontmatter:
            mark(offset, offset + len(line))
            if number > 1 and stripped.strip() in ('---', '...'):
                frontmatter = False
        elif fence:
            mark(offset, offset + len(line))
            if match and match[1][0] == fence[0] and len(match[1]) >= fence[1] and not match[2].strip():
                fence = None
        elif html:
            mark(offset, offset + len(line))
            if re.search(r'</' + html + r'\s*>', line, re.I):
                html = None
        elif match:
            mark(offset, offset + len(line))
            fence = (match[1][0], len(match[1]), number)
            quoted = False
        elif re.match(r'^ {0,3}>', line):
            mark(offset, offset + len(line))
            quoted = True
        elif quoted and not blank and not re.match(r'^ {0,3}(?:#{1,6}\s|[-+*]\s|\d+[.)]\s)', line):
            # A paragraph in a blockquote may continue without another > marker.
            mark(offset, offset + len(line))
        elif re.match(r'^(?: {4}| {0,3}\t)', line):
            mark(offset, offset + len(line))
            quoted = False
        else:
            quoted = False
            opening = re.match(r'^ {0,3}<(pre|script|style)\b', line, re.I)
            if opening:
                mark(offset, offset + len(line))
                tag = opening[1].lower()
                html = None if re.search(r'</' + tag + r'\s*>', line, re.I) else tag
        offset += len(line)

    block = protected[:]
    # Parse exact backtick runs only outside block constructs. A literal ``` in
    # fenced source code must never become an inline delimiter in surrounding prose.
    i = 0
    while i < len(text):
        if protected[i] or text[i] != '`':
            i += 1
            continue
        end = i + 1
        while end < len(text) and text[end] == '`' and not block[end]:
            end += 1
        width, j, closing = end - i, end, None
        while j < len(text) and not block[j]:
            if text[j] != '`':
                j += 1
                continue
            run = j + 1
            while run < len(text) and text[run] == '`' and not block[run]:
                run += 1
            if run - j == width:
                closing = run
                break
            j = run
        if closing is not None:
            mark(i, closing)
            i = closing
        else:
            i = end
    code = protected[:]

    def escaped(index):
        previous = index - 1
        while previous >= 0 and text[previous] == '\\':
            previous -= 1
        return (index - previous - 1) % 2 == 1

    pairs = {'«': '»', '“': '”', '„': '“', '‘': '’', '"': '"'}
    i = 0
    while i < len(text):
        close = pairs.get(text[i])
        if protected[i] or not close or escaped(i):
            i += 1
            continue
        j = i + 1
        while j < len(text) and not block[j]:
            if not code[j] and text[j] == close and not escaped(j):
                mark(i, j + 1)
                i = j
                break
            j += 1
        i += 1

    # Preserve destinations, including relative paths and balanced parentheses.
    i = 0
    while i + 1 < len(text):
        if protected[i] or text[i:i + 2] != '](':
            i += 1
            continue
        depth, j = 1, i + 2
        while j < len(text) and not block[j] and text[j] not in '\r\n':
            if not escaped(j):
                if text[j] == '(':
                    depth += 1
                elif text[j] == ')':
                    depth -= 1
                    if depth == 0:
                        mark(i + 1, j + 1)
                        break
            j += 1
        i += 1
    for pattern in (r'(?:https?://|mailto:)\S+', r'<[^>\n]+>',
                    r'^ {0,3}\[[^\]\n]+\]:[^\r\n]*',
                    r'(?<=\])\[[^\]\n]*\]', r'<!--[\s\S]*?-->'):
        for match in re.finditer(pattern, text, re.M):
            if not code[match.start()]:
                mark(match.start(), match.end())
    return protected, fence
