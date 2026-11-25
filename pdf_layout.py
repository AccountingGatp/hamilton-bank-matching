"""Render PDF text in column order while keeping Chase section markers separate."""
import fitz


def page_layout(page):
    entries = []
    for block in page.get_text('dict', flags=fitz.TEXTFLAGS_DICT & ~fitz.TEXT_PRESERVE_IMAGES)['blocks']:
        for line in block.get('lines', []):
            if line['dir'][0] < .99:
                continue
            for span in line['spans']:
                text = span['text'].strip()
                if text:
                    x, baseline = span['origin']
                    entries.append((baseline, x, text))
    lines = []
    row = []
    last_baseline = None

    def flush():
        if not row:
            return
        rendered = ''
        for x, text in sorted(row):
            column = round(x / 3.2)
            rendered += ' ' * max(1 if rendered else 0, column - len(rendered)) + text
        lines.append(rendered)

    for baseline, x, text in sorted(entries):
        marker = text.startswith(('*start*', '*end*'))
        if marker or (last_baseline is not None and baseline - last_baseline > 2):
            flush()
            row.clear()
        if marker:
            lines.append(text)
            last_baseline = None
        else:
            row.append((x, text))
            last_baseline = baseline
    flush()
    return '\n'.join(lines)


def pdf_text(path):
    try:
        with fitz.open(path) as document:
            if document.needs_pass and not document.authenticate(''):
                raise ValueError('Upload a PDF that opens without a password.')
            return '\n\f\n'.join(page_layout(page) for page in document)
    except fitz.FileDataError:
        raise ValueError('Unable to read the PDF. Export it again and retry.') from None
