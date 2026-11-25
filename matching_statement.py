"""Create a Matching Statement workbook from Chase and ResMan exports."""
import argparse
from calendar import monthrange
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal
from itertools import combinations
from pathlib import Path
import re
import sys

from openpyxl import load_workbook
import xlsxwriter

from pdf_layout import pdf_text


MONEY = r'\$?-?[\d,]+\.\d{2}'
DAY = r'\d{2}/\d{2}'
FULL_DAY = r'\d{1,2}/\d{1,2}/\d{4}'
LONG_DAY = r'[A-Za-z]+ \d{1,2}, \d{4}'
MONTH = r'[A-Za-z]+ \d{4}'
PERIOD_DATE = rf'(?:{LONG_DAY}|{FULL_DAY}|\d{{4}}-\d{{2}}-\d{{2}}|{MONTH})'
SECTIONS = {
    'deposits and additions': ('Deposits and Additions', 1),
    'atm debit withdrawal': ('ATM & Debit Card Withdrawals', -1),
    'electronic withdrawal': ('Electronic Withdrawals', -1),
    'checks paid': ('Checks Paid', -1),
    'other withdrawals': ('Other Withdrawals', -1),
    'fees': ('Fees', -1),
}


@dataclass(frozen=True)
class Transaction:
    key: str
    day: date
    cents: int
    description: str
    reference: str = ''
    kind: str = ''
    alt_day: date | None = None

    def label(self):
        return f'{self.key} {self.day:%m/%d/%Y} {self.reference} {self.description}'.strip()


def cents(value):
    return int(round(Decimal(str(value or 0).replace('$', '').replace(',', '')) * 100))


def full_day(value):
    return datetime.strptime(value, '%m/%d/%Y').date()


def short_day(value, start, end, *, optional=False):
    candidates = []
    for year in range(start.year, end.year + 1):
        try:
            day = full_day(f'{value}/{year}')
        except ValueError:
            continue
        if start <= day <= end:
            candidates.append(day)
    if len(candidates) == 1:
        return candidates[0]
    if optional:
        return None
    raise ValueError(f'Chase posting date {value} is missing or ambiguous in the statement period.')


def register_period(label):
    label = re.sub(r'^\s*(?:date range|period|from)\s*:?\s*', '', label.strip(), flags=re.I)
    period = re.fullmatch(
        rf'({PERIOD_DATE})(?:\s*(?:through|to|[-–—])\s*({PERIOD_DATE}))?',
        label, re.I,
    )
    if not period:
        return None

    def boundary(value, is_end):
        for fmt in ('%m/%d/%Y', '%Y-%m-%d', '%B %d, %Y', '%b %d, %Y'):
            try:
                return datetime.strptime(value, fmt).date()
            except ValueError:
                continue
        for fmt in ('%B %Y', '%b %Y'):
            try:
                day = datetime.strptime(value, fmt).date()
            except ValueError:
                continue
            return day.replace(day=monthrange(day.year, day.month)[1]) if is_end else day
        return None

    start = boundary(period[1], False)
    end = boundary(period[2] or period[1], True)
    return (start, end) if start and end and start <= end else None


def required(text, pattern, label):
    match = re.search(pattern, text, re.M | re.I)
    if not match:
        raise ValueError(f'Missing {label}. Check the input format and account.')
    return match


def read_chase(path):
    text = pdf_text(path)
    dates = required(text, rf'({LONG_DAY})\s+through\s+({LONG_DAY})', 'Chase statement dates')
    start, end = (datetime.strptime(value, '%B %d, %Y').date() for value in dates.groups())
    if start > end:
        raise ValueError('Chase statement start date follows its end date.')
    account = required(text, r'Account Number:\s*(\d+)', 'Chase account number')[1][-4:]
    opening = cents(required(text, rf'Beginning Balance\s+({MONEY})', 'bank opening balance')[1])
    close = required(text, rf'Ending Balance\s+(\d+)\s+({MONEY})', 'bank ending balance and count')
    closing = cents(close[2])
    activity = []
    active = pending = None
    page = 1

    def flush():
        nonlocal pending
        if pending is None:
            return
        day, amount, parts, key, section, number = pending
        if amount is None:
            raise ValueError(f'Chase transaction {key} is missing its amount.')
        description = ' '.join(parts)
        reference = re.search(r'(?:Online (?:ACH )?Payment |Transaction#:\s*|Trn:\s*|Trace#:\s*)([A-Za-z0-9]+)', description)
        alt = re.search(rf'(?<!\d)({DAY})(?![\d/])', description)
        posting_day = short_day(day, start, end)
        activity.append(Transaction(
            key, posting_day, amount * SECTIONS[section][1], description,
            number or (reference[1] if reference else ''), SECTIONS[section][0],
            short_day(alt[1], posting_day - timedelta(days=31), posting_day, optional=True) if alt else None,
        ))
        pending = None

    for line_no, raw in enumerate(text.replace('\f', '\nPAGE_BREAK\n').splitlines(), 1):
        line = raw.strip()
        if line == 'PAGE_BREAK':
            flush()
            active = None
            page += 1
            continue
        if line.startswith('*start*'):
            flush()
            active = line.removeprefix('*start*').strip()
            continue
        if line.startswith('*end*'):
            # Chase sometimes prints this marker before the final amount or continuation.
            continue
        if active not in SECTIONS or not line or re.fullmatch(r'\d{18,}', line):
            continue
        if line.startswith(('Total ', 'Page ', '*')):
            flush()
            active = None
            continue
        row = re.match(rf'^(?P<date>{DAY})(?:\s{{2,}}(?P<body>.*))?$', line)
        check = re.match(rf'^(\d+)\*?\s+.*?({DAY})\s+({MONEY})$', line) if active == 'checks paid' else None
        amount = re.search(rf'\s{{2,}}({MONEY})$', line)
        if check:
            flush()
            pending = [check[2], cents(check[3]), [f'Check No. {check[1]}'], f'PDF p{page} L{line_no}', active, check[1]]
        elif row:
            flush()
            body = row['body'] or ''
            if amount:
                body = line[row.start('body'):amount.start()].strip()
            pending = [row['date'], cents(amount[1]) if amount else None, [body] if body else [],
                       f'PDF p{page} L{line_no}', active, '']
        elif pending is not None:
            if pending[1] is None and amount:
                pending[1] = cents(amount[1])
                line = line[:amount.start()].strip()
            if line:
                pending[2].append(line)
    flush()

    for section, (label, sign) in SECTIONS.items():
        summary = re.search(rf'^\s*{re.escape(label)}\s+(\d+)\s+({MONEY})\s*$', text, re.M | re.I)
        rows = [item for item in activity if item.kind == label]
        if summary is None and not rows:
            continue
        if summary is None or (len(rows), sum(x.cents for x in rows)) != (int(summary[1]), sign * abs(cents(summary[2]))):
            raise ValueError(f'{label} extracted count or amount differs from the Chase summary.')
    if len(activity) != int(close[1]) or opening + sum(x.cents for x in activity) != closing:
        raise ValueError('Extracted Chase rows do not tie to the statement count and balances.')
    if any(not start <= item.day <= end for item in activity):
        raise ValueError('A Chase transaction is outside the statement dates.')
    return account, start, end, opening, closing, activity


def excel_rows(path):
    workbook = load_workbook(path, data_only=True, read_only=True)
    try:
        return list(workbook.active.values)
    finally:
        workbook.close()


def read_register(path):
    rows = excel_rows(path)
    header_index = next(i for i, row in enumerate(rows) if row and row[0] == 'Date')
    header = rows[header_index]
    account_index = next(
        i for i, row in enumerate(rows[:header_index]) if 'Account Number' in row and 'Account Name' in row
    )
    account_header = rows[account_index]
    account_values = rows[account_index + 1]
    masked = str(account_values[account_header.index('Account Number')] or '')
    account = required(masked, r'(\d{4})\s*$', 'ResMan account number')[1]
    account_name = str(account_values[account_header.index('Account Name')] or f'Account {account}').strip()
    period = next(
        (parsed for row in rows[3:account_index] if row and isinstance(row[0], str)
         if (parsed := register_period(row[0])) is not None),
        None,
    )
    if period is None:
        raise ValueError('The ResMan register must show a reporting month or a start and end date in its header.')
    columns = {value: header.index(value) for value in (
        'Date', 'Type', 'Reference', 'Description/Payee', 'Debit (+)', 'Credit (-)', 'Balance',
    )}
    activity = []
    balances = {}
    for row_number, values in enumerate(rows, 1):
        def cell(column):
            position = columns[column]
            return values[position] if len(values) > position else None

        balance_label = cell('Debit (+)')
        if balance_label in ('Beginning Balance', 'Ending Balance', 'Net Change'):
            balances[balance_label] = cents(cell('Balance'))
        day = cell('Date')
        if not isinstance(day, datetime):
            continue
        amount = cents(cell('Debit (+)')) - cents(cell('Credit (-)'))
        if not amount:
            continue
        activity.append(Transaction(
            f'Register row {row_number}', day.date(), amount, str(cell('Description/Payee') or ''),
            str(cell('Reference') or ''), str(cell('Type') or ''),
        ))
    net = sum(item.cents for item in activity)
    if balances['Net Change'] != net or balances['Beginning Balance'] + net != balances['Ending Balance']:
        raise ValueError('ResMan rows do not tie to the reported net change and balances.')
    if any(not period[0] <= item.day <= period[1] for item in activity):
        raise ValueError('Register entries fall outside the reporting dates shown in the ResMan export.')
    return str(rows[0][0]), account, account_name, period, balances['Beginning Balance'], activity


def read_prior(path, start):
    text = pdf_text(path)
    account = required(text, r'X{3,}(\d{4})\b', 'prior ResMan account number')[1]
    closing_day = full_day(required(text, rf'Statement Date\s+({FULL_DAY})', 'prior statement date')[1])
    if closing_day >= start:
        raise ValueError('Prior reconciliation must end before the current Chase statement starts.')
    bank = cents(required(text, rf'Statement Ending Balance\s+({MONEY})', 'prior bank balance')[1])
    books = cents(required(text, rf'Register Balance \(as of {FULL_DAY}\)\s+({MONEY})', 'prior register balance')[1])
    outstanding, cleared = [], []
    in_outstanding = False
    credit_position = None
    for number, raw in enumerate(text.splitlines(), 1):
        line = raw.strip()
        if line.startswith('Outstanding Transactions'):
            in_outstanding = True
        if line.startswith('Total Outstanding Transactions'):
            break
        if 'Debits (+)' in raw and 'Credits (-)' in raw:
            credit_position = raw.index('Credits (-)')
        row = re.match(rf'^\s*({FULL_DAY})(\*)?\s+(\S+)\s+(.*?)\s{{2,}}({MONEY})\s*$', raw)
        if not row or not (in_outstanding or row[2]):
            continue
        fields = re.split(r'\s{2,}', row[4], maxsplit=1)
        reference, label = fields if len(fields) == 2 else ('', fields[0])
        if row[2] and len(fields) == 1 and re.fullmatch(r'[0-9./-]+', label):
            reference, label = label, ''
        # Some descriptions resume on the adjacent PDF line. Future rows are
        # identified by date, document type, signed amount and reference instead.
        sign = -1 if row.end(5) >= credit_position else 1
        target = outstanding if in_outstanding else cleared
        target.append(Transaction(f'Prior PDF L{number}', full_day(row[1]), cents(row[5]) * sign,
                                  label, reference, row[3]))
    if sum(x.cents for x in outstanding) - sum(x.cents for x in cleared) != books - bank:
        raise ValueError('Prior outstanding and future-cleared rows do not explain the opening difference.')
    return account, bank, books, closing_day, outstanding, cleared


def normalize(value):
    return re.sub(r'[^a-z0-9]', '', value.lower())


def read_aging(path, property_name):
    rows = excel_rows(path)
    if normalize(str(rows[0][0])) != normalize(property_name):
        raise ValueError('AP Aging is for a different property.')
    header = next(row for row in rows if row and row[0] == 'Invoice Number')
    date_label = required(str(rows[3][0]), rf'Accounting Date\s+({FULL_DAY})', 'AP aging date')[1]
    aging_day = full_day(date_label)
    vendor = ''
    invoices = []
    for number, row in enumerate(rows, 1):
        if not row or not row[0] or row == header:
            continue
        if number < 8:
            continue
        if not any(row[1:]):
            vendor = str(row[0])
        elif isinstance(row[header.index('Invoice Date')], datetime):
            invoices.append((vendor, str(row[0]).strip(), cents(row[header.index('Balance')]), number))
    return aging_day, invoices


def exact_reference(bank, book):
    ref = book.reference
    return bool(ref and ref == bank.reference)


def inferred_day_gap(bank, book):
    return min(abs((day - book.day).days) for day in (bank.day, bank.alt_day or bank.day))


def same_words(first, second):
    return any(word in second.lower() for word in re.findall(r'[a-z]{5,}', first.lower())
               if word not in {'credit', 'debit', 'payment', 'presidio', 'flats', 'company', 'online'})


def reconcile(bank, books, outstanding, prior_cleared, window):
    matched, review = [], []
    banks = {item.key: item for item in bank}
    registers = {item.key: item for item in [*books, *outstanding]}

    def pair(bank_item, book_items, rule, evidence=''):
        del banks[bank_item.key]
        for item in book_items:
            del registers[item.key]
        matched.append((bank_item, book_items, rule, evidence))

    current_rows = set(books)
    for old in prior_cleared:
        candidates = [item for item in registers.values()
                      if item in current_rows and (item.day, item.kind, item.cents, item.reference)
                      == (old.day, old.kind, old.cents, old.reference)]
        if len(candidates) != 1:
            raise ValueError(f'Cannot verify the prior-cleared register item {old.label()}.')
        item = candidates[0]
        del registers[item.key]
        matched.append((None, [item], 'Cleared on prior reconciliation', old.label()))

    def unique(rule, eligible):
        pairs = [(bank_item, book_item) for bank_item in banks.values()
                 for book_item in registers.values() if bank_item.cents == book_item.cents
                 and eligible(bank_item, book_item)]
        bank_counts = Counter(bank_item.key for bank_item, _ in pairs)
        book_counts = Counter(book_item.key for _, book_item in pairs)
        for bank_item, book_item in pairs:
            if bank_counts[bank_item.key] == book_counts[book_item.key] == 1:
                pair(bank_item, [book_item], rule)

    unique('Exact amount and payment reference', lambda bank_item, book_item: exact_reference(bank_item, book_item)
           and (book_item.day <= bank_item.day + timedelta(days=window)))
    unique('Prior outstanding item cleared', lambda bank_item, book_item:
           book_item in outstanding and inferred_day_gap(bank_item, book_item) <= window
           and same_words(book_item.description, bank_item.description))
    unique('Exact amount and bank posting date', lambda bank_item, book_item:
           book_item not in outstanding and bank_item.day == book_item.day)
    unique('Exact amount and printed transaction date', lambda bank_item, book_item:
           book_item not in outstanding and bank_item.alt_day == book_item.day)
    unique(f'Exact amount within {window} calendar days', lambda bank_item, book_item:
           book_item not in outstanding and abs((bank_item.day - book_item.day).days) <= window)

    proposed = []
    for bank_item in banks.values():
        for principal in registers.values():
            delta = bank_item.cents - principal.cents
            if not 0 < abs(delta) <= 100 or principal.cents * bank_item.cents <= 0:
                continue
            if inferred_day_gap(bank_item, principal) > window or not same_words(principal.description, bank_item.description):
                continue
            adjustments = [item for item in registers.values() if item.kind == 'JE' and item.cents == delta
                           and item.day == principal.day and 'difference' in item.description.lower()
                           and same_words(principal.description, item.description)]
            if len(adjustments) == 1:
                proposed.append((bank_item, [principal, adjustments[0]]))
    accept_groups(proposed, pair, 'Exact sum with explicit difference JE',
                  'Verify the principal and adjustment against the payment detail.')

    proposed = []
    for bank_item in banks.values():
        if bank_item.cents <= 0 or 'paylease' not in bank_item.description.lower():
            continue
        pool = [item for item in registers.values()
                if item.kind == 'DEP' and 0 < item.cents < bank_item.cents
                and abs((bank_item.day - item.day).days) <= window
                and (item.description == item.reference or 'paylease' in item.description.lower())
                and not any(other.cents == item.cents for other in banks.values())]
        groups = [list(group) for group in combinations(pool, 2)
                  if sum(item.cents for item in group) == bank_item.cents]
        if len(groups) == 1:
            proposed.append((bank_item, groups[0]))
    accept_groups(proposed, pair, 'Provisional PayLease deposit group',
                  'Unique two-row sum in the date window. Confirm the PayLease settlement detail.')

    for bank_item in banks.values():
        for book_item in registers.values():
            if bank_item.cents == book_item.cents:
                gap = abs((bank_item.day - book_item.day).days)
                review.append((bank_item, book_item, f'Same signed amount, {gap} day(s) apart. Verify source before pairing.'))
    for bank_item, group, rule, evidence in matched:
        if rule.startswith('Provisional'):
            review.extend((bank_item, item, evidence) for item in group)

    groups = [
        [item for item in registers.values() if item.cents < 0],
        [item for item in banks.values() if item.cents < 0],
        [item for item in banks.values() if item.cents > 0],
        [item for item in registers.values() if item.cents > 0],
    ]
    for index, source in ((0, [*books, *outstanding]), (1, bank)):
        processed = [item for bank_item, book_items, _, _ in matched
                     for item in (book_items if index == 0 else [bank_item]) if item is not None]
        missing = groups[index] + groups[3 - index]
        if Counter(source) != Counter(processed + missing):
            raise ValueError('Each source transaction must be assigned to exactly one result.')
    return groups, matched, review


def accept_groups(proposed, pair, rule, evidence):
    bank_counts = Counter(bank_item.key for bank_item, _ in proposed)
    book_counts = Counter(item.key for _, group in proposed for item in group)
    for bank_item, group in proposed:
        if bank_counts[bank_item.key] == 1 and all(book_counts[item.key] == 1 for item in group):
            pair(bank_item, group, rule, evidence)


def write_report(path, property_name, account, account_name, start, end, bank_open, bank_close, book_open, book_close,
                 groups, matched, review, aging, source_files, counts, window, prior):
    titles = [
        'Add: Payments in books pending debit in bank',
        'Less: Payments in bank missing from books',
        'Add: Deposits in bank missing from books',
        'Less: Deposits in books pending credit in bank',
    ]
    subtotal = [sum(abs(item.cents) for item in group) for group in groups]
    adjusted = book_close + subtotal[0] - subtotal[1] + subtotal[2] - subtotal[3]
    if adjusted != bank_close:
        raise ValueError(f'Adjusted balance differs from Chase by {(adjusted - bank_close) / 100:,.2f}.')

    counterpart = defaultdict(list)
    for bank_item, book_item, note in review:
        counterpart[bank_item.key].append(f'{book_item.key}: {note}')
        counterpart[book_item.key].append(f'{bank_item.key}: {note}')

    potential_dupes = defaultdict(list)
    for group in groups:
        signatures = defaultdict(list)
        for item in group:
            merchant = normalize(re.sub(DAY, '', item.description))
            signatures[item.cents, merchant].append(item)
        for same in signatures.values():
            if len(same) > 1:
                for item in same:
                    potential_dupes[item.key].extend(other.key for other in same if other != item)

    aging_date, invoices = aging
    if aging_date < end:
        invoices = []
    ap_suggestions = {}
    for item in groups[0] + groups[1]:
        ap_suggestions[item.key] = [
            (vendor, inv, row) for vendor, inv, balance, row in invoices
            if balance == abs(item.cents) and same_words(vendor, item.description)
        ]

    def comment(item, index):
        if item.key.startswith('Prior'):
            note = 'Still outstanding from the prior reconciliation. Verify clearance or supporting documentation.'
        elif index in (1, 2):
            note = 'No match in the supplied register. Verify supporting detail.'
        else:
            note = 'No match in the supplied statement. Check posting or next statement.'
        if counterpart[item.key]:
            note += ' ' + ' '.join(counterpart[item.key])
        if potential_dupes[item.key]:
            note += ' Repeated description and amount with ' + ', '.join(potential_dupes[item.key]) + '. Verify both items.'
        if ap_suggestions.get(item.key):
            detail = ', '.join(f'{vendor} invoice {inv} (row {row})' for vendor, inv, row in ap_suggestions[item.key])
            note += f' Possible open AP support: {detail}. Verify before posting.'
        return f'{note} [{item.key}]'

    with xlsxwriter.Workbook(str(path), {'strings_to_formulas': False, 'strings_to_urls': False}) as workbook:
        base = {'font_name': 'Arial', 'font_size': 11, 'valign': 'vcenter'}
        money = {'num_format': '#,##0.00;(#,##0.00)'}
        title = workbook.add_format({**base, 'font_size': 14, 'bold': True, 'align': 'center'})
        subtitle = workbook.add_format({**base, 'bold': True, 'align': 'center'})
        band = workbook.add_format({**base, 'bold': True})
        plain = workbook.add_format({**base, 'text_wrap': True})
        ledger_money = workbook.add_format({**base, **money})
        grid = workbook.add_format({**base, 'border': 1, 'text_wrap': True})
        grid_money = workbook.add_format({**base, **money, 'border': 1})
        borders = {}
        for style in ('header', 'total', 'subtotal', 'difference'):
            for col in range(7):
                borders[style, col] = workbook.add_format({
                    **base, 'bold': True, 'italic': style == 'subtotal',
                    'top': 1 if style == 'difference' else 2, 'bottom': 2,
                    'left': 2 if col == 0 else int(style == 'header'),
                    'right': 2 if col == 6 else int(style == 'header'),
                    **(money if col == 4 else {}),
                })

        tab_prefix = re.sub(r"[\[\]:*?/\\']", ' ', property_name).strip()[:26].rstrip()
        sheet = workbook.add_worksheet(f'{tab_prefix or "Account"} {account}')
        sheet.hide_gridlines(2)
        sheet.freeze_panes(3, 0)
        for column, width in enumerate((65, 13, 24, 16, 18, 14, 55)):
            sheet.set_column(column, column, width)
        sheet.merge_range('A1:G1', f'{property_name} — Bank Reconciliation', title)
        sheet.set_row(0, 22)
        sheet.merge_range('A2:G2', f'{account_name} | Chase …{account} | {start:%m/%d/%Y} to {end:%m/%d/%Y}', subtitle)
        sheet.set_row(1, 18)
        row = 3

        def section(label):
            nonlocal row
            sheet.merge_range(row, 0, row, 6, label, band)
            sheet.set_row(row, 19)
            row += 1

        def summary(label, value, formula=None, style=None, closing_date=None):
            nonlocal row
            for column in range(7):
                fmt = borders[style, column] if style else (grid_money if column == 4 else grid)
                if column == 4 and formula:
                    sheet.write_formula(row, column, formula, fmt, value / 100)
                else:
                    values = {0: label, 4: value / 100, 5: closing_date}
                    sheet.write(row, column, values.get(column), fmt)
            row += 1
            return f'E{row}'

        section('OPENING BALANCES')
        opening_date = prior[0].strftime('%m/%d/%Y')
        bank_open_cell = summary('Bank Opening Balance', bank_open, closing_date=opening_date)
        book_open_cell = summary('Books Opening Balance (ResMan)', book_open, closing_date=opening_date)
        opening_difference = summary('Opening Difference (Bank − Books)', bank_open - book_open,
                                     f'={bank_open_cell}-{book_open_cell}', 'difference')
        for col, label in enumerate(('Particulars', 'Date', 'Ch. No./Reference', '$', 'Net Amount',
                                     'Closing Date', 'GATP Comment')):
            sheet.write(row, col, label, borders['header', col])
        row += 1

        subtotal_cells = []
        for index, group in enumerate(groups):
            section(titles[index])
            first = row + 1
            for item in sorted(group, key=lambda x: (x.day, x.key)):
                values = (item.description, item.day.strftime('%m/%d/%Y'), item.reference,
                          abs(item.cents) / 100, None, None, comment(item, index))
                for column, value in enumerate(values):
                    sheet.write(row, column, value, grid_money if column == 3 else grid)
                sheet.set_row(row, 40)
                row += 1
            if not group:
                for column in range(7):
                    sheet.write(row, column, '— None —' if column == 0 else None, grid)
                row += 1
            subtotal_cells.append(summary(f'Subtotal — {len(group)} item(s)', subtotal[index],
                                          f'=SUM(D{first}:D{row})', 'subtotal'))
        closing_date = end.strftime('%m/%d/%Y')
        book_close_cell = summary('Balance as per Books (ResMan)', book_close, style='total', closing_date=closing_date)
        s1, s2, s4, s3 = subtotal_cells
        adjusted_cell = summary('Adjusted Balance', adjusted, f'={book_close_cell}+{s1}-{s2}+{s4}-{s3}', 'total')
        bank_close_cell = summary('Balance as per Bank Statement', bank_close, style='total', closing_date=closing_date)
        summary('Difference if any (Adjusted − Bank)', adjusted - bank_close,
                f'=ROUND({adjusted_cell}-{bank_close_cell},2)', 'difference')

        section('CLOSING BALANCES')
        summary('Bank Closing Balance', bank_close, closing_date=closing_date)
        summary('Books Closing Balance (ResMan)', book_close, closing_date=closing_date)
        balance_formula = f'=ROUND({bank_close_cell}-{book_close_cell},2)'
        summary('Closing Difference (Bank − Books)', bank_close - book_close, balance_formula, 'total')

        section('CLOSING DIFFERENCE BREAKDOWN')
        summary('Opening Difference (carried forward)', bank_open - book_open, f'={opening_difference}')
        summary('(+) Current Period Bank Net', bank_close - bank_open)
        summary('(−) Current Period Books Net', book_close - book_open)
        summary('Closing Difference (Bank − Books)', bank_close - book_close, balance_formula, 'total')
        sheet.set_landscape()
        sheet.fit_to_pages(1, 0)
        sheet.print_area(0, 0, row - 1, 6)

        def audit_table(name, headers, widths, records, currency=(), tall=False):
            tab = workbook.add_worksheet(name)
            tab.hide_gridlines(2)
            tab.freeze_panes(1, 0)
            for col, (label, width) in enumerate(zip(headers, widths)):
                tab.set_column(col, col, width)
                tab.write(0, col, label, band)
            for row_number, values in enumerate(records, 1):
                if tall:
                    tab.set_row(row_number, 44)
                for col, value in enumerate(values):
                    tab.write(row_number, col, value, ledger_money if col in currency else plain)
            if records:
                tab.autofilter(0, 0, len(records), len(headers) - 1)

        audit_table(
            'Match Detail',
            ['Rule', 'Bank Source', 'Bank Date', 'Bank Reference', 'Bank Amount', 'Bank Particulars',
             'ResMan or Prior Source', 'Books Amount', 'Notes'],
            [36, 19, 13, 22, 16, 80, 72, 16, 66],
            [(rule, b.key if b else '', b.day.isoformat() if b else '', b.reference if b else '',
              b.cents / 100 if b else None, b.description if b else '',
              '\n'.join(item.label() for item in items), sum(item.cents for item in items) / 100, note)
             for b, items, rule, note in sorted(matched, key=lambda m: (m[0].day if m[0] else m[1][0].day, m[2]))],
            (4, 7), True,
        )
        audit_table(
            'Review Candidates',
            ['Bank Date', 'Signed Bank Amount', 'Bank Source and Particulars', 'Books Source and Particulars',
             'Signed Books Amount', 'Review Reason'],
            [14, 21, 72, 70, 22, 70],
            [(b.day.isoformat(), b.cents / 100, b.label(), k.label(), k.cents / 100, note)
             for b, k, note in review],
            (1, 4), True,
        )
        _, prior_outstanding, prior_cleared = prior
        audit_table(
            'Controls and Sources',
            ['Control', 'Result', 'Detail'],
            [43, 31, 101],
            [
                ('Property', property_name, account_name),
                ('Account Last Four', account, 'Validated across Chase and both ResMan account reports.'),
                ('Statement Start', start.isoformat(), ''),
                ('Statement End', end.isoformat(), ''),
                ('Chase PDF', source_files[0].name, f'{start:%m/%d/%Y} through {end:%m/%d/%Y}'),
                ('ResMan register', source_files[1].name, f'Register activity after {end:%m/%d/%Y} is excluded from this statement.'),
                ('Prior reconciliation', source_files[2].name, 'Prior bank and register endings tie to both current opening balances.'),
                ('AP Aging Detail', source_files[3].name, f'As of {aging_date:%m/%d/%Y}. ' + (
                    'Snapshot predates the cutoff and was excluded from invoice suggestions.'
                    if aging_date < end else 'Open invoice evidence only.')),
                ('Bank transaction count', counts[0], 'Extracted count and signed totals checked against the Chase summary.'),
                ('Register export transaction count', counts[1], 'Extracted signed activity checked against the ResMan report totals.'),
                ('Register rows in reconciliation', counts[2], 'After prior cutoff through the Chase ending date.'),
                ('AP invoice rows', len(invoices), 'Candidate support requires equal open balance and overlapping payee text.'),
                ('AP candidate hits', sum(map(len, ap_suggestions.values())), 'Any hits appear in the matching statement comments.'),
                ('Prior outstanding rows', len(prior_outstanding), f'Signed: {sum(x.cents for x in prior_outstanding) / 100:,.2f}'),
                ('Prior-cleared current register rows', len(prior_cleared), f'Signed: {sum(x.cents for x in prior_cleared) / 100:,.2f}. See Match Detail.'),
                ('Date tolerance', f'{window} calendar days', 'Exact signed amount required. Repeated amounts remain open if the dates and references fail to resolve them.'),
                ('PayLease group matches', sum(rule.startswith('Provisional') for _, _, rule, _ in matched),
                 'Included in the working reconciliation. Confirm against processor settlement detail.'),
                ('Reconciliation difference', f'{(adjusted - bank_close)/100:,.2f}', 'Adjusted balance less statement ending balance.'),
            ],
        )
    return subtotal


def report_filename(property_name, account, start, end):
    property_slug = re.sub(r'[^A-Za-z0-9]+', '_', property_name).strip('_')[:80] or 'Property'
    return f'Matching_Statement_{property_slug}_{account}_{start}_to_{end}.xlsx'


def run(statement, register, prior_path, aging_path, output, window=5):
    bank_account, start, end, bank_open, bank_close, bank = read_chase(statement)
    property_name, book_account, account_name, (register_start, register_end), register_open, books = read_register(register)
    prior_account, prior_bank, prior_book, prior_day, outstanding, prior_cleared = read_prior(prior_path, start)
    if not bank_account == book_account == prior_account:
        raise ValueError('Chase, ResMan, and prior reconciliation account numbers must agree.')
    if register_start > prior_day + timedelta(days=1) or register_end < end:
        raise ValueError(f'Export the ResMan register for at least {prior_day + timedelta(days=1):%m/%d/%Y} '
                         f'through {end:%m/%d/%Y} to cover the reconciliation period.')
    book_open = register_open + sum(book.cents for book in books if book.day <= prior_day)
    if (bank_open, book_open) != (prior_bank, prior_book):
        raise ValueError('Current opening balances differ from the prior reconciliation.')
    in_cycle = [book for book in books if prior_day < book.day <= end]
    book_close = book_open + sum(book.cents for book in in_cycle)
    aging = read_aging(aging_path, property_name)
    groups, matched, review = reconcile(bank, in_cycle, outstanding, prior_cleared, window)
    source_files = statement, register, prior_path, aging_path
    if output is None:
        output = Path.cwd() / report_filename(property_name, bank_account, start, end)
    if output.suffix.lower() != '.xlsx' or output.resolve() in {path.resolve() for path in source_files}:
        raise ValueError('Choose a new .xlsx output path distinct from the inputs.')
    if output.exists():
        raise ValueError(f'Output already exists: {output}. Choose a new output path.')
    write_report(output, property_name, bank_account, account_name, start, end, bank_open, bank_close, book_open,
                 book_close, groups, matched, review, aging, source_files, (len(bank), len(books), len(in_cycle)),
                 window, (prior_day, outstanding, prior_cleared))
    return output


def ask_path(label):
    return Path(input(f'{label}: ').strip().strip('"'))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--statement', type=Path, help='Current Chase statement PDF')
    parser.add_argument('--register', type=Path, help='Current ResMan bank register XLSX')
    parser.add_argument('--prior', type=Path, help='Previous finalized ResMan reconciliation PDF')
    parser.add_argument('--aging', type=Path, help='Current AP Aging Detail XLSX')
    parser.add_argument('--output', type=Path, help='New output XLSX path')
    parser.add_argument('--window', type=int, default=5, help='Maximum timing variation in days (default 5)')
    args = parser.parse_args()
    try:
        if not 0 <= args.window <= 10:
            raise ValueError('Use a timing window of 0 to 10 days.')
        if not any((args.statement, args.register, args.prior, args.aging)):
            print('Paste each file path or drag the file into this window, then press Enter.')
            args.statement = ask_path('Current Chase statement PDF')
            args.register = ask_path('Current ResMan register XLSX')
            args.prior = ask_path('Previous ResMan reconciliation PDF')
            args.aging = ask_path('Current AP Aging Detail XLSX')
        if not all((args.statement, args.register, args.prior, args.aging)):
            raise ValueError('Supply --statement, --register, --prior and --aging together.')
        print(run(args.statement, args.register, args.prior, args.aging, args.output, args.window))
    except (OSError, ValueError, KeyError, StopIteration) as error:
        print(f'Reconciliation stopped. {error}', file=sys.stderr)
        return 1
    return 0


if __name__ == '__main__':
    sys.exit(main())
