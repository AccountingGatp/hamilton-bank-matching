"""Regression coverage for report periods and account names outside the pilot."""
from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

import fitz
from openpyxl import Workbook

from api.index import output_summary
from matching_statement import run


def write_pdf(path, lines):
    with fitz.open() as document:
        page = document.new_page()
        for y, line in enumerate(lines, 1):
            for x, text in line:
                page.insert_text((x, 24 + y * 16), text, fontsize=9)
        document.save(path)


def write_xlsx(path, rows):
    workbook = Workbook()
    for row in rows:
        workbook.active.append(row)
    workbook.save(path)


class StatementPeriodTest(unittest.TestCase):
    def test_cross_year_statement_and_wider_register(self):
        with TemporaryDirectory() as directory:
            folder = Path(directory)
            statement, prior, register, aging, output = (
                folder / name for name in ('bank.pdf', 'prior.pdf', 'reg.xlsx', 'aging.xlsx', 'report.xlsx')
            )
            write_pdf(statement, [
                [(28, 'December 18, 2024 through January 8, 2025')],
                [(28, 'Account Number:'), (150, '43210')],
                [(28, 'Beginning Balance'), (490, '$100.00')],
                [(28, 'Deposits and Additions'), (465, '2'), (490, '$15.00')],
                [(28, 'Ending Balance'), (465, '2'), (490, '$115.00')],
                [(28, '*start*deposits and additions')],
                [(28, '12/19'), (79, 'First receipt'), (490, '$5.00')],
                [(28, '01/06'), (79, 'Second receipt'), (490, '$10.00')],
                [(28, 'Total Deposits and Additions'), (490, '$15.00')],
                [(28, '*end*deposits and additions')],
            ])
            write_pdf(prior, [
                [(28, 'Savings XXXX3210')],
                [(28, 'Statement Date'), (200, '12/17/2024')],
                [(28, 'Statement Ending Balance'), (490, '100.00')],
                [(28, 'Register Balance (as of 12/17/2024)'), (490, '100.00')],
            ])
            write_xlsx(register, [
                ['Example / Property'],
                ['Example Management'],
                ['Bank Account Register'],
                ['12/01/2024 - 01/31/2025'],
                [],
                ['Printed 02/01/2025'],
                ['Bank', 'Account Number', 'Account Name'],
                ['Chase', 'XXXXX3210', 'Reserve Savings'],
                ['Date', 'Type', 'Reference', 'Description/Payee', 'Debit (+)', 'Credit (-)', 'Balance'],
                [date(2024, 12, 12), 'DEP', '1', 'Before prior cutoff', 20, None, 100],
                [date(2024, 12, 19), 'DEP', '2', 'First receipt', 5, None, 105],
                [date(2025, 1, 6), 'DEP', '3', 'Second receipt', 10, None, 115],
                [date(2025, 1, 10), 'DEP', '4', 'After statement cutoff', 7, None, 122],
                [None, None, None, None, 'Beginning Balance', None, 80],
                [None, None, None, None, 'Net Change', None, 42],
                [None, None, None, None, 'Ending Balance', None, 122],
            ])
            write_xlsx(aging, [
                ['Example / Property'],
                ['Example Management'],
                ['AP Aging Detail'],
                ['Accounting Date 01/31/2025 - Historical'],
                [],
                [],
                ['Invoice Number', 'Invoice Date', 'Balance'],
            ])
            run(statement, register, prior, aging, output)
            result = output_summary(output)
            self.assertEqual(result['period'], '2024-12-18 to 2025-01-08')
            self.assertEqual(result['filename'], 'Matching_Statement_Example_Property_3210_2024-12-18_to_2025-01-08.xlsx')
            self.assertEqual(result['bank'], 115)
            self.assertEqual(result['adjusted'], 115)
            self.assertEqual(result['difference'], 0)
            self.assertEqual(result['exceptions'], 0)
