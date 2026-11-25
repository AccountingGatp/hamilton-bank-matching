"""Serve the Matching Statement form and process uploaded files."""
from email import policy
from email.parser import BytesParser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler
import json
import logging
from pathlib import Path
import socket
import tempfile
from zipfile import BadZipFile

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from matching_statement import report_filename, run


FIELDS = {
    'statement': ('Chase statement', '.pdf'),
    'register': ('ResMan bank register', '.xlsx'),
    'prior': ('Prior ResMan reconciliation', '.pdf'),
    'aging': ('AP Aging Detail', '.xlsx'),
}
MAX_BYTES = 4_000_000
PAGE = Path(__file__).resolve().parent.parent / 'index.html'


def uploaded_files(request, directory):
    media_type = request.headers.get('Content-Type', '')
    if request.headers.get_content_type() != 'multipart/form-data':
        raise ValueError('Select the four input files using the upload form.')
    try:
        length = int(request.headers.get('Content-Length', '0'))
    except ValueError:
        raise ValueError('The upload size is invalid.') from None
    if not 0 < length <= MAX_BYTES:
        raise ValueError('Select files with a combined size under 4 MB.')
    body = request.rfile.read(length)
    message = BytesParser(policy=policy.default).parsebytes(
        f'Content-Type: {media_type}\r\nMIME-Version: 1.0\r\n\r\n'.encode('ascii') + body,
    )
    if not message.is_multipart() or message.defects:
        raise ValueError('The upload was incomplete. Select the files and try again.')

    files = {}
    for part in message.iter_parts():
        field = part.get_param('name', header='content-disposition')
        if field not in FIELDS or field in files:
            raise ValueError('Select exactly one file in each upload field.')
        label, suffix = FIELDS[field]
        filename = (part.get_filename() or '').replace('\\', '/').rsplit('/', 1)[-1]
        contents = part.get_payload(decode=True)
        signature = b'%PDF-' if suffix == '.pdf' else b'PK\x03\x04'
        if not filename.lower().endswith(suffix) or not contents or not contents.startswith(signature):
            raise ValueError(f'{label} must be a valid {suffix.upper()} file.')
        folder = directory / field
        folder.mkdir()
        path = folder / filename
        path.write_bytes(contents)
        files[field] = path
    for field, (label, _) in FIELDS.items():
        if field not in files:
            raise ValueError(f'Select the {label}.')
    return files


def output_summary(path):
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        rows = list(workbook.worksheets[0].values)
        balances = {row[0]: row[4] for row in rows}
        controls = {row[0]: row[1] for row in workbook['Controls and Sources'].values}
        return {
            'title': rows[0][0],
            'period': f'{controls["Statement Start"]} to {controls["Statement End"]}',
            'filename': report_filename(controls['Property'], controls['Account Last Four'],
                                        controls['Statement Start'], controls['Statement End']),
            'bank': balances['Balance as per Bank Statement'],
            'adjusted': balances['Adjusted Balance'],
            'difference': balances['Difference if any (Adjusted − Bank)'],
            'exceptions': sum(int(row[0].split()[2]) for row in rows
                              if isinstance(row[0], str) and row[0].startswith('Subtotal — ')),
            'paylease': controls['PayLease group matches'],
        }
    finally:
        workbook.close()


class handler(BaseHTTPRequestHandler):
    def respond(self, status, content_type, content, headers=None):
        self.send_response(status)
        self.send_header('Content-Type', content_type)
        self.send_header('Content-Length', str(len(content)))
        self.send_header('Cache-Control', 'no-store')
        self.send_header('X-Content-Type-Options', 'nosniff')
        self.send_header('Referrer-Policy', 'no-referrer')
        self.send_header('Content-Security-Policy',
                         "default-src 'none'; style-src 'unsafe-inline'; script-src 'unsafe-inline'; "
                         "connect-src 'self'; form-action 'self'; frame-ancestors 'none'")
        for key, value in (headers or {}).items():
            self.send_header(key, value)
        self.end_headers()
        self.wfile.write(content)

    def error(self, status, message):
        self.respond(status, 'application/json; charset=utf-8', json.dumps({'error': message}).encode())

    def do_GET(self):
        self.respond(HTTPStatus.OK, 'text/html; charset=utf-8', PAGE.read_bytes())

    def do_POST(self):
        if self.headers.get('X-Matching-Statement') != '1':
            self.error(HTTPStatus.FORBIDDEN, 'Open the Matching Statement page to upload files.')
            return
        try:
            with tempfile.TemporaryDirectory(prefix='matching-statement-') as temporary:
                directory = Path(temporary)
                files = uploaded_files(self, directory)
                output = run(files['statement'], files['register'], files['prior'], files['aging'],
                                 directory / 'Matching_Statement.xlsx')
                summary = output_summary(output)
                report = output.read_bytes()
        except (BadZipFile, InvalidFileException, KeyError, StopIteration, IndexError):
            self.error(HTTPStatus.UNPROCESSABLE_ENTITY,
                       'Check that each file is in the correct field and uses the same Chase and ResMan export layouts.')
            return
        except ValueError as error:
            self.error(HTTPStatus.UNPROCESSABLE_ENTITY, str(error))
            return
        except (socket.timeout, OSError):
            logging.exception('Reconciliation could not read or process the submitted files')
            self.error(HTTPStatus.INTERNAL_SERVER_ERROR, 'Check that all files are readable and run the reconciliation again.')
            return
        self.respond(
            HTTPStatus.OK,
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            report,
            {'X-Matching-Summary': json.dumps(summary, ensure_ascii=True),
             'Content-Disposition': 'attachment; filename=' + summary['filename']},
        )

