"""Extract SKU / size / quantity lines from sales or inventory reports.

Reads PDF, Excel (.xlsx/.xlsm/.xls) and CSV files of any layout and produces a
clean workbook with two columns:

    Product   = 7-digit SKU starting with "10" + normalised size (e.g. 1012345M, 1012345-34)
    Quantity  = quantity for that product (summed if it appears more than once)

Lines where a SKU was found but the size or quantity couldn't be worked out are
written to a "Needs Review" sheet instead of being silently dropped.

Command line:  python converter.py report1.pdf report2.xlsx [-o OUTPUT_DIR] [--combined]
"""
import argparse
import csv
import os
import re
from collections import OrderedDict
from dataclasses import dataclass

SKU_RE = re.compile(r'(?<!\d)10\d{5}(?!\d)')
TOTAL_RE = re.compile(r'\b(sub\s*-?\s*)?totals?\b', re.I)
QTY_RE = re.compile(r'(\()?(-)?(\d{1,3}(?:,\d{3})+|\d{1,5})(\.0+)?(-)?(\))?')
WAIST_RE = re.compile(r'(-)?(W(?:AIST)?\s*)?([2-5]\d)(\s*W)?(?:\s*[X/\-]\s*(\d{2})\s*L?)?')
EXTRA_RE = re.compile(r'(\d)X(?:L|LG|LARGE)?|XL(\d)|(X{2,6})L(?:G|ARGE)?')

LETTER_SIZES = {
    'XS': 'XS', 'XSM': 'XS', 'XSML': 'XS', 'XSMALL': 'XS', 'EXTRASMALL': 'XS',
    'S': 'S', 'SM': 'S', 'SML': 'S', 'SMA': 'S', 'SMALL': 'S',
    'M': 'M', 'MD': 'M', 'MED': 'M', 'MEDIUM': 'M',
    'L': 'L', 'LG': 'L', 'LRG': 'L', 'LGE': 'L', 'LARGE': 'L',
    'XL': 'XL', 'XLG': 'XL', 'XLRG': 'XL', 'XLARGE': 'XL', 'EXTRALARGE': 'XL',
}
# First word of a two-word size such as "X LARGE", "EXTRA LARGE", "W 34"
JOIN_FIRST = {'X', 'XX', 'XXX', 'XXXX', 'EXTRA', 'EX', 'W', 'WAIST'}

SKU_WORDS = {'sku', 'item', 'style', 'product', 'article', 'code', 'upc', 'number', 'no', 'num'}
OTHER_HEADER_WORDS = {'desc', 'description', 'color', 'colour', 'price', 'cost', 'retail',
                      'vendor', 'dept', 'department', 'class', 'name', 'store', 'location'}
SIZE_WORDS = {'size', 'sz', 'sizes', 'waist', 'variant', 'varian'}
# Quantity columns that are NOT the stock/sold figure we want
QTY_QUALIFIERS = {'order', 'transit', 'oo', 'allocated', 'reserved', 'rsrvd', 'rtv', 'non',
                  'sellable', 'damaged', 'tfr', 'cust', 'fin', 'return', 'returns', 'pack'}
NOT_QTY_RE = re.compile(r'price|cost|retail|amount|value|msrp|dollar|revenue|margin|wholesale|\$|%', re.I)


@dataclass
class Cell:
    text: str
    x0: float
    x1: float
    h: float = 0.0

    @property
    def center(self):
        return (self.x0 + self.x1) / 2


# ---------------------------------------------------------------- parsing helpers

def normalize_size(token, allow_plain_number=False):
    """Return S/M/L/XL/XL2/XL3/-34 style size, or None if the token isn't a size.

    Plain numbers ("34") are only treated as waist sizes when allow_plain_number is
    True (i.e. we know we're looking at a size column); "W34", "34W", "34x32", "-34"
    are always accepted.
    """
    t = token.strip().upper().replace('.', '').strip('()[]:,;#')
    t = re.sub(r'^(SIZE|SZ)[:\s\-]*', '', t)
    if not t:
        return None
    compact = re.sub(r'[\s\-_]', '', t)
    if compact in LETTER_SIZES:
        return LETTER_SIZES[compact]
    m = EXTRA_RE.fullmatch(compact)
    if m:
        n = int(m.group(1) or m.group(2)) if (m.group(1) or m.group(2)) else len(m.group(3))
        if 1 <= n <= 6:
            return 'XL' if n == 1 else f'XL{n}'
        return None
    m = WAIST_RE.fullmatch(t)
    if m:
        explicit = m.group(1) or m.group(2) or m.group(4) or m.group(5)
        if explicit or allow_plain_number:
            return f'-{m.group(3)}'
    return None


def parse_qty(text, allow_decimal=True):
    t = text.strip()
    m = QTY_RE.fullmatch(t)
    if not m:
        return None
    if (m.group(1) is None) != (m.group(6) is None):
        return None
    if m.group(4) and (not allow_decimal or ',' in m.group(3)):
        return None  # "6,325.00" is money, not a quantity
    n = int(m.group(3).replace(',', ''))
    if m.group(1) or m.group(2) or m.group(5):
        n = -n
    return n


def split_tokens(text):
    return [t for t in re.split(r'[\s,;|]+', text) if t]


def find_sizes(tokens, allow_plain_number=False):
    """tokens: list of strings. Returns [(index, size)] for every size found."""
    found = []
    i = 0
    while i < len(tokens):
        tok = tokens[i]
        if i + 1 < len(tokens) and tok.upper().strip('.-') in JOIN_FIRST:
            s = normalize_size(tok + ' ' + tokens[i + 1], allow_plain_number)
            if s:
                found.append((i, s))
                i += 2
                continue
        s = normalize_size(tok, allow_plain_number)
        if s:
            found.append((i, s))
        i += 1
    return found


def qty_priority(label, next_label=''):
    """Lower = more likely to be the quantity column. None = not a quantity column."""
    low = label.lower()
    if NOT_QTY_RE.search(low):
        return None
    words = set(re.findall(r'[a-z]+', low))
    # PDF may split "Unit Price" / "Total Cost" into separate words
    if words <= {'unit', 'units', 'total', 'ext', 'extended'} and re.match(
            r'\s*(price|cost|retail|amount|value)', next_label.lower()):
        return None
    if words & QTY_QUALIFIERS:
        return 7  # "Qty On Order", "In Transit Qty", "Pack" etc. - only if nothing better exists
    if words & {'hand', 'oh', 'soh', 'onhand', 'available', 'avail', 'ats'}:
        return 1
    if words & {'qty', 'qnty', 'quantity', 'quan', 'qtys', 'quantities'}:
        return 2
    if words & {'units', 'unit'}:
        return 3
    if 'sold' in words:
        return 4
    if words & {'pcs', 'pieces', 'count', 'stock', 'inv', 'balance', 'bal'}:
        return 5
    if words & {'total', 'ordered', 'shipped', 'ord', 'shp'}:
        return 6
    return None


def merge_phrases(cells):
    """Join PDF words that sit right next to each other ("On" "Hand" -> "On Hand")."""
    out = []
    for c in cells:
        if out and c.x0 - out[-1].x1 <= max(3.0, 0.6 * (c.h or 10)):
            p = out[-1]
            out[-1] = Cell(p.text + ' ' + c.text, p.x0, c.x1, max(p.h, c.h))
        else:
            out.append(Cell(c.text, c.x0, c.x1, c.h))
    return out


# ---------------------------------------------------------------- extraction

class Extractor:
    def __init__(self, qty_column=None, keep_zero=False):
        self.qty_column = (qty_column or '').strip().lower() or None
        self.keep_zero = keep_zero
        self.products = OrderedDict()
        self.seen_skus = set()  # SKUs read with a size, including 0-quantity ones
        self.review = []

    # -- output
    def add(self, sku, size, qty):
        self.seen_skus.add(sku)
        if qty == 0 and not self.keep_zero:
            return
        key = sku + size
        self.products[key] = self.products.get(key, 0) + qty

    def flag(self, source, sku, size, qty, reason, text):
        self.review.append({'source': source, 'sku': sku, 'size': size or '',
                            'qty': '' if qty is None else qty, 'reason': reason, 'text': text})

    # -- header detection
    def _header(self, cells, is_pdf):
        # Size-grid header: S M L XL ...  or  28 30 32 34 ...
        size_cols = {}
        for i, c in enumerate(cells):
            s = normalize_size(c.text)
            if s:
                size_cols[i] = s
        if len(size_cols) < 2 or len(size_cols) * 4 < len(cells):
            nums = [(i, int(c.text.strip())) for i, c in enumerate(cells)
                    if re.fullmatch(r'[2-5]\d', c.text.strip())]
            if (len(nums) >= 3 and len(nums) * 2 >= len(cells)
                    and all(b[1] > a[1] for a, b in zip(nums, nums[1:]))):
                size_cols = {i: f'-{n}' for i, n in nums}
            else:
                size_cols = {}
        if size_cols:
            return {'kind': 'matrix', 'anchors': cells, 'size_cols': size_cols}

        # Column-name header: Item | Description | Size | Qty ...
        phrases = merge_phrases(cells) if is_pdf else cells
        if len(phrases) < 2:
            return None
        keywords = 0
        qty_idx, qty_best, size_idx = None, 99, None
        for i, c in enumerate(phrases):
            low = c.text.lower()
            nxt = phrases[i + 1].text if i + 1 < len(phrases) else ''
            words = set(re.findall(r'[a-z]+', low))
            pr = qty_priority(c.text, nxt)
            if self.qty_column and self.qty_column in low:
                pr = 0
            if pr is not None:
                keywords += 1
                if pr < qty_best:
                    qty_idx, qty_best = i, pr
            elif words & SIZE_WORDS:
                keywords += 1
                if size_idx is None:
                    size_idx = i
            elif words & (SKU_WORDS | OTHER_HEADER_WORDS):
                keywords += 1
        if keywords >= 2 and (qty_idx is not None or size_idx is not None):
            return {'kind': 'columns', 'anchors': phrases,
                    'qty': qty_idx, 'qty_pr': qty_best, 'size': size_idx}
        return None

    @staticmethod
    def _merge_headers(old, new):
        """Second line of a multi-line header: add to the first rather than replace it."""
        off = len(old['anchors'])
        merged = dict(old, anchors=old['anchors'] + new['anchors'])
        if new['qty'] is not None and (old['qty'] is None or new['qty_pr'] < old['qty_pr']):
            merged['qty'], merged['qty_pr'] = new['qty'] + off, new['qty_pr']
        if old['size'] is None and new['size'] is not None:
            merged['size'] = new['size'] + off
        return merged

    @staticmethod
    def _anchor(ctx, cell, tol):
        """Index of the header cell this cell sits under (left-, right- or centre-aligned)."""
        best, dist = None, None
        for i, a in enumerate(ctx['anchors']):
            d = min(abs(a.x0 - cell.x0), abs(a.x1 - cell.x1), abs(a.center - cell.center))
            if dist is None or d < dist:
                best, dist = i, d
        if dist is None or (tol is not None and dist > tol):
            return None
        return best

    # -- main loop
    def process_rows(self, rows, source, is_pdf=False):
        tol = 80.0 if is_pdf else 0.01
        ctx = None
        pending = None  # SKU row that had no size; following lines may carry the sizes

        def flush():
            nonlocal pending
            if pending and not pending['resolved']:
                self.flag(source, pending['sku'], None, pending['qty'],
                          'Size not found', pending['text'])
            pending = None

        prev_was_header = False
        for cells in rows:
            cells = [c for c in cells if c.text.strip()]
            if not cells:
                continue
            line = ' '.join(c.text for c in cells)
            hits = [(i, m) for i, c in enumerate(cells) for m in SKU_RE.finditer(c.text)]
            # The same SKU repeated on a line (e.g. in the description and a code column)
            # is still one item.
            if len({m.group() for _, m in hits}) == 1 and len(hits) > 1:
                exact = [h for h in hits if cells[h[0]].text.strip() == h[1].group()]
                hits = [exact[0] if exact else hits[0]]

            if not hits:
                hdr = self._header(cells, is_pdf)
                if hdr:
                    if hdr['kind'] == 'columns':
                        flush()
                        if prev_was_header and ctx and ctx['kind'] == 'columns':
                            hdr = self._merge_headers(ctx, hdr)
                    ctx = hdr
                elif pending:
                    self._continuation(cells, ctx, tol, pending)
                prev_was_header = bool(hdr)
                continue
            prev_was_header = False

            if TOTAL_RE.search(line):
                continue
            flush()

            # One row may hold several SKUs (side-by-side layouts): split it into segments.
            starts = sorted({i for i, _ in hits})
            for n, (ci, m) in enumerate(hits):
                if n and hits[n - 1][0] == ci:
                    continue  # second SKU inside the same cell - ignore
                k = starts.index(ci)
                lo = 0 if k == 0 else ci
                hi = starts[k + 1] if k + 1 < len(starts) else len(cells)
                seg = cells[lo:hi]
                result = self._segment(seg, ci - lo, m, ctx, tol, source, line)
                if result is not None:
                    flush()
                    pending = {'sku': m.group(), 'qty': result['qty'], 'text': line, 'resolved': False}
        flush()
        # A bare mention of a SKU (group heading, wrapped description) with no size and no
        # quantity isn't a problem when that SKU was picked up properly elsewhere.
        self.review = [r for r in self.review
                       if not (r['reason'] == 'Size not found' and r['qty'] == ''
                               and r['sku'] in self.seen_skus)]

    def _segment(self, cells, sku_ci, m, ctx, tol, source, line):
        """Handle one SKU. Returns None when done, or {'qty': guess} when the size is
        still unknown so later lines can supply it."""
        sku = m.group()
        sku_cell = cells[sku_ci]
        mapped = [self._anchor(ctx, c, tol) if ctx else None for c in cells]

        size, size_ci, size_tok = None, None, None

        # 1. a "Size" column
        if ctx and ctx['kind'] == 'columns' and ctx['size'] is not None:
            toks, owners = [], []
            for i, c in enumerate(cells):
                if i != sku_ci and mapped[i] == ctx['size']:
                    for t in split_tokens(c.text):
                        toks.append(t)
                        owners.append(i)
            found = find_sizes(toks, allow_plain_number=True)
            if found:
                size, size_ci = found[-1][1], owners[found[-1][0]]

        # 2. size glued to the SKU: 1012345-M, 1012345XL, 1012345-34, "1012345 Tee Black M"
        if size is None:
            rest = sku_cell.text[m.end():]
            # "1003169#NAVY\34", "1007238#CA BLACK\L": size after the backslash
            after = rest + ' ' + ' '.join(c.text for c in cells[sku_ci + 1:sku_ci + 4])
            bm = re.match(r'\s*#[^\\\s]*(?:\s+[^\\\s]+){0,3}?\\(\S+)', after)
            if bm:
                size = normalize_size(bm.group(1), allow_plain_number=True)
            # "1012345-M", "1012345XL", "1090123-34": size stuck onto the SKU
            glued = re.match(r'\S*', rest).group()
            if size is None and glued:
                parts = [p for p in re.split(r'[-_/\\#.]+', glued) if p]
                if parts:
                    size = normalize_size(parts[-1], allow_plain_number=glued[0] in '-_/\\#.')
            # "1007287 XL Black Polo" inside one cell
            if size is None:
                toks = [t for t in re.split(r'[\s\-_/]+', rest.strip(' -_/')) if t]
                found = find_sizes(toks)
                if found:
                    size = found[-1][1]
            if size is not None:
                size_ci = sku_ci

        # 3. size-grid row: quantities sit under S / M / L ... headings
        if size is None and ctx and ctx['kind'] == 'matrix':
            got = self._matrix(cells, mapped, ctx, sku, skip=sku_ci)
            if got:
                return None

        # 4. any size-looking word elsewhere on the line (the last one wins)
        tokens = []  # (cell index, token)
        for i, c in enumerate(cells):
            if i == sku_ci:
                continue
            if ctx and ctx['kind'] == 'columns' and ctx['qty'] is not None and mapped[i] == ctx['qty']:
                continue
            tokens.extend((i, t) for t in split_tokens(c.text))
        if size is None:
            found = find_sizes([t for _, t in tokens])
            if found:
                size_tok = found[-1][0]
                size = found[-1][1]
                size_ci = tokens[size_tok][0]

        qty = self._quantity(cells, mapped, ctx, sku_ci, size_ci, size_tok)

        if size is None:
            return {'qty': qty}  # pending - later lines may hold the size
        if qty is None:
            self.flag(source, sku, size, None, 'Quantity not found', line)
            return None
        self.add(sku, size, qty)
        return None

    def _quantity(self, cells, mapped, ctx, sku_ci, size_ci, size_tok):
        # A known quantity column wins
        if ctx and ctx['kind'] == 'columns' and ctx['qty'] is not None:
            for i, c in enumerate(cells):
                if i != sku_ci and mapped[i] == ctx['qty']:
                    q = parse_qty(c.text)
                    if q is not None:
                        return q
        # Otherwise: first whole number after the size (or SKU), then anything before it.
        # (Same token list as _segment builds, so size_tok indexes line up.)
        toks = []
        for i, c in enumerate(cells):
            if i == sku_ci:
                continue
            if ctx and ctx['kind'] == 'columns' and ctx['qty'] is not None and mapped[i] == ctx['qty']:
                continue
            for t in split_tokens(c.text):
                toks.append((i, t))
        if size_tok is not None:
            after = toks[size_tok + 1:] if size_tok + 1 < len(toks) else []
            before = toks[:size_tok]
            # a two-word size ("X LARGE") occupies the next token too
            if after and size_tok + 1 < len(toks) and normalize_size(
                    toks[size_tok][1] + ' ' + toks[size_tok + 1][1]):
                after = after[1:]
        else:
            pivot = max(sku_ci, size_ci if size_ci is not None else sku_ci)
            after = [x for x in toks if x[0] > pivot]
            before = [x for x in toks if x[0] <= pivot and x[0] != size_ci]
        for allow_dec in (False, True):
            for group in (after, before):
                for _, t in group:
                    q = parse_qty(t, allow_decimal=allow_dec)
                    if q is not None:
                        return q
        return None

    def _matrix(self, cells, mapped, ctx, sku, skip=None):
        got = 0
        for i, c in enumerate(cells):
            if i == skip or mapped[i] not in ctx['size_cols']:
                continue
            q = parse_qty(c.text)
            if q is not None:
                self.add(sku, ctx['size_cols'][mapped[i]], q)
                got += 1
        return got

    def _continuation(self, cells, ctx, tol, pending):
        """A line without a SKU right after a SKU line that had no size."""
        if TOTAL_RE.search(' '.join(c.text for c in cells)):
            return
        if ctx and ctx['kind'] == 'matrix':
            mapped = [self._anchor(ctx, c, tol) for c in cells]
            if self._matrix(cells, mapped, ctx, pending['sku']):
                pending['resolved'] = True
                return
        tokens = [t for c in cells for t in split_tokens(c.text)]
        allow_plain = False
        if ctx and ctx['kind'] == 'columns' and ctx['size'] is not None:
            allow_plain = any(self._anchor(ctx, c, tol) == ctx['size'] for c in cells)
        found = find_sizes(tokens, allow_plain_number=allow_plain)
        if not found:
            return
        idx, size = found[-1]
        for t in tokens[idx + 1:] + tokens[:idx]:
            q = parse_qty(t)
            if q is not None and not (allow_plain and normalize_size(t, True) == size):
                self.add(pending['sku'], size, q)
                pending['resolved'] = True
                return


# ---------------------------------------------------------------- readers

def _fmt_value(v):
    if v is None:
        return ''
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else repr(v)
    return str(v).strip()


def read_pdf(path):
    """Returns (rows, raw_rows). rows = list of lines, each a list of word Cells."""
    import pdfplumber
    rows = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            words = page.extract_words(x_tolerance=2, y_tolerance=3, keep_blank_chars=False)
            words.sort(key=lambda w: (w['top'], w['x0']))
            line, line_top = [], None
            for w in words:
                if line_top is None or abs(w['top'] - line_top) > 3:
                    if line:
                        rows.append(sorted(line, key=lambda c: c.x0))
                    line, line_top = [], w['top']
                line.append(Cell(w['text'], w['x0'], w['x1'], w['bottom'] - w['top']))
            if line:
                rows.append(sorted(line, key=lambda c: c.x0))
    return rows


def read_excel(path):
    """Returns list of (sheet_name, rows); each row is a list of Cells positioned by column."""
    ext = os.path.splitext(path)[1].lower()
    sheets = []
    if ext == '.csv':
        with open(path, newline='', encoding='utf-8-sig', errors='replace') as f:
            sample = f.read(4096)
            f.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=',;\t|')
            except csv.Error:
                dialect = csv.excel
            data = list(csv.reader(f, dialect))
        sheets.append(('csv', [[Cell(v.strip(), j, j + 1) for j, v in enumerate(r) if v.strip()]
                               for r in data]))
    elif ext == '.xls':
        import xlrd
        book = xlrd.open_workbook(path)
        for sh in book.sheets():
            rows = []
            for r in range(sh.nrows):
                rows.append([Cell(_fmt_value(v), j, j + 1)
                             for j, v in enumerate(sh.row_values(r)) if _fmt_value(v)])
            sheets.append((sh.name, rows))
    else:
        import openpyxl
        book = openpyxl.load_workbook(path, data_only=True, read_only=True)
        for sh in book.worksheets:
            rows = []
            for row in sh.iter_rows(values_only=True):
                rows.append([Cell(_fmt_value(v), j, j + 1)
                             for j, v in enumerate(row) if _fmt_value(v)])
            sheets.append((sh.title, rows))
        book.close()
    return sheets


def extract_file(path, qty_column=None, keep_zero=False):
    """Returns (Extractor, raw_rows_or_None, warning_or_None)."""
    ex = Extractor(qty_column=qty_column, keep_zero=keep_zero)
    name = os.path.basename(path)
    ext = os.path.splitext(path)[1].lower()
    raw, warning = None, None
    if ext == '.pdf':
        rows = read_pdf(path)
        raw = [[c.text for c in merge_phrases(r)] for r in rows]
        if not rows:
            warning = 'No text found - this looks like a scanned/image PDF and would need OCR.'
        ex.process_rows(rows, name, is_pdf=True)
    elif ext in ('.xlsx', '.xlsm', '.xls', '.csv'):
        for sheet, rows in read_excel(path):
            ex.process_rows(rows, f'{name} [{sheet}]')
    else:
        raise ValueError(f'Unsupported file type: {ext}')
    if not ex.products and not warning:
        warning = 'No SKUs (10xxxxx) with a size and quantity were found.'
    return ex, raw, warning


# ---------------------------------------------------------------- writers

def write_clean(path, products, review):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment

    wb = Workbook()
    ws = wb.active
    ws.title = 'Products'
    ws.append(['Product', 'Quantity'])
    for prod, qty in products.items():
        ws.append([prod, qty])
    head_font = Font(bold=True, color='FFFFFF')
    head_fill = PatternFill('solid', fgColor='305496')
    for c in ws[1]:
        c.font = head_font
        c.fill = head_fill
        c.alignment = Alignment(horizontal='center')
    for row in ws.iter_rows(min_row=2, max_col=1):
        row[0].number_format = '@'
    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 12
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f'A1:B{ws.max_row}'

    if review:
        rs = wb.create_sheet('Needs Review')
        rs.append(['Source', 'SKU', 'Size', 'Quantity', 'Problem', 'Original line'])
        for r in review:
            rs.append([r['source'], r['sku'], r['size'], r['qty'], r['reason'], r['text']])
        for c in rs[1]:
            c.font = head_font
            c.fill = PatternFill('solid', fgColor='C65911')
        for col, w in zip('ABCDEF', (28, 10, 8, 10, 20, 90)):
            rs.column_dimensions[col].width = w
        rs.freeze_panes = 'A2'
    wb.save(path)


def write_raw(path, raw_rows):
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = 'PDF data'
    for r in raw_rows:
        ws.append(r)
    wb.save(path)


def convert(paths, out_dir=None, combined=False, qty_column=None, keep_zero=False,
            save_raw=False, log=print):
    """Process every file. Returns list of output paths written."""
    written = []
    all_products, all_review = OrderedDict(), []
    for p in paths:
        name = os.path.basename(p)
        try:
            ex, raw, warning = extract_file(p, qty_column, keep_zero)
        except Exception as e:  # keep going with the other files
            log(f'ERROR  {name}: {e}')
            continue
        target_dir = out_dir or os.path.dirname(os.path.abspath(p))
        stem = os.path.splitext(name)[0]
        out = os.path.join(target_dir, f'{stem}_clean.xlsx')
        try:
            write_clean(out, ex.products, ex.review)
            written.append(out)
            if save_raw and raw is not None:
                raw_out = os.path.join(target_dir, f'{stem}_raw.xlsx')
                write_raw(raw_out, raw)
                written.append(raw_out)
        except PermissionError:
            log(f'ERROR  {name}: cannot write {out} - is it open in Excel?')
            continue
        msg = f'OK     {name}: {len(ex.products)} products, {sum(ex.products.values())} units'
        if ex.review:
            msg += f', {len(ex.review)} line(s) need review'
        log(msg + f'  ->  {os.path.basename(out)}')
        if warning:
            log(f'       note: {warning}')
        for k, v in ex.products.items():
            all_products[k] = all_products.get(k, 0) + v
        all_review.extend(ex.review)

    if combined and len(paths) > 1 and all_products:
        from datetime import datetime
        target_dir = out_dir or os.path.dirname(os.path.abspath(paths[0]))
        out = os.path.join(target_dir, f'Combined_clean_{datetime.now():%Y%m%d_%H%M%S}.xlsx')
        write_clean(out, all_products, all_review)
        written.append(out)
        log(f'OK     Combined file: {len(all_products)} products  ->  {os.path.basename(out)}')
    return written


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('files', nargs='+')
    ap.add_argument('-o', '--output-dir')
    ap.add_argument('--combined', action='store_true', help='also write one combined workbook')
    ap.add_argument('--qty-column', help='header text of the quantity column, if auto-detect picks the wrong one')
    ap.add_argument('--keep-zero', action='store_true', help='keep products whose quantity is 0')
    ap.add_argument('--raw', action='store_true', help='also save the raw PDF-to-Excel conversion')
    a = ap.parse_args()
    convert(a.files, a.output_dir, a.combined, a.qty_column, a.keep_zero, a.raw)


if __name__ == '__main__':
    main()
