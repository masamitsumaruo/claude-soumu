import openpyxl
import sqlite3
import os

EXCEL_PATH = r'C:\Users\acuma\Dropbox\AI活用\総務資料\事務用品持出明細表.xlsx'
DB_PATH = os.path.join(os.path.dirname(__file__), 'soumu.db')


def import_products():
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    ws = wb['商品一覧']

    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    inserted = 0
    updated = 0
    skipped = 0

    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        code, name, site, formal = row[0], row[1], row[2], row[3]
        if code is None or name is None:
            skipped += 1
            continue
        try:
            code = int(code)
        except (ValueError, TypeError):
            skipped += 1
            continue

        existing = c.execute('SELECT code FROM products WHERE code=?', (code,)).fetchone()
        if existing:
            c.execute('UPDATE products SET name=?, site=?, formal_name=? WHERE code=?',
                      (name, site or '', formal or '', code))
            updated += 1
        else:
            c.execute('INSERT INTO products (code, name, site, formal_name) VALUES (?,?,?,?)',
                      (code, name, site or '', formal or ''))
            inserted += 1

    conn.commit()
    total = c.execute('SELECT COUNT(*) FROM products').fetchone()[0]
    conn.close()
    print(f'Inserted: {inserted}, Updated: {updated}, Skipped: {skipped}, Total products in DB: {total}')


if __name__ == '__main__':
    import_products()
