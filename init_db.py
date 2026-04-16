import sqlite3
import json
import os

DB_PATH = os.path.join(os.path.dirname(__file__), 'soumu.db')
JSON_PATH = os.path.join(os.path.dirname(__file__), 'excel_data.json')

def init_db():
    if os.path.exists(DB_PATH):
        os.remove(DB_PATH)

    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    c.executescript('''
        CREATE TABLE products (
            code INTEGER PRIMARY KEY,
            name TEXT NOT NULL,
            site TEXT,
            formal_name TEXT
        );
        CREATE TABLE inventory (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            product_name TEXT NOT NULL,
            no INTEGER UNIQUE,
            stock INTEGER DEFAULT 0,
            order_qty INTEGER DEFAULT 10,
            remaining INTEGER DEFAULT 0,
            order_judgment TEXT DEFAULT '',
            supply_qty INTEGER DEFAULT 0,
            usage_qty INTEGER DEFAULT 0,
            order_threshold INTEGER DEFAULT 0,
            note TEXT DEFAULT ''
        );
        CREATE TABLE checkout (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            year INTEGER,
            date TEXT,
            code INTEGER,
            product_name TEXT,
            supply INTEGER DEFAULT 0,
            usage INTEGER DEFAULT 0,
            user_name TEXT,
            note TEXT DEFAULT '',
            FOREIGN KEY (code) REFERENCES products(code)
        );
        CREATE TABLE ringi (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_date TEXT,
            status TEXT DEFAULT 'draft',
            approver TEXT DEFAULT '',
            approved_date TEXT
        );
        CREATE TABLE todos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            description TEXT DEFAULT '',
            priority TEXT DEFAULT 'medium',
            category TEXT DEFAULT '',
            due_date TEXT,
            done INTEGER DEFAULT 0,
            created_at TEXT,
            completed_at TEXT
        );
        CREATE TABLE ringi_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ringi_id INTEGER,
            supplier TEXT,
            product_name TEXT,
            quantity INTEGER,
            unit_price INTEGER DEFAULT 0,
            amount INTEGER DEFAULT 0,
            is_new INTEGER DEFAULT 0,
            purpose TEXT DEFAULT '',
            requester TEXT DEFAULT '',
            approved INTEGER DEFAULT 0,
            search_url TEXT DEFAULT '',
            code INTEGER,
            FOREIGN KEY (ringi_id) REFERENCES ringi(id)
        );
    ''')

    with open(JSON_PATH, 'r', encoding='utf-8') as f:
        data = json.load(f)

    # Products
    for row in data['商品一覧'][1:]:
        code = int(row[0]) if row[0] else None
        if code is None:
            continue
        c.execute('INSERT OR IGNORE INTO products (code, name, site, formal_name) VALUES (?,?,?,?)',
                  (code, row[1], row[2], row[3]))

    # Inventory
    inv_data = data['在庫チェック表(発注担当者用)']
    for row in inv_data[2:]:
        name = row[1]
        if not name:
            continue
        no = int(row[2]) if row[2] else None
        stock = int(row[3]) if row[3] else 0
        order_qty = int(row[4]) if row[4] else 10
        remaining = int(row[5]) if row[5] else 0
        judgment = row[6] if row[6] else ''
        supply = int(row[7]) if row[7] else 0
        usage = int(row[8]) if row[8] else 0
        threshold = int(row[9]) if row[9] else 0
        note = row[10] if row[10] else ''
        c.execute('''INSERT OR IGNORE INTO inventory
            (product_name, no, stock, order_qty, remaining, order_judgment, supply_qty, usage_qty, order_threshold, note)
            VALUES (?,?,?,?,?,?,?,?,?,?)''',
            (name, no, stock, order_qty, remaining, judgment, supply, usage, threshold, note))

    # Checkout
    checkout_data = data['持出明細表']
    for row in checkout_data[4:]:
        date_str = row[2]
        if not date_str:
            continue
        code = int(row[3]) if row[3] else None
        if code is None:
            continue
        year = int(row[1]) if row[1] else 2026
        usage = int(row[6]) if row[6] else 0
        supply = int(row[5]) if row[5] else 0
        c.execute('''INSERT INTO checkout (year, date, code, product_name, supply, usage, user_name, note)
            VALUES (?,?,?,?,?,?,?,?)''',
            (year, date_str[:10], code, row[4], supply, usage, row[7], row[8]))

    conn.commit()
    conn.close()
    print(f'Database initialized: {DB_PATH}')

if __name__ == '__main__':
    init_db()
