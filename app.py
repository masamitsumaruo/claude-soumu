import sqlite3
import os
import csv
import io
import urllib.parse
from datetime import datetime, date
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_file, make_response

app = Flask(__name__)
app.secret_key = 'soumu-bunbougu-2026'
DB_PATH = os.path.join(os.path.dirname(__file__), 'soumu.db')


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


# ============ Dashboard ============
@app.route('/')
def index():
    db = get_db()
    product_count = db.execute('SELECT COUNT(*) FROM products').fetchone()[0]
    inv_count = db.execute('SELECT COUNT(*) FROM inventory').fetchone()[0]
    checkout_count = db.execute('SELECT COUNT(*) FROM checkout').fetchone()[0]
    ringi_count = db.execute('SELECT COUNT(*) FROM ringi').fetchone()[0]
    todo_count = db.execute('SELECT COUNT(*) FROM todos WHERE done=0').fetchone()[0]
    alerts = db.execute("SELECT product_name, remaining, order_threshold FROM inventory WHERE remaining <= order_threshold AND order_threshold > 0").fetchall()
    db.close()
    return render_template('index.html', product_count=product_count, inv_count=inv_count,
                           checkout_count=checkout_count, ringi_count=ringi_count,
                           todo_count=todo_count, alerts=alerts)


# ============ Products (商品一覧) ============
@app.route('/products')
def products():
    db = get_db()
    rows = db.execute('SELECT * FROM products ORDER BY code').fetchall()
    db.close()
    return render_template('products.html', products=rows)


@app.route('/products/add', methods=['POST'])
def product_add():
    db = get_db()
    code = int(request.form['code'])
    name = request.form['name']
    site = request.form['site']
    formal = request.form['formal_name']
    try:
        db.execute('INSERT INTO products (code, name, site, formal_name) VALUES (?,?,?,?)',
                   (code, name, site, formal))
        db.commit()
        flash('商品を追加しました', 'success')
    except sqlite3.IntegrityError:
        flash('このコードは既に存在します', 'danger')
    db.close()
    return redirect(url_for('products'))


@app.route('/products/edit/<int:code>', methods=['POST'])
def product_edit(code):
    db = get_db()
    db.execute('UPDATE products SET name=?, site=?, formal_name=? WHERE code=?',
               (request.form['name'], request.form['site'], request.form['formal_name'], code))
    db.commit()
    db.close()
    flash('商品を更新しました', 'success')
    return redirect(url_for('products'))


@app.route('/products/delete/<int:code>', methods=['POST'])
def product_delete(code):
    db = get_db()
    db.execute('DELETE FROM products WHERE code=?', (code,))
    db.commit()
    db.close()
    flash('商品を削除しました', 'success')
    return redirect(url_for('products'))


# ============ Checkout (持出明細表) ============
@app.route('/checkout')
def checkout():
    db = get_db()
    rows = db.execute('SELECT * FROM checkout ORDER BY date DESC').fetchall()
    products = db.execute('SELECT code, name FROM products ORDER BY name').fetchall()
    db.close()
    return render_template('checkout.html', rows=rows, products=products)


@app.route('/checkout/add', methods=['POST'])
def checkout_add():
    db = get_db()
    code = int(request.form['code'])
    product = db.execute('SELECT name FROM products WHERE code=?', (code,)).fetchone()
    product_name = product['name'] if product else request.form.get('product_name', '')
    checkout_date = request.form['date']
    year = int(checkout_date[:4]) if checkout_date else date.today().year
    usage = int(request.form['usage']) if request.form['usage'] else 0
    supply = int(request.form.get('supply', 0) or 0)
    user_name = request.form['user_name']

    db.execute('''INSERT INTO checkout (year, date, code, product_name, supply, usage, user_name, note)
        VALUES (?,?,?,?,?,?,?,?)''',
        (year, checkout_date, code, product_name, supply, usage, user_name, request.form.get('note', '')))

    # Update inventory remaining
    inv = db.execute('SELECT * FROM inventory WHERE no=?', (code,)).fetchone()
    if inv:
        new_remaining = inv['remaining'] - usage + supply
        judgment = '注文して下さい' if new_remaining <= inv['order_threshold'] and inv['order_threshold'] > 0 else ''
        db.execute('UPDATE inventory SET remaining=?, usage_qty=usage_qty+?, supply_qty=supply_qty+?, order_judgment=? WHERE no=?',
                   (new_remaining, usage, supply, judgment, code))

    db.commit()
    db.close()
    flash('持出を記録しました', 'success')
    return redirect(url_for('checkout'))


@app.route('/checkout/delete/<int:id>', methods=['POST'])
def checkout_delete(id):
    db = get_db()
    db.execute('DELETE FROM checkout WHERE id=?', (id,))
    db.commit()
    db.close()
    flash('記録を削除しました', 'success')
    return redirect(url_for('checkout'))


# ============ Inventory (在庫チェック表) ============
@app.route('/inventory')
def inventory():
    db = get_db()
    rows = db.execute('SELECT * FROM inventory ORDER BY no').fetchall()
    db.close()
    return render_template('inventory.html', rows=rows)


@app.route('/inventory/edit/<int:id>', methods=['POST'])
def inventory_edit(id):
    db = get_db()
    remaining = int(request.form['remaining'])
    threshold = int(request.form['order_threshold'])
    order_qty = int(request.form['order_qty'])
    judgment = '注文して下さい' if remaining <= threshold and threshold > 0 else ''
    db.execute('''UPDATE inventory SET remaining=?, order_threshold=?, order_qty=?, order_judgment=?, note=? WHERE id=?''',
               (remaining, threshold, order_qty, judgment, request.form.get('note', ''), id))
    db.commit()
    db.close()
    flash('在庫を更新しました', 'success')
    return redirect(url_for('inventory'))


@app.route('/inventory/add', methods=['POST'])
def inventory_add():
    db = get_db()
    max_no = db.execute('SELECT COALESCE(MAX(no),0) FROM inventory').fetchone()[0]
    db.execute('''INSERT INTO inventory (product_name, no, stock, order_qty, remaining, order_threshold, note)
        VALUES (?,?,?,?,?,?,?)''',
        (request.form['product_name'], max_no + 1, int(request.form.get('stock', 0)),
         int(request.form.get('order_qty', 10)), int(request.form.get('remaining', 0)),
         int(request.form.get('order_threshold', 0)), request.form.get('note', '')))
    db.commit()
    db.close()
    flash('在庫項目を追加しました', 'success')
    return redirect(url_for('inventory'))


@app.route('/inventory/delete/<int:id>', methods=['POST'])
def inventory_delete(id):
    db = get_db()
    db.execute('DELETE FROM inventory WHERE id=?', (id,))
    db.commit()
    db.close()
    flash('在庫項目を削除しました', 'success')
    return redirect(url_for('inventory'))


# ============ Ringi (稟議書) ============
@app.route('/ringi')
def ringi_list():
    db = get_db()
    ringis = db.execute('SELECT * FROM ringi ORDER BY id DESC').fetchall()
    db.close()
    return render_template('ringi_list.html', ringis=ringis)


@app.route('/ringi/create')
def ringi_create_form():
    db = get_db()
    # Get items that need ordering
    alerts = db.execute("""SELECT i.no, i.product_name, i.remaining, i.order_qty, i.order_threshold,
        p.site, p.formal_name, p.code
        FROM inventory i LEFT JOIN products p ON i.no = p.code
        WHERE (i.remaining <= i.order_threshold AND i.order_threshold > 0) OR i.order_judgment != ''
        ORDER BY i.no""").fetchall()
    products = db.execute('SELECT * FROM products ORDER BY name').fetchall()
    db.close()
    return render_template('ringi_create.html', alerts=alerts, products=products)


@app.route('/ringi/create', methods=['POST'])
def ringi_create():
    db = get_db()
    today = date.today().isoformat()
    cursor = db.execute('INSERT INTO ringi (created_date, status) VALUES (?, ?)', (today, 'draft'))
    ringi_id = cursor.lastrowid

    codes = request.form.getlist('code[]')
    suppliers = request.form.getlist('supplier[]')
    names = request.form.getlist('product_name[]')
    quantities = request.form.getlist('quantity[]')
    prices = request.form.getlist('unit_price[]')
    is_news = request.form.getlist('is_new[]') if 'is_new[]' in request.form else []
    purposes = request.form.getlist('purpose[]')
    requesters = request.form.getlist('requester[]')
    urls = request.form.getlist('search_url[]')

    for i in range(len(names)):
        if not names[i]:
            continue
        qty = int(quantities[i]) if quantities[i] else 0
        price = int(prices[i]) if prices[i] else 0
        code = int(codes[i]) if codes[i] else 0
        is_new = 1 if str(i) in is_news or (i < len(is_news) and is_news[i] == '1') else 0
        db.execute('''INSERT INTO ringi_items
            (ringi_id, supplier, product_name, quantity, unit_price, amount, is_new, purpose, requester, search_url, code)
            VALUES (?,?,?,?,?,?,?,?,?,?,?)''',
            (ringi_id, suppliers[i] if i < len(suppliers) else '',
             names[i], qty, price, qty * price, is_new,
             purposes[i] if i < len(purposes) else '',
             requesters[i] if i < len(requesters) else '',
             urls[i] if i < len(urls) else '', code))

    db.commit()
    db.close()
    flash('稟議書を作成しました', 'success')
    return redirect(url_for('ringi_detail', id=ringi_id))


@app.route('/ringi/<int:id>')
def ringi_detail(id):
    db = get_db()
    ringi = db.execute('SELECT * FROM ringi WHERE id=?', (id,)).fetchone()
    items = db.execute('SELECT * FROM ringi_items WHERE ringi_id=?', (id,)).fetchall()
    db.close()
    if not ringi:
        flash('稟議書が見つかりません', 'danger')
        return redirect(url_for('ringi_list'))
    return render_template('ringi_detail.html', ringi=ringi, items=items)


@app.route('/ringi/<int:id>/approve', methods=['POST'])
def ringi_approve(id):
    db = get_db()
    approver = request.form.get('approver', '')
    today = date.today().isoformat()
    db.execute('UPDATE ringi SET status=?, approver=?, approved_date=? WHERE id=?',
               ('approved', approver, today, id))

    # Update inventory
    items = db.execute('SELECT * FROM ringi_items WHERE ringi_id=?', (id,)).fetchall()
    for item in items:
        inv = db.execute('SELECT * FROM inventory WHERE no=?', (item['code'],)).fetchone()
        if inv:
            new_remaining = inv['remaining'] + item['quantity']
            judgment = '' if new_remaining > inv['order_threshold'] else inv['order_judgment']
            db.execute('UPDATE inventory SET remaining=?, supply_qty=supply_qty+?, order_judgment=? WHERE no=?',
                       (new_remaining, item['quantity'], judgment, item['code']))
        else:
            # New item - insert into inventory
            max_no = db.execute('SELECT COALESCE(MAX(no),0) FROM inventory').fetchone()[0]
            new_no = item['code'] if item['code'] else max_no + 1
            db.execute('''INSERT OR IGNORE INTO inventory
                (product_name, no, stock, order_qty, remaining, order_threshold, note)
                VALUES (?,?,?,10,?,0,'新規追加')''',
                (item['product_name'], new_no, item['quantity'], item['quantity']))

    db.commit()
    db.close()
    flash('稟議書を承認し、在庫を更新しました', 'success')
    return redirect(url_for('ringi_detail', id=id))


@app.route('/ringi/<int:id>/reject', methods=['POST'])
def ringi_reject(id):
    db = get_db()
    db.execute('UPDATE ringi SET status=? WHERE id=?', ('rejected', id))
    db.commit()
    db.close()
    flash('稟議書を却下しました', 'warning')
    return redirect(url_for('ringi_detail', id=id))


@app.route('/ringi/<int:id>/delete', methods=['POST'])
def ringi_delete(id):
    db = get_db()
    db.execute('DELETE FROM ringi_items WHERE ringi_id=?', (id,))
    db.execute('DELETE FROM ringi WHERE id=?', (id,))
    db.commit()
    db.close()
    flash('稟議書を削除しました', 'success')
    return redirect(url_for('ringi_list'))


@app.route('/ringi/<int:id>/pdf')
def ringi_pdf(id):
    db = get_db()
    ringi = db.execute('SELECT * FROM ringi WHERE id=?', (id,)).fetchone()
    items = db.execute('SELECT * FROM ringi_items WHERE ringi_id=?', (id,)).fetchall()
    db.close()
    if not ringi:
        flash('稟議書が見つかりません', 'danger')
        return redirect(url_for('ringi_list'))
    return render_template('ringi_pdf.html', ringi=ringi, items=items)


@app.route('/ringi/<int:id>/csv')
def ringi_csv(id):
    db = get_db()
    items = db.execute('SELECT * FROM ringi_items WHERE ringi_id=?', (id,)).fetchall()
    db.close()

    output = io.StringIO()
    writer = csv.writer(output, quoting=csv.QUOTE_ALL)
    writer.writerow(['コード', '商品名', '購入先', '数量', '単価', '金額'])
    for item in items:
        writer.writerow([item['code'], item['product_name'], item['supplier'],
                         item['quantity'], item['unit_price'], item['amount']])

    response = make_response(output.getvalue())
    response.headers['Content-Type'] = 'text/csv; charset=utf-8-sig'
    response.headers['Content-Disposition'] = f'attachment; filename=ringi_{id}.csv'
    return response


# ============ Todo ============
@app.route('/todo')
def todo_list():
    db = get_db()
    filter_status = request.args.get('status', 'active')
    filter_priority = request.args.get('priority', 'all')
    filter_category = request.args.get('category', '')

    counts = {
        'all': db.execute('SELECT COUNT(*) FROM todos').fetchone()[0],
        'active': db.execute('SELECT COUNT(*) FROM todos WHERE done=0').fetchone()[0],
        'done': db.execute('SELECT COUNT(*) FROM todos WHERE done=1').fetchone()[0],
    }

    query = 'SELECT * FROM todos WHERE 1=1'
    params = []
    if filter_status == 'active':
        query += ' AND done=0'
    elif filter_status == 'done':
        query += ' AND done=1'
    if filter_priority != 'all':
        query += ' AND priority=?'
        params.append(filter_priority)
    if filter_category:
        query += ' AND category=?'
        params.append(filter_category)

    query += " ORDER BY done ASC, CASE priority WHEN 'high' THEN 0 WHEN 'medium' THEN 1 WHEN 'low' THEN 2 END, due_date IS NULL, due_date ASC, created_at DESC"
    todos = db.execute(query, params).fetchall()
    categories = [r[0] for r in db.execute("SELECT DISTINCT category FROM todos WHERE category != '' ORDER BY category").fetchall()]
    today = date.today().isoformat()
    db.close()
    return render_template('todo.html', todos=todos, counts=counts, categories=categories,
                           filter_status=filter_status, filter_priority=filter_priority,
                           filter_category=filter_category, today=today)


@app.route('/todo/add', methods=['POST'])
def todo_add():
    db = get_db()
    now = datetime.now().isoformat()
    db.execute('INSERT INTO todos (title, description, priority, category, due_date, created_at) VALUES (?,?,?,?,?,?)',
               (request.form['title'], request.form.get('description', ''),
                request.form.get('priority', 'medium'), request.form.get('category', ''),
                request.form.get('due_date') or None, now))
    db.commit()
    db.close()
    flash('タスクを追加しました', 'success')
    return redirect(url_for('todo_list'))


@app.route('/todo/toggle/<int:id>', methods=['POST'])
def todo_toggle(id):
    db = get_db()
    todo = db.execute('SELECT done FROM todos WHERE id=?', (id,)).fetchone()
    if todo:
        new_done = 0 if todo['done'] else 1
        completed = datetime.now().isoformat() if new_done else None
        db.execute('UPDATE todos SET done=?, completed_at=? WHERE id=?', (new_done, completed, id))
        db.commit()
    db.close()
    return redirect(request.referrer or url_for('todo_list'))


@app.route('/todo/edit/<int:id>', methods=['POST'])
def todo_edit(id):
    db = get_db()
    db.execute('UPDATE todos SET title=?, description=?, priority=?, category=?, due_date=? WHERE id=?',
               (request.form['title'], request.form.get('description', ''),
                request.form.get('priority', 'medium'), request.form.get('category', ''),
                request.form.get('due_date') or None, id))
    db.commit()
    db.close()
    flash('タスクを更新しました', 'success')
    return redirect(url_for('todo_list'))


@app.route('/todo/delete/<int:id>', methods=['POST'])
def todo_delete(id):
    db = get_db()
    db.execute('DELETE FROM todos WHERE id=?', (id,))
    db.commit()
    db.close()
    flash('タスクを削除しました', 'success')
    return redirect(url_for('todo_list'))


# ============ Search URL API ============
@app.route('/api/search_url')
def search_url():
    product_name = request.args.get('name', '')
    supplier = request.args.get('supplier', '')
    if supplier == 'アスクル':
        url = 'https://www.askul.co.jp/s/' + urllib.parse.quote(product_name)
    elif supplier == 'Amazon':
        url = 'https://www.amazon.co.jp/s?k=' + urllib.parse.quote(product_name)
    elif supplier == 'モノタロウ':
        url = 'https://www.monotaro.com/s/?c=&q=' + urllib.parse.quote(product_name)
    elif supplier == '楽天':
        url = 'https://search.rakuten.co.jp/search/mall/' + urllib.parse.quote(product_name)
    else:
        url = 'https://www.google.com/search?q=' + urllib.parse.quote(product_name + ' 購入')
    return jsonify({'url': url})


@app.route('/api/product/<int:code>')
def api_product(code):
    db = get_db()
    p = db.execute('SELECT * FROM products WHERE code=?', (code,)).fetchone()
    db.close()
    if p:
        return jsonify(dict(p))
    return jsonify({}), 404


if __name__ == '__main__':
    if not os.path.exists(DB_PATH):
        from init_db import init_db
        init_db()
    app.run(debug=True, host='0.0.0.0', port=5000)
