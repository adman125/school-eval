from flask import Flask, request, jsonify, send_from_directory, make_response, Response
import os, sqlite3, datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, 'static')
DATA_DIR = os.path.join(BASE_DIR, 'data')
DB_PATH = os.path.join(DATA_DIR, 'eval.db')

app = Flask(__name__, static_folder=STATIC_DIR, static_url_path='/static')

@app.after_request
def add_cors(response):
    response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type'
    response.headers['Access-Control-Allow-Methods'] = 'GET,POST,OPTIONS'
    return response

@app.before_request
def handle_options():
    if request.method == 'OPTIONS':
        return make_response('', 204)

def get_db():
    os.makedirs(DATA_DIR, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    conn.execute('''
        CREATE TABLE IF NOT EXISTS submissions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            evaluator TEXT NOT NULL,
            dept TEXT,
            position TEXT,
            customer_role TEXT,
            indicator TEXT,
            score INTEGER,
            comment TEXT,
            submitted_at TEXT,
            UNIQUE(evaluator, position, indicator)
        )
    ''')
    conn.commit()
    conn.close()

init_db()

@app.route('/')
def index():
    return send_from_directory(STATIC_DIR, 'index.html')

@app.route('/admin')
def admin():
    return send_from_directory(STATIC_DIR, 'admin.html')

@app.route('/api/submit', methods=['POST'])
def submit():
    data = request.json
    evaluator = data.get('evaluator', '').strip()
    rows = data.get('rows', [])
    if not evaluator or not rows:
        return jsonify({'ok': False, 'msg': '数据为空'}), 400
    now = datetime.datetime.now().isoformat()
    conn = get_db()
    try:
        for r in rows:
            conn.execute('''
                INSERT INTO submissions
                    (evaluator, dept, position, customer_role, indicator, score, comment, submitted_at)
                VALUES (?,?,?,?,?,?,?,?)
                ON CONFLICT(evaluator, position, indicator)
                DO UPDATE SET score=excluded.score, comment=excluded.comment, submitted_at=excluded.submitted_at
            ''', (evaluator, r.get('dept',''), r.get('position',''), r.get('role',''),
                  r.get('indicator',''), r.get('score') or None, r.get('comment',''), now))
        conn.commit()
        return jsonify({'ok': True, 'saved': len(rows)})
    except Exception as e:
        return jsonify({'ok': False, 'msg': str(e)}), 500
    finally:
        conn.close()

@app.route('/api/stats')
def stats():
    conn = get_db()
    submitted = conn.execute("SELECT COUNT(DISTINCT evaluator) as n FROM submissions").fetchone()['n']
    total_rows = conn.execute("SELECT COUNT(*) as n FROM submissions WHERE score IS NOT NULL").fetchone()['n']
    avg = conn.execute("SELECT AVG(score) as a FROM submissions WHERE score IS NOT NULL").fetchone()['a']
    conn.close()
    return jsonify({'submitted': submitted, 'total_rows': total_rows,
                    'avg': round(avg, 2) if avg else None, 'total_evaluators': 469})

@app.route('/api/who_submitted')
def who_submitted():
    conn = get_db()
    rows = conn.execute("SELECT DISTINCT evaluator, MAX(submitted_at) as last_at FROM submissions GROUP BY evaluator ORDER BY last_at DESC").fetchall()
    conn.close()
    return jsonify([{'name': r['evaluator'], 'at': r['last_at']} for r in rows])

@app.route('/api/detail')
def detail():
    ev = request.args.get('evaluator', '')
    pos = request.args.get('position', '')
    score = request.args.get('score', '')
    dept = request.args.get('dept', '')
    limit = int(request.args.get('limit', 100))
    offset = int(request.args.get('offset', 0))
    sql = "SELECT * FROM submissions WHERE 1=1"
    params = []
    if ev:    sql += " AND evaluator LIKE ?";  params.append(f'%{ev}%')
    if pos:   sql += " AND position LIKE ?";   params.append(f'%{pos}%')
    if score: sql += " AND score=?";           params.append(int(score))
    if dept:  sql += " AND dept LIKE ?";       params.append(f'%{dept}%')
    sql += " ORDER BY submitted_at DESC LIMIT ? OFFSET ?"
    params += [limit, offset]
    conn = get_db()
    rows = conn.execute(sql, params).fetchall()
    csql = "SELECT COUNT(*) as n FROM submissions WHERE 1=1"
    cp = []
    if ev:    csql += " AND evaluator LIKE ?";  cp.append(f'%{ev}%')
    if pos:   csql += " AND position LIKE ?";   cp.append(f'%{pos}%')
    if score: csql += " AND score=?";           cp.append(int(score))
    if dept:  csql += " AND dept LIKE ?";       cp.append(f'%{dept}%')
    total = conn.execute(csql, cp).fetchone()['n']
    conn.close()
    return jsonify({'total': total, 'rows': [dict(r) for r in rows]})

@app.route('/api/by_position')
def by_position():
    dept = request.args.get('dept', '')
    pos  = request.args.get('position', '')
    sql = "SELECT dept, position, customer_role, indicator, COUNT(*) as cnt, AVG(score) as avg_score FROM submissions WHERE score IS NOT NULL"
    params = []
    if dept: sql += " AND dept LIKE ?";     params.append(f'%{dept}%')
    if pos:  sql += " AND position LIKE ?"; params.append(f'%{pos}%')
    sql += " GROUP BY dept, position, customer_role, indicator ORDER BY dept, position, avg_score"
    conn = get_db()
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])

@app.route('/api/export_csv')
def export_csv():
    conn = get_db()
    rows = conn.execute("SELECT * FROM submissions ORDER BY submitted_at DESC").fetchall()
    conn.close()
    lines = ['评价人\t部门\t被评岗位\t评价者角色\t评价指标\t评分\t评语\t提交时间']
    for r in rows:
        lines.append('\t'.join([
            r['evaluator'], r['dept'] or '', r['position'] or '',
            r['customer_role'] or '', (r['indicator'] or '').replace('\n',' ')[:80],
            str(r['score'] or ''), r['comment'] or '', r['submitted_at'] or ''
        ]))
    return Response('\ufeff' + '\n'.join(lines),
        mimetype='text/tab-separated-values',
        headers={'Content-Disposition': 'attachment; filename=eval_export.txt'})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
