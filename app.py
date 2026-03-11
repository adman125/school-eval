import os
import sqlite3
import datetime
import csv
from io import StringIO
from flask import Flask, request, jsonify, send_from_directory, make_response, Response

# 路径配置
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, 'static')
DATA_DIR = os.path.join(BASE_DIR, 'data')
DB_PATH = os.path.join(DATA_DIR, 'eval.db')

app = Flask(__name__, static_folder=STATIC_DIR, static_url_path='/static')

# CORS 跨域处理
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

# 数据库连接
def get_db():
    os.makedirs(DATA_DIR, exist_ok=True)
    # 增加 timeout 防止多用户同时写入时数据库锁定
    conn = sqlite3.connect(DB_PATH, timeout=20)
    conn.row_factory = sqlite3.Row
    return conn

# 初始化数据库
def init_db():
    conn = get_db()
    # 修复了原代码中 SQL 内部包含 # 注释导致的语法错误
    conn.execute('''
        CREATE TABLE IF NOT EXISTS submissions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            evaluator TEXT NOT NULL,
            dept TEXT,
            position TEXT NOT NULL,
            customer_role TEXT,
            indicator TEXT NOT NULL,
            score REAL,
            comment TEXT,
            submitted_at TEXT NOT NULL,
            UNIQUE(evaluator, position, indicator)
        )
    ''')
    # 增加索引提升查询性能
    conn.execute("CREATE INDEX IF NOT EXISTS idx_eval_pos ON submissions(evaluator, position)")
    conn.commit()
    conn.close()

init_db()

# --- 路由部分 ---

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
                DO UPDATE SET 
                    score=excluded.score, 
                    comment=excluded.comment, 
                    submitted_at=excluded.submitted_at
            ''', (evaluator, r.get('dept',''), r.get('position',''), r.get('role',''),
                  r.get('indicator',''), r.get('score') if r.get('score') is not None else None, 
                  r.get('comment',''), now))
        conn.commit()
        return jsonify({'ok': True, 'saved': len(rows)})
    except Exception as e:
        return jsonify({'ok': False, 'msg': str(e)}), 500
    finally:
        conn.close()

@app.route('/api/stats')
def stats():
    conn = get_db()
    try:
        submitted = conn.execute("SELECT COUNT(DISTINCT evaluator) as n FROM submissions").fetchone()['n']
        total_rows = conn.execute("SELECT COUNT(*) as n FROM submissions WHERE score IS NOT NULL").fetchone()['n']
        avg_res = conn.execute("SELECT AVG(score) as a FROM submissions WHERE score IS NOT NULL").fetchone()['a']
        return jsonify({
            'submitted': submitted, 
            'total_rows': total_rows,
            'avg': round(avg_res, 2) if avg_res else None, 
            'total_evaluators': 469
        })
    finally:
        conn.close()

@app.route('/api/who_submitted')
def who_submitted():
    conn = get_db()
    rows = conn.execute("""
        SELECT DISTINCT evaluator, MAX(submitted_at) as last_at 
        FROM submissions 
        GROUP BY evaluator 
        ORDER BY last_at DESC
    """).fetchall()
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
    if score: sql += " AND score=?";           params.append(score)
    if dept:  sql += " AND dept LIKE ?";       params.append(f'%{dept}%')
    
    sql += " ORDER BY submitted_at DESC LIMIT ? OFFSET ?"
    params.extend([limit, offset])
    
    conn = get_db()
    rows = conn.execute(sql, params).fetchall()
    
    # 获取总数用于分页
    csql = "SELECT COUNT(*) as n FROM submissions WHERE 1=1"
    cp = []
    if ev:    csql += " AND evaluator LIKE ?";  cp.append(f'%{ev}%')
    if pos:   csql += " AND position LIKE ?";   cp.append(f'%{pos}%')
    if score: csql += " AND score=?";           cp.append(score)
    if dept:  csql += " AND dept LIKE ?";       cp.append(f'%{dept}%')
    total = conn.execute(csql, cp).fetchone()['n']
    conn.close()
    
    return jsonify({'total': total, 'rows': [dict(r) for r in rows]})

@app.route('/api/by_position')
def by_position():
    dept = request.args.get('dept', '')
    pos  = request.args.get('position', '')
    sql = """
        SELECT dept, position, customer_role, indicator, COUNT(*) as cnt, AVG(score) as avg_score 
        FROM submissions 
        WHERE score IS NOT NULL
    """
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

    # 使用 csv 模块处理，自动解决评论中包含引号、换行、制表符导致的格式乱行问题
    output = StringIO()
    # 写入 UTF-8 BOM 头，确保 Excel 打开中文不乱码
    output.write('\ufeff')
    writer = csv.writer(output, delimiter='\t')
    
    writer.writerow(['评价人', '部门', '被评岗位', '评价者角色', '评价指标', '评分', '评语', '提交时间'])
    for r in rows:
        writer.writerow([
            r['evaluator'], 
            r['dept'] or '', 
            r['position'] or '',
            r['customer_role'] or '', 
            (r['indicator'] or '').replace('\n', ' '),
            r['score'] if r['score'] is not None else '', 
            (r['comment'] or '').replace('\n', ' '), 
            r['submitted_at']
        ])
    
    response = Response(output.getvalue(), mimetype='text/tab-separated-values')
    response.headers['Content-Disposition'] = 'attachment; filename=eval_export.xls'
    return response

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    # 生产环境建议将 debug 设置为 False
    app.run(host='0.0.0.0', port=port, debug=True)
