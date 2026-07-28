import os
import re
import random
import string
from datetime import datetime
import docx
from flask import Flask, render_template, request, redirect, url_for, session, jsonify, send_file

app = Flask(__name__, static_folder='.', static_url_path='')
app.secret_key = 'bm_exam_secure_secret_key_2026'

# Global configuration and state
CONFIG = {
    'INITIAL_TOKEN': 'MURNI2'  # Default login token
}

# Real-time state of student sessions (helps proctors monitor exams and manage locks)
ACTIVE_STUDENTS = {}

@app.after_request
def add_header(r):
    """Ensure browser never caches exam or login pages to prevent back-button bypasses."""
    r.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    r.headers["Pragma"] = "no-cache"
    r.headers["Expires"] = "0"
    return r

def clean_filename(name):
    """Clean student name to make it safe for file writing and session indexing."""
    cleaned = re.sub(r'[^a-zA-Z0-9_\-]', '_', name)
    return cleaned.strip('_').lower()

def parse_docx_math(el):
    """Recursively parse Office Math elements from DOCX XML to output clean HTML tags."""
    ns = {
        'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    }
    tag = el.tag.split('}')[-1]
    
    if tag == 't':
        return el.text or ''
    if tag in ('r', 'oMath'):
        return ''.join(parse_docx_math(c) for c in el)
    if tag == 'f':
        num = el.find('m:num', ns)
        den = el.find('m:den', ns)
        num_str = parse_docx_math(num) if num is not None else ''
        den_str = parse_docx_math(den) if den is not None else ''
        return f'<span class="math-fraction"><span class="math-num">{num_str}</span><span class="math-den">{den_str}</span></span>'
    if tag == 'sSup':
        base = el.find('m:e', ns)
        sup = el.find('m:sup', ns)
        base_str = parse_docx_math(base) if base is not None else ''
        sup_str = parse_docx_math(sup) if sup is not None else ''
        return f'{base_str}<sup>{sup_str}</sup>'
    if tag == 'd':
        e = el.find('m:e', ns)
        e_str = parse_docx_math(e) if e is not None else ''
        return f'({e_str})'
    
    return ''.join(parse_docx_math(c) for c in el)

def load_questions():
    """Load questions and options dynamically from DOCX file."""
    path = 'soal matematika/soal ulangan pilihan ganda.docx'
    if not os.path.exists(path):
        return []
    
    doc = docx.Document(path)
    paragraphs_text = []
    
    for p in doc.paragraphs:
        parts = []
        for child in p._element:
            parts.append(parse_docx_math(child))
        paragraphs_text.append(''.join(parts).strip())
        
    questions = []
    # Loop paragraphs in groups of 3 (Q, Options ACE, Options BD)
    for idx in range(0, len(paragraphs_text), 3):
        if idx + 2 >= len(paragraphs_text):
            break
        q_text = paragraphs_text[idx]
        opt_ace = paragraphs_text[idx+1]
        opt_bd = paragraphs_text[idx+2]
        
        # Regex/indexing extraction of choices
        a_text = ""
        c_text = ""
        e_text = ""
        
        c_idx = opt_ace.find('c.')
        e_idx = opt_ace.find('e.')
        
        if c_idx != -1 and e_idx != -1:
            a_text = opt_ace[:c_idx].strip()
            c_text = opt_ace[c_idx+2:].strip()
            e_idx_rel = c_text.find('e.')
            if e_idx_rel != -1:
                e_text = c_text[e_idx_rel+2:].strip()
                c_text = c_text[:e_idx_rel].strip()
        else:
            a_text = opt_ace
            
        b_text = ""
        d_text = ""
        d_idx = opt_bd.find('d.')
        if d_idx != -1:
            b_text = opt_bd[:d_idx].strip()
            d_text = opt_bd[d_idx+2:].strip()
        else:
            b_text = opt_bd
            
        def clean_opt(txt):
            txt = txt.strip()
            txt = re.sub(r'^[a-eA-E]\.\s*', '', txt)
            return txt

        choices = {
            'A': clean_opt(a_text),
            'B': clean_opt(b_text),
            'C': clean_opt(c_text),
            'D': clean_opt(d_text),
            'E': clean_opt(e_text)
        }
        
        questions.append({
            'index': len(questions) + 1,
            'question': q_text,
            'choices': choices
        })
        
    return questions

def load_answers():
    """Load answer keys dynamically from DOCX file."""
    path = 'soal matematika/Kunci jawaban.docx'
    if not os.path.exists(path):
        return {}
    
    doc = docx.Document(path)
    answers = {}
    q_index = 1
    
    for p in doc.paragraphs:
        txt = p.text.strip().upper()
        if not txt:
            continue
        if 'KUNCI' in txt or 'JAWABAN' in txt:
            continue
        match = re.search(r'[A-E]', txt)
        if match:
            answers[q_index] = match.group(0)
            q_index += 1
            
    return answers


@app.route('/')
def index():
    if 'siswa' in session:
        return redirect(url_for('ujian'))
    return redirect(url_for('login'))


@app.route('/login', methods=['GET', 'POST'])
def login():
    # If student is already in session but locked in backend, clear their session keys so they must re-authenticate
    if 'nama_clean' in session:
        nama_clean = session['nama_clean']
        if nama_clean in ACTIVE_STUDENTS and ACTIVE_STUDENTS[nama_clean]['status'] == 'terkunci':
            session.pop('siswa', None)
            session.pop('nama_clean', None)
            session.pop('status', None)

    if request.method == 'POST':
        nama = request.form.get('nama', '').strip()
        kelas = request.form.get('kelas', '').strip()
        jurusan = request.form.get('jurusan', '').strip()
        token = request.form.get('token', '').strip().upper()
        
        if not nama or not kelas or not jurusan or not token:
            return render_template('login.html', error='Semua kolom wajib diisi!')
            
        nama_clean = clean_filename(nama)
        
        # Check if student name already registered in active lists
        if nama_clean in ACTIVE_STUDENTS:
            student_state = ACTIVE_STUDENTS[nama_clean]
            
            if student_state['status'] == 'selesai':
                return render_template('login.html', error='Anda telah menyelesaikan ujian ini!')
                
            if student_state['status'] == 'terkunci':
                # Verify against student-specific token_baru
                expected_token = student_state['token_baru']
                if token == expected_token:
                    # Unlock student and log them back in
                    ACTIVE_STUDENTS[nama_clean]['status'] = 'aktif'
                    ACTIVE_STUDENTS[nama_clean]['token_baru'] = ''
                    ACTIVE_STUDENTS[nama_clean]['lock_time'] = None
                    
                    # Remove student token file
                    filepath = f"data_token/{nama_clean}.txt"
                    if os.path.exists(filepath):
                        try:
                            os.remove(filepath)
                        except Exception:
                            pass
                            
                    session['siswa'] = nama
                    session['kelas'] = kelas
                    session['jurusan'] = jurusan
                    session['nama_clean'] = nama_clean
                    session['status'] = 'aktif'
                    
                    return redirect(url_for('ujian'))
                else:
                    return render_template('login.html', error='Token masuk pengawas Anda salah atau tidak valid!')
            
            # If student is 'aktif' but session was lost/cleared (e.g. browser crash or remotely unlocked)
            if student_state['status'] == 'aktif':
                if token == CONFIG['INITIAL_TOKEN']:
                    session['siswa'] = nama
                    session['kelas'] = kelas
                    session['jurusan'] = jurusan
                    session['nama_clean'] = nama_clean
                    session['status'] = 'aktif'
                    return redirect(url_for('ujian'))
                else:
                    return render_template('login.html', error='Token masuk salah!')

        # Normal login with INITIAL_TOKEN (first-time entrance)
        if token != CONFIG['INITIAL_TOKEN']:
            return render_template('login.html', error='Token masuk salah!')
            
        # Get questions and shuffle the order for this specific student
        questions_all = load_questions()
        q_order = [q['index'] for q in questions_all]
        random.shuffle(q_order)
            
        # Register student in the global active list
        ACTIVE_STUDENTS[nama_clean] = {
            'nama': nama,
            'kelas': kelas,
            'jurusan': jurusan,
            'status': 'aktif',
            'token_baru': '',
            'lock_time': None,
            'score': None,
            'answers': {},
            'question_order': q_order
        }
        
        session['siswa'] = nama
        session['kelas'] = kelas
        session['jurusan'] = jurusan
        session['nama_clean'] = nama_clean
        session['status'] = 'aktif'
        
        return redirect(url_for('ujian'))
        
    return render_template('login.html')


@app.route('/ujian')
def ujian():
    if 'siswa' not in session:
        return redirect(url_for('login'))
        
    nama_clean = session['nama_clean']
    
    # Sync with global state
    if nama_clean in ACTIVE_STUDENTS:
        status = ACTIVE_STUDENTS[nama_clean]['status']
        session['status'] = status
    else:
        # Fallback if server restarted but cookie persists
        questions_all = load_questions()
        q_order = [q['index'] for q in questions_all]
        random.shuffle(q_order)
        
        ACTIVE_STUDENTS[nama_clean] = {
            'nama': session['siswa'],
            'kelas': session['kelas'],
            'jurusan': session['jurusan'],
            'status': session.get('status', 'aktif'),
            'token_baru': session.get('token_baru', ''),
            'lock_time': None,
            'score': None,
            'answers': {},
            'question_order': q_order
        }
        status = session.get('status', 'aktif')
        
    if status == 'terkunci':
        # Clear student session so they are logged out of the browser UI
        session.pop('siswa', None)
        session.pop('nama_clean', None)
        session.pop('status', None)
        return redirect(url_for('login', error='Anda telah keluar karena membuka tab pengerjaan. Silakan hubungi pengawas untuk mendapatkan token baru.'))
        
    if status == 'selesai':
        return redirect(url_for('selesai'))
        
    questions_all = load_questions()
    
    # Sort the questions list based on this student's recorded question_order
    student_order = ACTIVE_STUDENTS[nama_clean].get('question_order', [])
    if not student_order:
        student_order = [q['index'] for q in questions_all]
        random.shuffle(student_order)
        ACTIVE_STUDENTS[nama_clean]['question_order'] = student_order
        
    q_map = {q['index']: q for q in questions_all}
    questions = [q_map[idx] for idx in student_order if idx in q_map]
    
    return render_template('ujian.html', questions=questions, status=status)


@app.route('/lock', methods=['POST'])
def lock():
    if 'nama_clean' not in session:
        return jsonify({'error': 'Unauthorized'}), 401
        
    nama_clean = session['nama_clean']
    
    if nama_clean in ACTIVE_STUDENTS and ACTIVE_STUDENTS[nama_clean]['status'] != 'selesai':
        ACTIVE_STUDENTS[nama_clean]['status'] = 'terkunci'
        
        # Generate 6 character token
        token_baru = ''.join(random.choices(string.ascii_uppercase + string.digits, k=6))
        ACTIVE_STUDENTS[nama_clean]['token_baru'] = token_baru
        ACTIVE_STUDENTS[nama_clean]['lock_time'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        # Write file in data_token/
        os.makedirs('data_token', exist_ok=True)
        filename = f"data_token/{nama_clean}.txt"
        
        content = (
            "=============================\n"
            "TOKEN BARU SISWA\n"
            "=============================\n"
            f"Nama    : {ACTIVE_STUDENTS[nama_clean]['nama']}\n"
            f"Kelas   : {ACTIVE_STUDENTS[nama_clean]['kelas']}\n"
            f"Jurusan : {ACTIVE_STUDENTS[nama_clean]['jurusan']}\n"
            f"Token   : {token_baru}\n"
            f"Waktu   : {ACTIVE_STUDENTS[nama_clean]['lock_time']}\n"
        )
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(content)
            
        # Pop student credentials to log them out of Flask session
        session.pop('siswa', None)
        session.pop('nama_clean', None)
        session.pop('status', None)
            
        return jsonify({'status': 'locked', 'token_file': filename})
        
    return jsonify({'status': 'ignored'})


@app.route('/unlock', methods=['POST'])
def unlock():
    if 'siswa' not in session:
        return redirect(url_for('login'))
        
    nama_clean = session['nama_clean']
    token_input = request.form.get('token_input', '').strip().upper()
    
    expected_token = ''
    if nama_clean in ACTIVE_STUDENTS:
        expected_token = ACTIVE_STUDENTS[nama_clean]['token_baru']
        
    if token_input and token_input == expected_token:
        # Unlock success
        if nama_clean in ACTIVE_STUDENTS:
            ACTIVE_STUDENTS[nama_clean]['status'] = 'aktif'
            ACTIVE_STUDENTS[nama_clean]['token_baru'] = ''
            ACTIVE_STUDENTS[nama_clean]['lock_time'] = None
            
        session['status'] = 'aktif'
        session['token_baru'] = ''
        
        # Remove student token file
        filepath = f"data_token/{nama_clean}.txt"
        if os.path.exists(filepath):
            try:
                os.remove(filepath)
            except Exception:
                pass
                
        return redirect(url_for('ujian'))
    else:
        # Fail
        return render_template('ujian.html', 
                               questions=load_questions(), 
                               status='terkunci', 
                               error='Token salah! Silakan periksa kembali atau minta ulang pengawas.')


@app.route('/student/status')
def student_status():
    if 'nama_clean' not in session:
        return jsonify({'status': 'none'})
    nama_clean = session['nama_clean']
    if nama_clean in ACTIVE_STUDENTS:
        return jsonify({'status': ACTIVE_STUDENTS[nama_clean]['status']})
    return jsonify({'status': 'none'})


@app.route('/submit', methods=['POST'])
def submit():
    if 'siswa' not in session:
        return redirect(url_for('login'))
        
    nama_clean = session['nama_clean']
    
    # Check if student is locked in backend registry
    if nama_clean in ACTIVE_STUDENTS and ACTIVE_STUDENTS[nama_clean]['status'] == 'terkunci':
        session.pop('siswa', None)
        session.pop('nama_clean', None)
        session.pop('status', None)
        return redirect(url_for('login', error='Sesi Anda terkunci karena meninggalkan halaman ujian secara otomatis. Silakan minta kode token baru kepada pengawas kelas untuk masuk kembali.'))
        
    if nama_clean in ACTIVE_STUDENTS and ACTIVE_STUDENTS[nama_clean]['status'] == 'selesai':
        return redirect(url_for('selesai'))
        
    questions = load_questions()
    correct_answers = load_answers()
    
    student_answers = {}
    correct_count = 0
    total_count = len(questions)
    
    details = []
    
    for q in questions:
        q_idx = q['index']
        # Retrieve answer submitted by student
        ans = request.form.get(f'q{q_idx}', '').strip().upper()
        student_answers[q_idx] = ans
        
        expected = correct_answers.get(q_idx, '')
        is_correct = (ans == expected)
        if is_correct:
            correct_count += 1
            
        details.append({
            'index': q_idx,
            'question': q['question'],
            'student_answer': ans,
            'correct_answer': expected,
            'is_correct': is_correct,
            'choices': q['choices']
        })
        
    score = round((correct_count / total_count * 100), 2) if total_count > 0 else 0.0
    
    # Save statistics/answers in active student record
    if nama_clean in ACTIVE_STUDENTS:
        ACTIVE_STUDENTS[nama_clean]['status'] = 'selesai'
        ACTIVE_STUDENTS[nama_clean]['score'] = score
        ACTIVE_STUDENTS[nama_clean]['answers'] = student_answers
        
    session['status'] = 'selesai'
    session['score'] = score
    
    # Write to hasil ujian/...
    kelas = session['kelas']
    jurusan = session['jurusan']
    nama_siswa = session['siswa']
    
    # Folder path
    result_dir = os.path.join('hasil ujian', kelas, jurusan, nama_siswa)
    os.makedirs(result_dir, exist_ok=True)
    
    # Write results txt
    txt_path = os.path.join(result_dir, 'hasil.txt')
    with open(txt_path, 'w', encoding='utf-8') as f:
        f.write("==================================================\n")
        f.write("HASIL UJIAN SISWA - SMK BUDI MURNI 2\n")
        f.write("==================================================\n")
        f.write(f"Nama     : {nama_siswa}\n")
        f.write(f"Kelas    : {kelas}\n")
        f.write(f"Jurusan  : {jurusan}\n")
        f.write(f"Tanggal  : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"Skor     : {score} / 100\n")
        f.write(f"Benar    : {correct_count} dari {total_count} soal\n")
        f.write("--------------------------------------------------\n")
        f.write("DETAIL JAWABAN:\n")
        for d in details:
            status_symbol = "✓ BENAR" if d['is_correct'] else "✗ SALAH"
            f.write(f"Soal {d['index']}: {status_symbol}\n")
            f.write(f"  Jawaban Siswa: {d['student_answer'] or '-'}\n")
            f.write(f"  Jawaban Kunci: {d['correct_answer']}\n\n")
            
    # Write results html report
    html_path = os.path.join(result_dir, 'hasil_ujian.html')
    
    report_html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Laporan Hasil Ujian - {nama_siswa}</title>
    <style>
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 20px; background: #f4f6f9; color: #333; }}
        .report-card {{ background: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); max-width: 800px; margin: auto; }}
        .header {{ text-align: center; border-bottom: 2px solid #007bff; padding-bottom: 20px; margin-bottom: 30px; }}
        .school-title {{ font-size: 24px; font-weight: bold; color: #0c2340; text-transform: uppercase; margin: 5px 0; }}
        .report-title {{ font-size: 18px; color: #007bff; margin: 5px 0; }}
        .meta-table {{ width: 100%; border-collapse: collapse; margin-bottom: 30px; }}
        .meta-table td {{ padding: 8px 12px; border: 1px dashed #ddd; font-size: 15px; }}
        .meta-label {{ font-weight: bold; width: 150px; background: #f8f9fa; }}
        .score-box {{ text-align: center; margin: 30px 0; padding: 20px; background: #e8f0fe; border-radius: 8px; border: 1px solid #b3d1ff; }}
        .score-val {{ font-size: 48px; font-weight: bold; color: #1a73e8; }}
        .detail-item {{ background: #fafafa; border: 1px solid #eee; border-radius: 8px; padding: 15px; margin-bottom: 15px; }}
        .detail-item.correct {{ border-left: 5px solid #28a745; }}
        .detail-item.incorrect {{ border-left: 5px solid #dc3545; }}
        .question-txt {{ font-weight: 500; font-size: 16px; margin-bottom: 10px; }}
        .answer-comparison {{ display: flex; gap: 20px; font-size: 14px; margin-top: 10px; }}
        .ans-badge {{ padding: 4px 10px; border-radius: 4px; font-weight: 500; }}
        .badge-student {{ background: #e9ecef; border: 1px solid #ced4da; }}
        .badge-correct {{ background: #d4edda; border: 1px solid #c3e6cb; color: #155724; }}
        .status-tag {{ font-weight: bold; float: right; font-size: 14px; text-transform: uppercase; }}
        .status-correct {{ color: #28a745; }}
        .status-incorrect {{ color: #dc3545; }}
        /* Math formulas rendering */
        .math-fraction {{ display: inline-block; vertical-align: middle; text-align: center; padding: 0 4px; }}
        .math-num {{ display: block; border-bottom: 1px solid #333; padding: 0 2px; }}
        .math-den {{ display: block; padding: 0 2px; }}
        @media print {{
            body {{ background: white; padding: 0; }}
            .report-card {{ box-shadow: none; max-width: 100%; padding: 0; }}
        }}
    </style>
</head>
<body>
    <div class="report-card">
        <div class="header">
            <div class="school-title">SMK Budi Murni 2</div>
            <div class="report-title">Laporan Lembar Hasil Ujian Matematika</div>
        </div>
        
        <table class="meta-table">
            <tr>
                <td class="meta-label">Nama Siswa</td>
                <td>{nama_siswa}</td>
                <td class="meta-label">Kelas</td>
                <td>{kelas}</td>
            </tr>
            <tr>
                <td class="meta-label">Jurusan</td>
                <td>{jurusan}</td>
                <td class="meta-label">Waktu Selesai</td>
                <td>{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</td>
            </tr>
        </table>
        
        <div class="score-box">
            <div>NILAI AKHIR:</div>
            <div class="score-val">{score}</div>
            <div style="font-size:14px; color:#555;">Keterangan: Benar {correct_count} dari {total_count} soal</div>
        </div>
        
        <h3>Analisis Lembar Jawaban:</h3>
"""
    
    for d in details:
        status_cls = "correct" if d['is_correct'] else "incorrect"
        status_lbl = "BENAR" if d['is_correct'] else "SALAH"
        status_lbl_cls = "status-correct" if d['is_correct'] else "status-incorrect"
        
        student_ans_val = d['student_answer']
        student_ans_text = f"{student_ans_val}. {d['choices'].get(student_ans_val, '')}" if student_ans_val else "-"
        correct_ans_text = f"{d['correct_answer']}. {d['choices'].get(d['correct_answer'], '')}"
        
        report_html += f"""
        <div class="detail-item {status_cls}">
            <span class="status-tag {status_lbl_cls}">{status_lbl}</span>
            <div class="question-txt">Soal {d['index']}: {d['question']}</div>
            <div class="answer-comparison">
                <div>Jawaban Siswa: <span class="ans-badge badge-student">{student_ans_text}</span></div>
                <div>Jawaban Kunci: <span class="ans-badge badge-correct">{correct_ans_text}</span></div>
            </div>
        </div>
        """
        
    report_html += """
    </div>
</body>
</html>
"""
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(report_html)
        
    return redirect(url_for('selesai'))


@app.route('/selesai')
def selesai():
    if 'siswa' not in session:
        return redirect(url_for('login'))
        
    score = session.get('score', 0.0)
    questions = load_questions()
    correct_answers = load_answers()
    
    nama_clean = session['nama_clean']
    student_ans = {}
    if nama_clean in ACTIVE_STUDENTS:
        student_ans = ACTIVE_STUDENTS[nama_clean]['answers']
        
    details = []
    for q in questions:
        idx = q['index']
        ans = student_ans.get(idx, '')
        expected = correct_answers.get(idx, '')
        details.append({
            'index': idx,
            'question': q['question'],
            'student_answer': ans,
            'correct_answer': expected,
            'is_correct': (ans == expected)
        })
        
    return render_template('selesai.html', score=score, details=details)


# PROCTOR MODULE (SUPERVISOR LOGINS & STATE CONTROLS)

@app.route('/pengawas/login', methods=['GET', 'POST'])
def proctor_login():
    if 'proctor' in session:
        return redirect(url_for('proctor_dashboard'))
        
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()
        
        if username == 'pengawas1' and password == 'budimurnijaya':
            session['proctor'] = 'active'
            return redirect(url_for('proctor_dashboard'))
        else:
            return render_template('pengawas.html', is_login=True, error='Log masuk salah!')
            
    return render_template('pengawas.html', is_login=True)


@app.route('/pengawas')
def proctor_dashboard():
    if 'proctor' not in session:
        return redirect(url_for('proctor_login'))
        
    return render_template('pengawas.html', is_login=False, 
                           initial_token=CONFIG['INITIAL_TOKEN'], 
                           students=list(ACTIVE_STUDENTS.values()))


@app.route('/pengawas/api/students')
def proctor_api_students():
    if 'proctor' not in session:
        return jsonify({'error': 'Unauthorized'}), 401
        
    # Return active student list as JSON (allows refreshing the proctor page dynamically)
    return jsonify(list(ACTIVE_STUDENTS.values()))


@app.route('/pengawas/api/unlock/<nama>', methods=['POST'])
def proctor_api_unlock(nama):
    if 'proctor' not in session:
        return jsonify({'error': 'Unauthorized'}), 401
    
    nama_clean = clean_filename(nama)
    if nama_clean in ACTIVE_STUDENTS:
        ACTIVE_STUDENTS[nama_clean]['status'] = 'aktif'
        ACTIVE_STUDENTS[nama_clean]['token_baru'] = ''
        ACTIVE_STUDENTS[nama_clean]['lock_time'] = None
        
        # Remove file
        filepath = f"data_token/{nama_clean}.txt"
        if os.path.exists(filepath):
            try:
                os.remove(filepath)
            except Exception:
                pass
        return jsonify({'status': 'success'})
        
    return jsonify({'error': 'Student not found'}), 404


@app.route('/pengawas/api/reset/<nama>', methods=['POST'])
def proctor_api_reset(nama):
    if 'proctor' not in session:
        return jsonify({'error': 'Unauthorized'}), 401
        
    nama_clean = clean_filename(nama)
    if nama_clean in ACTIVE_STUDENTS:
        # Delete token file if exists
        filepath = f"data_token/{nama_clean}.txt"
        if os.path.exists(filepath):
            try:
                os.remove(filepath)
            except Exception:
                pass
                
        # Completely remove/reset student records so they can log back in
        del ACTIVE_STUDENTS[nama_clean]
        return jsonify({'status': 'success'})
        
    return jsonify({'error': 'Student not found'}), 404


@app.route('/pengawas/api/update_token', methods=['POST'])
def proctor_api_update_token():
    if 'proctor' not in session:
        return jsonify({'error': 'Unauthorized'}), 401
        
    new_token = request.form.get('new_token', '').strip().upper()
    if not new_token:
        return jsonify({'error': 'Token invalid'}), 400
        
    CONFIG['INITIAL_TOKEN'] = new_token
    return jsonify({'status': 'success', 'new_token': new_token})


@app.route('/pengawas/logout')
def proctor_logout():
    session.pop('proctor', None)
    return redirect(url_for('proctor_login'))


if __name__ == '__main__':
    # Start on all interfaces (host='0.0.0.0') on port 5000
    app.run(host='0.0.0.0', port=5000, debug=True)
