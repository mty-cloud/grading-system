#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
作业批改系统 - Web 应用
基于智谱AI视觉模型的实验报告智能批改系统
"""
import os, re, json, base64, sys, webbrowser, threading, time
from datetime import datetime

def resource_path(rel):
    try:
        return os.path.join(sys._MEIPASS, rel)
    except:
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), rel)

# API Key
ZHIPU_API_KEY = "ae0c33227c2542659ae03872244959cb.hcjuR7slRxiPPROU"
ZHIPU_MODEL = "glm-4v-flash"
_env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env')
if os.path.exists(_env_path):
    try:
        with open(_env_path) as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#') and '=' in line:
                    k, v = line.split('=', 1)
                    if k.strip() == 'ZHIPU_API_KEY': ZHIPU_API_KEY = v.strip()
                    elif k.strip() == 'ZHIPU_MODEL': ZHIPU_MODEL = v.strip()
    except: pass

import requests
from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory, jsonify
import docx
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

def _exe_dir():
    try:
        return os.path.dirname(os.path.abspath(sys.executable))
    except:
        return os.path.dirname(os.path.abspath(__file__))

app = Flask(__name__, template_folder=resource_path('templates'), static_folder=resource_path('static'))
app.secret_key = "grading_secret_key_2024"
BASE_DIR = _exe_dir()
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
IMAGE_FOLDER = os.path.join(BASE_DIR, "images")
REPORTS_FOLDER = os.path.join(BASE_DIR, "reports")
REPORT_FILE = os.path.join(BASE_DIR, "批改报告.html")
ZHIPU_BASE_URL = "https://open.bigmodel.cn/api/paas/v4/"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(IMAGE_FOLDER, exist_ok=True)
os.makedirs(REPORTS_FOLDER, exist_ok=True)

# ---- Prompt 构建 ----
def build_grading_prompt(text):
    excerpt = text[:3000] if text else "(无)"
    return f"""你是严格的实验报告评分助手。请根据以下实验报告的文字内容和截图，给出0-100分的评分。

## 文字内容
```
{excerpt}
```

## 要求
1. 根据你的判断客观评分，不要刻意打高分
2. 文字质量（实验目的、步骤、总结是否完整，有无深度）和截图质量都要考虑
3. 差的报告给低分（0-40），一般的给中等分（40-70），优秀的给高分（70-100）
4. 截图数量和质量是重要判断依据
5. 如果文字明显敷衍、大量重复、内容过少，必须给低分

## 输出JSON格式
{{"image_score":0-50,"text_score":0-30,"tech_score":0-20,"total_score":0-100,"analysis":"一句话评价","text_feedback":"文字评价","image_feedback":"截图评价"}}"""

def build_text_only_prompt(text):
    excerpt = text[:4000] if text else "(无)"
    return f"""你是严格的实验报告评分助手。以下实验报告完全没有截图，请仅根据文字内容评分。

## 文字内容
```
{excerpt}
```

## 要求
1. 根据文字内容的质量独立打分，不要因为没截图就心软
2. 缺截图是严重缺陷，最高75分
3. 内容完整充实可以给较高分，内容空洞敷衍必须给低分

## 输出JSON格式
{{"content_score":0-30,"tech_score":0-25,"summary_score":0-25,"format_score":0-20,"total_score":0-100,"analysis":"一句话评价","text_feedback":"具体评价"}}"""

# ---- 辅助函数 ----
def is_zhipu_available():
    return bool(ZHIPU_API_KEY)

def get_available_vision_models():
    return [ZHIPU_MODEL] if ZHIPU_API_KEY else []

def image_b64(path):
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")

# ---- AI分析截图 ----
def analyze_image_with_zhipu(image_path, text):
    if not ZHIPU_API_KEY:
        return {"error": "未配置API Key"}
    try:
        b64 = image_b64(image_path)
        mime = "image/png"
        if image_path.lower().endswith(('.jpg','.jpeg')): mime = "image/jpeg"
        elif image_path.lower().endswith('.gif'): mime = "image/gif"
        elif image_path.lower().endswith('.webp'): mime = "image/webp"
        payload = {
            "model": ZHIPU_MODEL,
            "messages": [{"role": "user", "content": [
                {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{b64}"}},
                {"type": "text", "text": build_grading_prompt(text)}
            ]}],
            "temperature": 0.1, "max_tokens": 1024
        }
        headers = {"Authorization": f"Bearer {ZHIPU_API_KEY}", "Content-Type": "application/json"}
        resp = requests.post(f"{ZHIPU_BASE_URL}chat/completions", headers=headers, json=payload, timeout=120)
        if resp.status_code != 200:
            return {"error": f"API错误({resp.status_code})"}
        content = resp.json().get("choices", [{}])[0].get("message", {}).get("content", "")
        if not content:
            return {"error": "API返回为空"}
        cleaned = re.sub(r'^```(?:json)?\s*|\s*```$', '', content.strip(), flags=re.MULTILINE)
        m = re.search(r'\{.*"total_score".*\}', cleaned, re.DOTALL)
        if m:
            return json.loads(m.group())
        return {"error": f"无法解析输出: {content[:200]}"}
    except requests.Timeout:
        return {"error": "请求超时"}
    except Exception as e:
        return {"error": str(e)}

# ---- AI纯文字评分 ----
def ai_grade_text_only(text):
    if not ZHIPU_API_KEY:
        return auto_grade_fallback(text, [])
    try:
        p = build_text_only_prompt(text)
        payload = {"model": ZHIPU_MODEL, "messages": [{"role": "user", "content": [{"type": "text", "text": p}]}], "temperature": 0.1, "max_tokens": 1024}
        headers = {"Authorization": f"Bearer {ZHIPU_API_KEY}", "Content-Type": "application/json"}
        print("  📝 正在AI文字评分（无截图）...")
        resp = requests.post(f"{ZHIPU_BASE_URL}chat/completions", headers=headers, json=payload, timeout=60)
        if resp.status_code != 200:
            return auto_grade_fallback(text, [])
        content = resp.json().get("choices", [{}])[0].get("message", {}).get("content", "")
        if not content:
            return auto_grade_fallback(text, [])
        cleaned = re.sub(r'^```(?:json)?\s*|\s*```$', '', content.strip(), flags=re.MULTILINE)
        m = re.search(r'\{.*"total_score".*\}', cleaned, re.DOTALL)
        if m:
            data = json.loads(m.group())
            score = data.get("total_score", 50)
            analysis = data.get("analysis", "")
            feedback = data.get("text_feedback", "")
            comments = []
            if score >= 90: comments.append("🎉 文字内容质量优秀！")
            elif score >= 75: comments.append("✅ 文字内容良好。")
            elif score >= 60: comments.append("⚠️ 文字内容基本合格。")
            elif score >= 40: comments.append("⚠️ 文字内容不够充分。")
            else: comments.append("❌ 文字内容严重不足。")
            comments.append("⚠️ 该报告未提交任何实验截图，已根据评分规则扣分")
            if analysis: comments.append(f"📝 {analysis}")
            if feedback: comments.append(f"📝 评价: {feedback}")
            print(f"  ✅ AI文字评分完成: {score}分")
            return score, "；".join(comments)
        return auto_grade_fallback(text, [])
    except Exception as e:
        print(f"  ⚠️ AI文字评分异常: {e}")
        return auto_grade_fallback(text, [])

# ---- AI综合评分（图片+文字） ----
def ai_grade_with_zhipu(text, images):
    if not images:
        print("  ⚠️ 该文档没有截图，使用AI文字评分...")
        return ai_grade_text_only(text)

    all_findings = []
    total_scores = []
    text_feedbacks = []
    image_feedbacks = []

    for idx, img in enumerate(images):
        print(f"  📤 正在综合评分 截图{idx+1}/{len(images)}: {img['filename']}")
        result = analyze_image_with_zhipu(img["path"], text)
        if "error" in result:
            print(f"  ⚠️ 截图{idx+1} 分析失败: {result['error']}")
            continue
        score = result.get("total_score", 60)
        total_scores.append(score)
        text_fb = str(result.get("text_feedback", "") or "")
        image_fb = str(result.get("image_feedback", "") or "")
        if text_fb: text_feedbacks.append(text_fb)
        if image_fb: image_feedbacks.append(image_fb)
        all_findings.append({
            "index": idx+1, "filename": img["filename"], "score": score,
            "image_score": result.get("image_score", 0),
            "text_score": result.get("text_score", 0),
            "tech_score": result.get("tech_score", 0),
            "analysis": result.get("analysis", ""),
            "text_feedback": text_fb, "image_feedback": image_fb,
        })
        print(f"  ✅ 截图{idx+1}: {score}分")

    if not total_scores:
        print("  ⚠️ 所有截图分析失败，回退到文字评分")
        return ai_grade_text_only(text)

    avg = sum(total_scores) / len(total_scores)
    final_score = round(max(0, min(100, avg)))

    ts_avg = sum(f.get("text_score", 0) for f in all_findings) / len(all_findings)
    is_avg = sum(f.get("image_score", 0) for f in all_findings) / len(all_findings)
    tc_avg = sum(f.get("tech_score", 0) for f in all_findings) / len(all_findings)

    comments = []
    if final_score >= 90: comments.append("🎉 实验完成度极高！")
    elif final_score >= 75: comments.append("✅ 实验基本完成。")
    elif final_score >= 60: comments.append("⚠️ 实验部分完成，有改进空间。")
    else: comments.append("❌ 实验存在明显问题，需要完善。")
    comments.append(f"截图 {is_avg:.0f}/50 + 内容 {ts_avg:.0f}/30 + 技术 {tc_avg:.0f}/20 = {final_score} 分")

    seen = set()
    for fb in text_feedbacks:
        if fb not in seen:
            seen.add(fb)
            if len(seen) <= 2: comments.append(f"📝 文字: {fb}")
    seen = set()
    for fb in image_feedbacks:
        if fb not in seen:
            seen.add(fb)
            if len(seen) <= 2: comments.append(f"🖼️ 截图: {fb}")
    for f in all_findings:
        a = f.get("analysis", "")
        if a: comments.append(f"截图{f['index']}: {a}")

    return final_score, "；".join(comments)

# ---- 关键词回退 ----
def auto_grade_fallback(text, images):
    has_img = len(images) > 0
    score = 50 if has_img else 20
    has_summary = bool(re.search(r'实验总结|实验心得|实验小结', text))
    has_problems = bool(re.search(r'遇到|问题|错误|报错|故障', text))
    has_code = bool(re.search(r'start-dfs|start-yarn|jps|hadoop|hive', text))
    has_steps = bool(re.search(r'实验步骤|操作步骤|实验原理', text))
    wc = len(text)
    if has_summary: score += 10
    if has_problems: score += 10
    if has_code: score += 5
    if has_steps: score += 5
    if len(images) >= 3: score += 10
    elif len(images) >= 1: score += 5
    if wc > 2000: score += 5
    if wc < 500: score -= 15
    if not has_img: score = min(score, 50)
    score = max(0, min(100, score))
    c = []
    if score >= 90: c.append("[关键词] 实验完成度极高")
    elif score >= 75: c.append("[关键词] 实验基本完成")
    elif score >= 60: c.append("[关键词] 实验部分完成")
    elif score >= 40: c.append("[关键词] 实验内容不足")
    else: c.append("[关键词] 实验严重不足")
    if not has_img: c.append("❌ 缺少实验截图！")
    return score, "；".join(c)

# ---- 文档解析 ----
def extract_student_info(doc, filename=""):
    text = ""
    for p in doc.paragraphs: text += p.text + "\n"
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells: text += c.text.strip() + " "
    info = {"student_id": "", "name": "", "class_name": "", "course": "", "experiment": ""}
    for p in [r'学\s*号[：:\s/]*_*(\d{10,12})', r'学号/班级[：:\s]*_*(\d{10,12})', r'学号[：:\s]*(\d{10,12})']:
        m = re.search(p, text)
        if m: info["student_id"] = m.group(1); break
    if not info["student_id"] and filename:
        m = re.search(r'(\d{10,12})', filename)
        if m: info["student_id"] = m.group(1)
    for p in [r'姓\s*名[：:\s]*_*([一-龥]{2,4})', r'实验人[：:\s]*_*([一-龥]{2,4})', r'姓名[：:\s]*([一-龥]{2,4})']:
        m = re.search(p, text)
        if m: info["name"] = m.group(1); break
    if not info["name"] and filename:
        m = re.search(r'\d{10,12}([一-龥]{2,4})', filename)
        if m: info["name"] = m.group(1)
    for p in [r'学号/班级[：:\s]*_*\d{10,12}_*[/_]*([^\s_]{3,10})', r'班\s*级[：:\s]*([^\s]{2,10})']:
        m = re.search(p, text)
        if m: info["class_name"] = m.group(1); break
    for p in [r'课程名称[：:\s]+([^\s]{2,20})', r'课程[：:\s]*([^\s]{2,20})']:
        m = re.search(p, text)
        if m: info["course"] = m.group(1); break
    first_para = ""
    for p in doc.paragraphs:
        if p.text.strip(): first_para = p.text.strip(); break
    for p in [r'实验[一二三四五六七八九十]+[：:]\s*(.+?)(?:\n|$)', r'实验名称[：:\s]+([^\s]{2,30})']:
        m = re.search(p, text)
        if m: info["experiment"] = m.group(1); break
    if not info["experiment"] and first_para and len(first_para) > 4:
        info["experiment"] = first_para
    return info, text

def extract_images(doc, doc_id):
    images = []
    for rid, rel in doc.part.rels.items():
        if "image" in str(rel.reltype):
            data = rel.target_part.blob
            ext = os.path.splitext(rel.target_ref)[1] or ".png"
            fname = f"{doc_id}_{rid}{ext}"
            fpath = os.path.join(IMAGE_FOLDER, fname)
            with open(fpath, "wb") as f: f.write(data)
            try: url = url_for('serve_image', filename=fname)
            except: url = f"/images/{fname}"
            images.append({"filename": fname, "path": fpath, "size": len(data), "url": url})
    return images

# ---- Excel导出 ----
def export_to_excel(results, fn="成绩汇总表.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.title = "成绩汇总"
    hf = Font(name='微软雅黑', bold=True, color='FFFFFF', size=12)
    hfill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    ha = Alignment(horizontal='center', vertical='center')
    tb = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    headers = ['序号', '学号', '姓名', '班级', '课程', '实验名称', '分数', '评语', '截图数量', '批改时间']
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = hf; cell.fill = hfill; cell.alignment = ha; cell.border = tb
    for i, r in enumerate(results, 1):
        row = i + 1
        ws.cell(row=row, column=1, value=i).border = tb
        ws.cell(row=row, column=2, value=r.get('student_id','')).border = tb
        ws.cell(row=row, column=3, value=r.get('name','')).border = tb
        ws.cell(row=row, column=4, value=r.get('class_name','')).border = tb
        ws.cell(row=row, column=5, value=r.get('course','')).border = tb
        ws.cell(row=row, column=6, value=r.get('experiment','')).border = tb
        sc = ws.cell(row=row, column=7, value=r.get('score',0)); sc.border = tb
        ws.cell(row=row, column=8, value=r.get('comment','')).border = tb
        ws.cell(row=row, column=9, value=r.get('image_count',0)).border = tb
        ws.cell(row=row, column=10, value=datetime.now().strftime('%Y-%m-%d %H:%M')).border = tb
        s = r.get('score', 0)
        if s >= 90: sc.fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
        elif s >= 75: sc.fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
        elif s >= 60: sc.fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
        else: sc.fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
    for i, w in enumerate([6, 16, 10, 16, 16, 30, 8, 40, 10, 18], 1):
        ws.column_dimensions[chr(64+i)].width = w
    fp = os.path.join(BASE_DIR, fn)
    wb.save(fp)
    return fp

# ---- Flask路由 ----
@app.route('/')
def index():
    docs = []
    if os.path.exists(REPORTS_FOLDER):
        for f in os.listdir(REPORTS_FOLDER):
            if f.endswith('.docx'):
                mt = datetime.fromtimestamp(os.path.getmtime(os.path.join(REPORTS_FOLDER, f)))
                docs.append({'filename': f, 'modified': mt.strftime('%Y-%m-%d %H:%M'), 'size': f"{os.path.getsize(os.path.join(REPORTS_FOLDER,f))/1024:.1f} KB"})
    return render_template('index.html', documents=docs, results=load_results())

@app.route('/upload', methods=['POST'])
def upload():
    if 'files' not in request.files:
        flash('请选择文件', 'error')
        return redirect(url_for('index'))
    files = request.files.getlist('files')
    if not files or files[0].filename == '':
        flash('请选择文件', 'error'); return redirect(url_for('index'))
    uploaded = []
    for f in files:
        if not f.filename.endswith('.docx'): continue
        fn = f.filename
        if '/' in fn: fn = fn.rsplit('/', 1)[-1]
        elif '\\' in fn: fn = fn.rsplit('\\', 1)[-1]
        fp = os.path.join(REPORTS_FOLDER, fn)
        c = 1
        while os.path.exists(fp):
            n, e = os.path.splitext(fn)
            fp = os.path.join(REPORTS_FOLDER, f"{n}_{c}{e}"); c += 1
        f.save(fp)
        uploaded.append(os.path.basename(fp))
    if uploaded: flash(f'成功上传 {len(uploaded)} 个文档', 'success')
    else: flash('未上传任何 .docx 文件', 'error')
    return redirect(url_for('index'))

@app.route('/grade')
def grade():
    api_avail = is_zhipu_available()
    vmodels = get_available_vision_models()
    grading_mode = "智谱AI视觉评分" if api_avail else "关键词回退评分"
    if not api_avail: flash('⚠️ 未配置智谱API Key，使用关键词回退评分', 'warning')
    all_res = []
    for fname in sorted(os.listdir(REPORTS_FOLDER)):
        if not fname.endswith('.docx'): continue
        fpath = os.path.join(REPORTS_FOLDER, fname)
        doc_id = os.path.splitext(fname)[0]
        try:
            doc = docx.Document(fpath)
            info, text = extract_student_info(doc, fname)
            images = extract_images(doc, doc_id)
            if api_avail and images:
                print(f"\n📝 正在AI评分: {fname} ({len(images)}张截图)")
                score, comment = ai_grade_with_zhipu(text, images)
            else:
                score, comment = auto_grade_fallback(text, images)
            all_res.append({'id': doc_id, 'filename': fname, **info, 'score': score, 'comment': comment,
                          'images': images, 'image_count': len(images), 'text_preview': text[:500]})
        except Exception as e:
            import traceback
            traceback.print_exc()
            all_res.append({'id': doc_id, 'filename': fname, 'student_id': '', 'name': f'[解析失败: {str(e)}]',
                          'class_name': '', 'course': '', 'experiment': '', 'score': 0,
                          'comment': f'文档解析失败：{str(e)}', 'images': [], 'image_count': 0, 'text_preview': ''})
    print(f"\n📊 评分完成: {len(all_res)} 个文档")
    save_results(all_res)
    return render_template('grade.html', results=all_res, grading_mode=grading_mode, api_available=api_avail, api_model=ZHIPU_MODEL, vision_models=vmodels)

@app.route('/update_score', methods=['POST'])
def update_score():
    data = request.get_json()
    results = load_results()
    for r in results:
        if r['id'] == data.get('id'):
            r['score'] = int(data.get('score', r['score']))
            if data.get('comment'): r['comment'] = data['comment']
            break
    save_results(results)
    return jsonify({'status': 'ok'})

@app.route('/export')
def export():
    results = load_results()
    if not results: flash('没有评分数据可以导出', 'error'); return redirect(url_for('index'))
    fp = export_to_excel(results)
    flash(f'Excel已导出', 'success')
    return send_from_directory(BASE_DIR, os.path.basename(fp), as_attachment=True)

@app.route('/export_html')
def export_html():
    results = load_results()
    if not results: flash('没有评分数据可以导出', 'error'); return redirect(url_for('index'))
    html = render_template('report.html', results=results, now=datetime.now())
    with open(REPORT_FILE, 'w', encoding='utf-8') as f: f.write(html)
    flash('HTML报告已生成', 'success')
    return redirect(url_for('index'))

@app.route('/images/<filename>')
def serve_image(filename):
    return send_from_directory(IMAGE_FOLDER, filename)

@app.route('/clear', methods=['POST'])
def clear_all():
    for folder in [REPORTS_FOLDER, IMAGE_FOLDER, UPLOAD_FOLDER]:
        for f in os.listdir(folder):
            try:
                fp = os.path.join(folder, f)
                if os.path.isfile(fp): os.unlink(fp)
            except: pass
    rf = os.path.join(BASE_DIR, '.grading_results.json')
    if os.path.exists(rf): os.unlink(rf)
    flash('已清空所有数据', 'success')
    return redirect(url_for('index'))

@app.route('/api/api_status')
def api_status_api():
    return jsonify({"available": is_zhipu_available(), "provider": "zhipuai", "model": ZHIPU_MODEL, "vision_models": get_available_vision_models()})

def save_results(results):
    data = []
    for r in results:
        d = {k: v for k, v in r.items() if k != 'images'}
        d['image_count'] = len(r.get('images', []))
        d['text_preview'] = r.get('text_preview', '')[:500]
        data.append(d)
    with open(os.path.join(BASE_DIR, '.grading_results.json'), 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def load_results():
    fp = os.path.join(BASE_DIR, '.grading_results.json')
    if os.path.exists(fp):
        with open(fp, 'r', encoding='utf-8') as f: return json.load(f)
    return []

def open_browser():
    time.sleep(2)
    try: webbrowser.open('http://localhost:8900')
    except: pass

if __name__ == '__main__':
    pkg = getattr(sys, 'frozen', False)
    print("=" * 50)
    print("  📚 作业批改系统")
    print(f"  🤖 评分引擎: 智谱AI ({ZHIPU_MODEL})" if ZHIPU_API_KEY else "  ⚠️ 未配置API Key")
    print("=" * 50)
    print("  👉  http://localhost:8900")
    print("=" * 50)
    threading.Thread(target=open_browser, daemon=True).start()
    try: app.run(debug=not pkg, host='127.0.0.1', port=8900)
    except Exception as e:
        print(f"❌ 启动失败: {e}")
        input("按回车键退出...")
