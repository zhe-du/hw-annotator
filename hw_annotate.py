# -*- coding: utf-8 -*-
"""hw_annotate.py —— 作业答案内联标注工具（手动 + answers.json 模式）

把「拿到 docx 试卷 → 做题 → 答案红色加粗内联标注到每道题 → 导出 PDF」固化成
一条可重复运行的命令行流水线。基于 2026-10-05 已跑通的原型（mark_answers.py /
docx2pdf.py）重写为通用 CLI。

用法：
    python hw_annotate.py run  <docx|目录> [-a answers.json ...] [--out 目录]
    python hw_annotate.py validate <docx|目录> [-a answers.json ...]

标注样式（与范例一致）：
    - 选择题：答案字母写在题干末尾，形如 （B）
    - 完形填空：答案字母紧跟每个 ____16____ 空格之后
    - 阅读表达：答案写在题目下方的横线上
    - 全部红色加粗；不产生文末「参考答案」表（旧块会删除）
"""

import argparse
import json
import os
import re
import shutil
import sys
import tempfile
import time
import zipfile
from xml.etree import ElementTree as ET
import gc

# PyInstaller windowed（GUI）模式下 stdout/stderr 为 None，写入会崩溃，重定向到空设备
if sys.stderr is None:
    sys.stderr = open(os.devnull, 'w', encoding='utf-8')
if sys.stdout is None:
    sys.stdout = open(os.devnull, 'w', encoding='utf-8')

# 程序目录（打包后为 exe 所在目录，源码运行时为脚本目录）
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# --------------------------------------------------------------------------
# 默认配置（可用 config.yaml / --config 覆盖，见 load_config）
# --------------------------------------------------------------------------
DEFAULTS = {
    "answer_color": "FF0000",
    "bold": True,
    "remove_reference_block": True,
    "reference_marker": "参考答案",
    "output_suffix": "_含答案",
    "backup_suffix": ".before",
    "pdf_page": "a4",
    "pdf_margin": 36,
    "pdf_compress": True,
}

# --------------------------------------------------------------------------
# OOXML 正则（源自实战）
# --------------------------------------------------------------------------
P_RE = re.compile(r'<w:p(?: [^>]*)?>.*?</w:p>', re.S)
T_RE = re.compile(r'<w:t(?: [^>]*)?>(.*?)</w:t>', re.S)
QUE_RE = re.compile(r'^\s*(\d{1,2})(?:\s*[.．、])?\s+(.*)$', re.S)
OPT_RE = re.compile(r'^\s*[A-D]\s*[.．]')
CLOZE_OPT_RE = re.compile(r'^\s*\d{1,2}\s*[.．]\s*[A-D]\s*[.．]')
BLANK_RUN_RE = re.compile(r'<w:r(?: [^>]*)?>(?:(?!</w:r>).)*?</w:r>', re.S)
EXPR_LINE_RE = re.compile(r'\s*_{20,}\s*')


# --------------------------------------------------------------------------
# 小工具
# --------------------------------------------------------------------------
def esc(s):
    return (s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
             .replace('"', '&quot;'))


def ans_run_letter(letter, cfg):
    """选择题 / 完形：形如 （B） 的红色加粗 run。"""
    b = '<w:b/>' if cfg.get('bold') else ''
    return ('<w:r><w:rPr>%s<w:color w:val="%s"/></w:rPr>'
            '<w:t xml:space="preserve">（%s）</w:t></w:r>'
            % (b, cfg.get('answer_color', 'FF0000'), esc(letter)))


def ans_run_text(text, cfg):
    """阅读表达：整段答案的红色加粗 run。"""
    b = '<w:b/>' if cfg.get('bold') else ''
    return ('<w:r><w:rPr>%s<w:color w:val="%s"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r>'
            % (b, cfg.get('answer_color', 'FF0000'), esc(text)))


def para_text(frag):
    """抽取一个 XML 片段里所有 <w:t> 的纯文本。"""
    return ''.join(m.group(1) for m in T_RE.finditer(frag)).replace('&amp;', '&')


def paragraphs(xml):
    """返回 [(start, end, frag, text), ...]。"""
    return [(m.start(), m.end(), m.group(0), para_text(m.group(0)))
            for m in P_RE.finditer(xml)]


# --------------------------------------------------------------------------
# S1 解析 / 读写
# --------------------------------------------------------------------------
def read_docx(path):
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        data = {n: z.read(n) for n in names}
    return names, data


def write_docx(path, names, data):
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.writestr(n, data[n])


# --------------------------------------------------------------------------
# S3 取答案
# --------------------------------------------------------------------------
def load_answers(path):
    with open(path, 'rb') as f:
        raw = f.read()
    text = None
    # 先按 UTF-8（含/不含 BOM）读，失败再按 GBK/GB18030（中文 ANSI）兜底
    for enc in ('utf-8-sig', 'gb18030'):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        raise ValueError('无法识别编码（既非 UTF-8 也非 GBK）')
    d = json.loads(text)
    for key in ('choice', 'cloze', 'expression'):
        if key not in d:
            d[key] = {}
    return d


def int_keys(d):
    return {int(k): v for k, v in d.items()}


# --------------------------------------------------------------------------
# S5 内联标注（含 S2 识别 + S4 对齐）
# --------------------------------------------------------------------------
def remove_reference_block(xml, cfg, log):
    marker = cfg.get('reference_marker', '参考答案')
    hp = xml.rfind(marker)
    if hp == -1:
        return xml
    start = max(xml.rfind('<w:p ', 0, hp), xml.rfind('<w:p>', 0, hp))
    sect = xml.find('<w:sectPr', hp)
    if sect == -1:
        log.append('!! sectPr 未找到，跳过删除参考答案块')
        return xml
    if start != -1:
        xml = xml[:start] + '<w:p/>' + xml[sect:]
        log.append('已删除文末参考答案块')
    return xml


def find_stem(texts, qn):
    for i, t in enumerate(texts):
        mt = QUE_RE.match(t)
        if mt and int(mt.group(1)) == qn and not CLOZE_OPT_RE.match(t):
            return i
    return None


def find_options_after(texts, i):
    for j in range(i + 1, len(texts)):
        if OPT_RE.match(texts[j]):
            return j
    return None


def annotate_xml(xml, answers, cfg, log):
    stats = {"choice": 0, "cloze": 0, "expression": 0,
             "unmatched": [], "missing": []}
    choice = int_keys(answers.get('choice', {}))
    cloze = int_keys(answers.get('cloze', {}))
    expr = int_keys(answers.get('expression', {}))

    if cfg.get('remove_reference_block'):
        xml = remove_reference_block(xml, cfg, log)

    paras = paragraphs(xml)
    texts = [p[3] for p in paras]

    # ---- 选择题：答案插到题干末尾（选项段的前一段末尾） ----
    inserts = []  # (pos, run_xml)
    for qn in sorted(choice):
        letter = choice[qn]
        i = find_stem(texts, qn)
        if i is None:
            stats['unmatched'].append('choice:%d' % qn)
            continue
        j = find_options_after(texts, i)
        if j is None:
            stats['unmatched'].append('choice:%d' % qn)
            continue
        pos = paras[j - 1][1] - len('</w:p>')  # 插入点必须在 </w:p> 之前
        inserts.append((pos, ans_run_letter(letter, cfg)))
        stats['choice'] += 1

    # ---- 完形填空：答案 run 紧跟每个空白 run 之后 ----
    used = set()
    for m in BLANK_RUN_RE.finditer(xml):
        t = para_text(m.group(0)).strip()
        mm = re.match(r'^_+(\d{1,2})_+$', t)
        if not mm:
            continue
        qn = int(mm.group(1))
        if qn in cloze and qn not in used:
            used.add(qn)
            inserts.append((m.end(), ans_run_letter(cloze[qn], cfg)))
            stats['cloze'] += 1
    stats['missing'] += ['cloze:%d' % qn for qn in sorted(set(cloze) - used)]

    # 反向插入，避免坐标漂移
    for pos, frag in sorted(inserts, key=lambda x: -x[0]):
        xml = xml[:pos] + frag + xml[pos:]

    # ---- 阅读表达：整行下划线替换成答案（保留 pPr）----
    paras = paragraphs(xml)
    texts = [p[3] for p in paras]
    blanks = [i for i, t in enumerate(texts) if EXPR_LINE_RE.fullmatch(t)]
    keys = sorted(expr)
    reps = []  # (start, end, newfrag)
    for k, bi in zip(keys, blanks):
        sx, ex, frag, _ = paras[bi]
        pm = re.search(r'<w:pPr>.*?</w:pPr>', frag, re.S)
        ppr = pm.group(0) if pm else ''
        mppr = re.match(r'(<w:p(?: [^>]*)?>).*?(</w:p>)$', frag, re.S)
        newfrag = mppr.group(1) + ppr + ans_run_text(expr[k], cfg) + '</w:p>'
        reps.append((sx, ex, newfrag))
        stats['expression'] += 1
    if len(blanks) < len(keys):
        log.append('!! 阅读表达横线 %d 条 < 答案 %d 条' % (len(blanks), len(keys)))
        stats['unmatched'] += ['expression:%d' % k for k in keys[len(blanks):]]
    for sx, ex, newfrag in sorted(reps, key=lambda x: -x[0]):
        xml = xml[:sx] + newfrag + xml[ex:]

    return xml, stats


# --------------------------------------------------------------------------
# S6 校验
# --------------------------------------------------------------------------
def validate_xml(xml):
    try:
        ET.fromstring(xml)
        return True, None
    except ET.ParseError as e:
        return False, str(e)


def validate_docx(path):
    """zip 完整 + document.xml 可解析。"""
    try:
        with zipfile.ZipFile(path) as z:
            bad = z.testzip()
            if bad:
                return False, 'zip 损坏: %s' % bad
            xml = z.read('word/document.xml').decode('utf-8')
        return validate_xml(xml)
    except Exception as e:
        return False, str(e)


# --------------------------------------------------------------------------
# S7 导出 PDF（PyMuPDF Story，中英混排）
# --------------------------------------------------------------------------
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def _val(el):
    return None if el is None else el.get(W + 'val')


def load_styles(z):
    styles = {}
    default_sz = None
    try:
        sr = ET.fromstring(z.read('word/styles.xml'))
    except KeyError:
        return styles, default_sz
    dd = sr.find(W + 'docDefaults')
    if dd is not None:
        sz = dd.find(W + 'rPrDefault/' + W + 'rPr/' + W + 'sz')
        default_sz = _val(sz)
    for s in sr.findall(W + 'style'):
        sid = s.get(W + 'styleId')
        nm = s.find(W + 'name')
        rpr = s.find(W + 'rPr')
        ppr = s.find(W + 'pPr')
        styles[sid] = {
            'name': _val(nm) or '',
            'sz': _val(rpr.find(W + 'sz')) if rpr is not None else None,
            'b': (rpr is not None and rpr.find(W + 'b') is not None),
            'jc': _val(ppr.find(W + 'jc')) if ppr is not None else None,
        }
    return styles, default_sz


def _html_esc(t):
    import html
    t = html.escape(t, quote=False)
    t = t.replace('\t', '\u00a0' * 4)
    t = re.sub(r' {2,}', lambda m: '\u00a0' * len(m.group(0)), t)
    return t


def run_html(r, styles, pstyle_sz):
    rpr = r.find(W + 'rPr')
    bold = italic = under = False
    vert = None
    color = None
    sz = pstyle_sz
    if rpr is not None:
        if rpr.find(W + 'b') is not None and _val(rpr.find(W + 'b')) != '0':
            bold = True
        if rpr.find(W + 'i') is not None and _val(rpr.find(W + 'i')) != '0':
            italic = True
        u = rpr.find(W + 'u')
        if u is not None and _val(u) not in (None, 'none'):
            under = True
        v = rpr.find(W + 'vertAlign')
        if v is not None:
            vert = _val(v)
        s = rpr.find(W + 'sz')
        if s is not None:
            sz = _val(s)
        c = rpr.find(W + 'color')
        if c is not None:
            cv = _val(c)
            if cv and cv.lower() != 'auto':
                color = cv
    parts = []
    for ch in r:
        if ch.tag == W + 't':
            parts.append(ch.text or '')
        elif ch.tag == W + 'tab':
            parts.append('\t')
        elif ch.tag in (W + 'br', W + 'cr'):
            parts.append('\n')
    if not parts:
        return ''
    txt = _html_esc(''.join(parts)).replace('\n', '<br/>')
    if not txt:
        return ''
    style = ''
    if sz is not None:
        try:
            style += 'font-size:%.1fpt;' % (int(sz) / 2.0)
        except ValueError:
            pass
    if color:
        style += 'color:#%s;' % color
    if style:
        txt = '<span style="%s">%s</span>' % (style, txt)
    if under:
        txt = '<u>' + txt + '</u>'
    if italic:
        txt = '<i>' + txt + '</i>'
    if bold:
        txt = '<b>' + txt + '</b>'
    if vert == 'superscript':
        txt = '<sup>' + txt + '</sup>'
    elif vert == 'subscript':
        txt = '<sub>' + txt + '</sub>'
    return txt


JC = {'left': 'left', 'start': 'left', 'center': 'center', 'right': 'right',
      'end': 'right', 'both': 'justify', 'distribute': 'justify'}


def para_html(p, styles):
    ppr = p.find(W + 'pPr')
    pstyle = None
    jc = None
    if ppr is not None:
        pstyle = _val(ppr.find(W + 'pStyle'))
        jc = _val(ppr.find(W + 'jc'))
    info = styles.get(pstyle, {}) if pstyle else {}
    name = (info.get('name') or '').lower()
    pstyle_sz = info.get('sz')
    if jc is None:
        jc = info.get('jc')
    body = ''.join(run_html(r, styles, pstyle_sz) for r in p.findall(W + 'r'))
    for h in p.findall(W + 'hyperlink'):
        body += ''.join(run_html(r, styles, pstyle_sz) for r in h.findall(W + 'r'))
    align = JC.get(jc, 'left') if jc else 'left'
    style = 'text-align:%s;' % align
    tag = 'p'
    if 'heading 1' in name or name == 'title':
        tag = 'h1'
    elif 'heading 2' in name:
        tag = 'h2'
    elif 'heading 3' in name:
        tag = 'h3'
    elif 'heading 4' in name:
        tag = 'h4'
    if not body.strip():
        body = '\u00a0'
        style += 'font-size:6pt;'
    return '<%s style="%s">%s</%s>' % (tag, style, body, tag)


def table_html(tbl, styles):
    rows = []
    for tr in tbl.findall(W + 'tr'):
        cells = []
        for tc in tr.findall(W + 'tc'):
            inner = ''.join(para_html(p, styles) for p in tc.findall(W + 'p'))
            cells.append('<td>%s</td>' % inner)
        rows.append('<tr>%s</tr>' % ''.join(cells))
    return '<table><tbody>%s</tbody></table>' % ''.join(rows)


CSS = """
* { font-family: sans-serif; font-size: @DEFPT@pt; line-height: 1.32; }
p { margin: 0 0 3pt 0; }
h1 { font-size: 18pt; margin: 8pt 0 6pt 0; }
h2 { font-size: 15pt; margin: 8pt 0 5pt 0; }
h3 { font-size: 13pt; margin: 7pt 0 4pt 0; }
h4 { font-size: 12pt; margin: 6pt 0 3pt 0; }
table { border-collapse: collapse; margin: 6pt 0; width: 100%; }
td { border: 0.6pt solid #333333; padding: 3pt 5pt; vertical-align: top; }
"""


def export_pdf(docx_path, pdf_path, cfg):
    import pymupdf as fitz
    z = zipfile.ZipFile(docx_path)
    styles, default_sz = load_styles(z)
    try:
        default_pt = int(default_sz) / 2.0
    except (TypeError, ValueError):
        default_pt = 10.5
    root = ET.fromstring(z.read('word/document.xml'))
    body = root.find(W + 'body')
    chunks = []
    for child in body:
        if child.tag == W + 'p':
            chunks.append(para_html(child, styles))
        elif child.tag == W + 'tbl':
            chunks.append(table_html(child, styles))
    htmlstr = '<div>%s</div>' % ''.join(chunks)
    css = CSS.replace('@DEFPT@', '%.1f' % default_pt)
    page = cfg.get('pdf_page', 'a4')
    margin = int(cfg.get('pdf_margin', 36))
    mediabox = fitz.paper_rect(page)
    where = mediabox + (margin, margin, -margin, -margin)
    raw = os.path.join(tempfile.gettempdir(),
                        'hw_annotate_%d_%d.pdf' % (os.getpid(), int(time.time() * 1000)))
    writer = fitz.DocumentWriter(raw)
    story = fitz.Story(html=htmlstr, user_css=css)
    more = 1
    while more:
        dev = writer.begin_page(mediabox)
        more, _ = story.place(where)
        story.draw(dev)
        writer.end_page()
    writer.close()
    del writer, story, dev  # 释放 DocumentWriter 句柄，否则 os.remove(raw) 报 WinError 32
    gc.collect()
    if cfg.get('pdf_compress', True):
        d = fitz.open(raw)
        d.subset_fonts()
        d.save(pdf_path, garbage=4, deflate=True, clean=True)
        d.close()
    else:
        shutil.move(raw, pdf_path)
    try:
        os.remove(raw)
    except OSError as e:
        sys.stderr.write('!! 临时文件清理失败（可忽略）: %s\n' % e)
    d = fitz.open(pdf_path)
    pages = d.page_count
    d.close()
    return pages


# --------------------------------------------------------------------------
# 配置
# --------------------------------------------------------------------------
def load_config(path):
    cfg = dict(DEFAULTS)
    if not path:
        return cfg
    # 相对路径：优先当前工作目录，其次程序目录（随包 config.yaml）
    if not os.path.isabs(path):
        p = os.path.join(os.getcwd(), path)
        if not os.path.isfile(p):
            p = os.path.join(BASE_DIR, path)
    else:
        p = path
    if not os.path.isfile(p):
        return cfg
    for line in open(p, 'r', encoding='utf-8-sig'):
        line = line.split('#')[0].strip()
        if not line or ':' not in line:
            continue
        k, v = line.split(':', 1)
        k, v = k.strip(), v.strip()
        if v == 'true':
            v = True
        elif v == 'false':
            v = False
        else:
            try:
                v = int(v)
            except ValueError:
                pass
        cfg[k] = v
    return cfg


# --------------------------------------------------------------------------
# 单文件处理
# --------------------------------------------------------------------------
def process_one(path, answers, cfg, out_dir, keep_before=True):
    stem, ext = os.path.splitext(os.path.basename(path))
    out_docx = os.path.join(out_dir, stem + cfg['output_suffix'] + ext)
    out_pdf = os.path.join(out_dir, stem + cfg['output_suffix'] + '.pdf')
    before = os.path.join(out_dir, stem + cfg['backup_suffix'] + ext)
    report = {
        'file': os.path.basename(path),
        'status': 'ok',
        'choice_total': len(answers.get('choice', {})),
        'cloze_total': len(answers.get('cloze', {})),
        'expression_total': len(answers.get('expression', {})),
        'matched': {}, 'unmatched': [], 'missing': [],
        'outputs': {}, 'errors': [],
    }

    # 读入（坏文件/缺 document.xml 时记 failed 并跳过，不中断整批）
    try:
        names, data = read_docx(path)
        xml = data['word/document.xml'].decode('utf-8')
    except Exception as e:
        report['status'] = 'failed'
        report['errors'].append('read_failed: %s' % e)
        # 解析失败时题量统计置 0，避免误导
        report['choice_total'] = 0
        report['cloze_total'] = 0
        report['expression_total'] = 0
        return report, []

    # 标注
    log = []
    new_xml, stats = annotate_xml(xml, answers, cfg, log)
    report['matched'] = {'choice': stats['choice'], 'cloze': stats['cloze'],
                         'expression': stats['expression']}
    report['unmatched'] = stats['unmatched']
    report['missing'] = stats['missing']

    # 校验标注后的 XML
    ok, err = validate_xml(new_xml)
    if not ok:
        report['status'] = 'failed'
        report['errors'].append('invalid_xml: %s' % err)
        return report, log

    # 写盘（先备份原文件快照）
    if keep_before:
        shutil.copy2(path, before)
        report['outputs']['before'] = before
    data['word/document.xml'] = new_xml.encode('utf-8')
    written = False
    for _ in range(3):
        try:
            write_docx(out_docx, names, data)
            written = True
            break
        except PermissionError:
            time.sleep(1)
    if not written:
        report['status'] = 'failed'
        report['errors'].append('file_locked: %s' % out_docx)
        return report, log
    report['outputs']['docx'] = out_docx

    # 校验输出 docx（失败则删除坏文件，避免残留）
    ok, err = validate_docx(out_docx)
    if not ok:
        report['status'] = 'failed'
        report['errors'].append('invalid_output_docx: %s' % err)
        try:
            os.remove(out_docx)
        except OSError:
            pass
        return report, log

    # 导出 PDF
    try:
        pages = export_pdf(out_docx, out_pdf, cfg)
        report['outputs']['pdf'] = out_pdf
        report['pdf_pages'] = pages
    except Exception as e:
        report['errors'].append('pdf_skipped: %s' % e)

    return report, log


# --------------------------------------------------------------------------
# 输入收集 / 答案匹配
# --------------------------------------------------------------------------
def collect_inputs(path):
    if os.path.isfile(path):
        return [path]
    out = []
    for root, _, files in os.walk(path):
        for f in sorted(files):
            if not f.lower().endswith('.docx'):
                continue
            if f.startswith('~$'):
                continue
            if '_含答案' in f or '.before' in f:
                continue
            out.append(os.path.join(root, f))
    return out


def match_answers(docx_path, answer_paths, errors=None):
    base = os.path.basename(docx_path)
    stem = os.path.splitext(base)[0]
    loaded = []
    for p in answer_paths:
        try:
            loaded.append((p, load_answers(p)))
        except Exception as e:
            msg = '无法读取答案 %s: %s' % (p, e)
            sys.stderr.write('!! %s\n' % msg)
            if errors is not None:
                errors.append(msg)
    # 1) 显式 file 字段匹配
    for p, d in loaded:
        if d.get('file') and d['file'] == base:
            return d
    # 2) 同名兄弟 json
    for cand in ('%s.json' % stem, '%s.answers.json' % stem):
        cpath = os.path.join(os.path.dirname(docx_path), cand)
        if os.path.isfile(cpath):
            try:
                return load_answers(cpath)
            except Exception as e:
                msg = '无法读取答案 %s: %s' % (cpath, e)
                sys.stderr.write('!! %s\n' % msg)
                if errors is not None:
                    errors.append(msg)
    # 3) 只剩一份答案时兜底（file 字段不匹配时给出警告）
    if len(loaded) == 1:
        p, d = loaded[0]
        if d.get('file') and d['file'] != base:
            sys.stderr.write('!! 警告：%s 的 file 字段(%s) 与 %s 不匹配，仍将使用\n'
                             % (p, d['file'], base))
        return d
    return None


def write_report(out_dir, reports):
    rp = os.path.join(out_dir, 'report.json')
    with open(rp, 'w', encoding='utf-8') as f:
        json.dump(reports, f, ensure_ascii=False, indent=2)
    lines = ['# 标注报告', '']
    for r in reports:
        lines.append('## %s — %s' % (r['file'], r['status']))
        lines.append('')
        lines.append('| 题型 | 标注 | 总量 |')
        lines.append('|---|---|---|')
        lines.append('| 选择题 | %d | %d |' % (r['matched'].get('choice', 0), r['choice_total']))
        lines.append('| 完形填空 | %d | %d |' % (r['matched'].get('cloze', 0), r['cloze_total']))
        lines.append('| 阅读表达 | %d | %d |' % (r['matched'].get('expression', 0), r['expression_total']))
        lines.append('')
        if r['unmatched']:
            lines.append('- 未识别：' + '、'.join(r['unmatched']))
        if r['missing']:
            lines.append('- 答案缺失：' + '、'.join(r['missing']))
        if r['errors']:
            lines.append('- 异常：' + '、'.join(r['errors']))
        lines.append('')
    with open(os.path.join(out_dir, 'report.md'), 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    return rp


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------
def cmd_run(args):
    cfg = load_config(args.config)
    inputs = collect_inputs(args.input)
    if not inputs:
        print('未找到可处理的 .docx 文件：%s' % args.input)
        return 1
    answer_paths = list(args.answers or [])
    out_dir = args.out or (os.path.dirname(args.input) if os.path.isfile(args.input) else args.input)
    os.makedirs(out_dir, exist_ok=True)

    t0 = time.time()
    reports = []
    answer_errors = []
    for path in inputs:
        answers = match_answers(path, answer_paths, answer_errors)
        if answers is None:
            print('[跳过] %s（未找到匹配的 answers.json）' % os.path.basename(path))
            reports.append({'file': os.path.basename(path), 'status': 'skipped',
                            'errors': ['no_answers'], 'matched': {}, 'unmatched': [],
                            'missing': [], 'choice_total': 0, 'cloze_total': 0,
                            'expression_total': 0, 'outputs': {}})
            continue
        report, log = process_one(path, answers, cfg, out_dir, keep_before=True)
        reports.append(report)
        print('[%s] %s' % (report['status'], os.path.basename(path)))
        for line in log:
            print('    ', line)
        for e in report['errors']:
            print('    !!', e)

    rp = write_report(out_dir, reports)
    print('报告：%s（%.2fs）' % (rp, time.time() - t0))
    failed = any(r['status'] == 'failed' for r in reports) or bool(answer_errors)
    return 1 if failed else 0


def cmd_validate(args):
    cfg = load_config(args.config)
    inputs = collect_inputs(args.input)
    answer_paths = list(args.answers or [])
    failed = 0
    answer_errors = []
    for path in inputs:
        answers = match_answers(path, answer_paths, answer_errors)
        if answers is None:
            print('[无答案] %s' % os.path.basename(path))
            continue
        try:
            names, data = read_docx(path)
            xml = data['word/document.xml'].decode('utf-8')
        except Exception as e:
            print('[失败] %s：%s' % (os.path.basename(path), e))
            failed += 1
            continue
        log = []
        _, stats = annotate_xml(xml, answers, cfg, log)
        print('== %s' % os.path.basename(path))
        print('   选择题 %d / 完形 %d / 阅读表达 %d' %
              (stats['choice'], stats['cloze'], stats['expression']))
        if stats['unmatched']:
            print('   未识别：', '、'.join(stats['unmatched']))
        if stats['missing']:
            print('   缺失：', '、'.join(stats['missing']))
    return 1 if (failed or answer_errors) else 0


def _fix_console_encoding():
    """中文 Windows 保持 GBK；非中文区域转 UTF-8 避免 UnicodeEncodeError。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            enc = getattr(stream, 'encoding', None) or ''
            if enc.lower().replace('-', '') in ('utf8', 'cp65001'):
                continue
            try:
                '中文'.encode(enc)
            except (LookupError, UnicodeEncodeError):
                stream.reconfigure(encoding='utf-8', errors='replace')
        except Exception:
            pass


def main(argv=None):
    _fix_console_encoding()
    ap = argparse.ArgumentParser(prog='hw_annotate',
                                 description='作业答案内联标注工具（手动 + answers.json）')
    sub = ap.add_subparsers(dest='cmd', required=True)

    p = sub.add_parser('run', help='处理指定文件或目录')
    p.add_argument('input', help='单个 .docx 路径或目录（递归 *.docx）')
    p.add_argument('-a', '--answers', nargs='*', help='answers.json（可多个）')
    p.add_argument('--out', help='输出目录（默认：输入文件所在目录）')
    p.add_argument('--config', default='config.yaml',
                   help='配置文件路径（默认 config.yaml，不存在则用内置默认）')
    p.set_defaults(func=cmd_run)

    v = sub.add_parser('validate', help='只校验不写盘')
    v.add_argument('input', help='单个 .docx 路径或目录')
    v.add_argument('-a', '--answers', nargs='*', help='answers.json')
    v.add_argument('--config', default='config.yaml', help='配置文件路径')
    v.set_defaults(func=cmd_validate)

    args = ap.parse_args(argv)
    sys.exit(args.func(args))


if __name__ == '__main__':
    main()
