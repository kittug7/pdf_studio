"""PDF Studio - single-file PDF toolkit for Windows (Python + Tkinter + PyMuPDF)."""
import io, os, re, shutil, subprocess, sys, tempfile, threading, queue
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
from pathlib import Path
import fitz
from PIL import Image, ImageOps

RED, BG = "#4f46e5", "#f3f4f8"
FN, FB = ("Segoe UI", 10), ("Segoe UI", 10, "bold")


# ───────────────────────── helpers ─────────────────────────
def uniq(p):
    p = Path(p); q, i = p, 1
    while q.exists():
        q = p.with_name(f"{p.stem} ({i}){p.suffix}"); i += 1
    return q

def pgs(spec, n):
    spec = (spec or "").strip()
    if not spec or spec.lower() == "all": return list(range(n))
    r = []
    for part in re.split(r"[,\s]+", spec):
        if not part: continue
        m = re.fullmatch(r"(\d*)-(\d*)", part)
        if m: a, b = int(m[1] or 1), int(m[2] or n)
        elif part.isdigit(): a = b = int(part)
        else: raise ValueError(f"Invalid page range: {part}")
        if a < 1 or a > n or b < 1: raise ValueError(f"Page {part} is out of range (PDF has {n} pages)")
        r += range(a - 1, min(b, n))
    return r

def opn(p):
    d = fitz.open(p)
    if d.needs_pass: raise ValueError(f"{Path(p).name} is password protected - use Unlock PDF first.")
    return d

def sv(d, path, **kw):
    d.save(str(path), garbage=3, deflate=True, **kw); d.close(); return [path]

def each(files, out, suffix, fn):
    r = []
    for f in files:
        d = opn(f); fn(d); r += sv(d, uniq(out / f"{Path(f).stem}_{suffix}.pdf"))
    return r

def rgb(h): h = h.lstrip("#"); return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))

def num(c):
    if c is None: return ""
    s = str(c).replace(",", "")
    return float(s) if re.fullmatch(r"-?\d+(\.\d+)?", s) else c


# ───────────────────────── tools (backend) ─────────────────────────
def merge(F, o, out):
    d = fitz.open()
    for f in F: d.insert_pdf(opn(f))
    return sv(d, uniq(out / "merged.pdf"))

def split(F, o, out):
    src = opn(F[0]); n = len(src); s = Path(F[0]).stem
    if o["mode"] == "ranges": groups = [pgs(x, n) for x in re.split(r"[,;]", o["ranges"]) if x.strip()]
    elif o["mode"] == "every":
        k = max(1, int(o["n"])); groups = [list(range(i, min(i + k, n))) for i in range(0, n, k)]
    else: groups = [[i] for i in range(n)]
    res = []
    for j, g in enumerate(groups):
        d = fitz.open()
        for i in g: d.insert_pdf(src, from_page=i, to_page=i)
        res += sv(d, uniq(out / f"{s}_part{j + 1}.pdf"))
    return res

def remove(F, o, out):
    def go(d):
        n = len(d); kill = set(pgs(o["pages"], n)); keep = [i for i in range(n) if i not in kill]
        if not keep: raise ValueError("You cannot remove every page.")
        d.select(keep)
    return each(F, out, "removed", go)

def extract(F, o, out):
    return each(F, out, "extracted", lambda d: d.select(pgs(o["pages"], len(d))))

def compress(F, o, out):
    q, mx = {"extreme": (35, 900), "recommended": (60, 1500), "low": (80, 2400)}[o["level"]]; res = []
    for f in F:
        d = opn(f); seen = set()
        for pg in d:
            for im in pg.get_images(full=True):
                x = im[0]
                if x in seen or im[1]: continue
                seen.add(x)
                try:
                    b = d.extract_image(x)["image"]
                    if len(b) < 20000: continue
                    pil = Image.open(io.BytesIO(b)).convert("RGB"); pil.thumbnail((mx, mx))
                    buf = io.BytesIO(); pil.save(buf, "JPEG", quality=q, optimize=True)
                    if buf.tell() < len(b): pg.replace_image(x, stream=buf.getvalue())
                except Exception: pass
        p = uniq(out / f"{Path(f).stem}_compressed.pdf")
        d.save(str(p), garbage=4, deflate=True, clean=True); d.close()
        if os.path.getsize(p) >= os.path.getsize(f): shutil.copy(f, p)
        res.append(p)
    return res

def repair(F, o, out):
    res = []
    for f in F:
        d = fitz.open(f)
        if d.needs_pass: raise ValueError("Password protected file.")
        res += sv(d, uniq(out / f"{Path(f).stem}_repaired.pdf"), clean=True)
    return res

def jpg2pdf(F, o, out):
    d = fitz.open(); sizes = {"fit": None, "a4": (595, 842), "letter": (612, 792)}[o["size"]]; mg = int(o["margin"])
    for f in F:
        im = ImageOps.exif_transpose(Image.open(f)).convert("RGB"); w, h = im.size
        if sizes: pw, ph = sizes if h >= w else sizes[::-1]
        else: pw, ph = w * .75, h * .75
        buf = io.BytesIO(); im.save(buf, "JPEG", quality=92)
        d.new_page(width=pw, height=ph).insert_image(fitz.Rect(mg, mg, pw - mg, ph - mg), stream=buf.getvalue())
    return sv(d, uniq(out / "images.pdf"))

def com_convert(ext, src, dst):
    import pythoncom, win32com.client as w
    pythoncom.CoInitialize()
    if ext in (".doc", ".docx", ".rtf", ".txt", ".odt"):
        a = w.DispatchEx("Word.Application"); a.Visible = False
        try: x = a.Documents.Open(src, ReadOnly=True); x.ExportAsFixedFormat(dst, 17); x.Close(False)
        finally: a.Quit()
    elif ext in (".ppt", ".pptx", ".odp"):
        a = w.DispatchEx("PowerPoint.Application")
        try: x = a.Presentations.Open(src, True, False, False); x.SaveAs(dst, 32); x.Close()
        finally: a.Quit()
    else:
        a = w.DispatchEx("Excel.Application"); a.Visible = False; a.DisplayAlerts = False
        try: x = a.Workbooks.Open(src, ReadOnly=True); x.ExportAsFixedFormat(0, dst); x.Close(False)
        finally: a.Quit()

def office(F, o, out):
    res = []
    for f in F:
        src = str(Path(f).resolve()); dst = uniq(out / (Path(f).stem + ".pdf")); ext = Path(f).suffix.lower()
        try: com_convert(ext, src, str(dst.resolve()))
        except Exception as e:
            lo = shutil.which("soffice") or next((p for p in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                  r"C:\Program Files (x86)\LibreOffice\program\soffice.exe") if os.path.exists(p)), None)
            if not lo: raise RuntimeError(f"Install Microsoft Office or LibreOffice to convert this file.\n({e})")
            subprocess.run([lo, "--headless", "--convert-to", "pdf", "--outdir", str(out), src], check=True, timeout=300)
            dst = out / (Path(f).stem + ".pdf")
        res.append(dst)
    return res

def html2pdf(F, o, out):
    b = next((p for p in (r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
              r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
              r"C:\Program Files\Google\Chrome\Application\chrome.exe",
              r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe") if os.path.exists(p)), None)
    if not b: raise RuntimeError("Microsoft Edge or Google Chrome is required.")
    srcs = list(F) + ([o["url"].strip()] if o["url"].strip() else []); res = []
    for s in srcs:
        url = s if s.startswith("http") else Path(s).resolve().as_uri()
        name = Path(s).stem if not s.startswith("http") else re.sub(r"\W+", "_", s.split("//")[-1])[:40]
        dst = uniq(out / f"{name}.pdf")
        subprocess.run([b, "--headless", "--disable-gpu", "--no-pdf-header-footer", f"--user-data-dir={tempfile.mkdtemp()}",
                        f"--print-to-pdf={dst}", url], check=True, timeout=120)
        res.append(dst)
    return res

def pdf2jpg(F, o, out):
    res = []
    for f in F:
        d = opn(f); folder = uniq(out / f"{Path(f).stem}_images"); folder.mkdir(parents=True)
        for i in pgs(o["pages"], len(d)):
            pix = d[i].get_pixmap(dpi=int(o["dpi"])); im = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            p = folder / f"page_{i + 1}.{o['fmt']}"; im.save(p, quality=90); res.append(p)
    return res

def pdf2word(F, o, out):
    from pdf2docx import Converter
    res = []
    for f in F:
        dst = uniq(out / f"{Path(f).stem}.docx"); c = Converter(f); c.convert(str(dst)); c.close(); res.append(dst)
    return res

def pdf2excel(F, o, out):
    import openpyxl
    res = []
    for f in F:
        d = opn(f); wb = openpyxl.Workbook(); wb.remove(wb.active); hit = False
        for pg in d:
            for j, t in enumerate(pg.find_tables().tables):
                ws = wb.create_sheet(f"P{pg.number + 1}_T{j + 1}"); hit = True
                for row in t.extract(): ws.append([num(c) for c in row])
        if not hit:
            ws = wb.create_sheet("Text")
            for pg in d:
                for ln in pg.get_text().splitlines(): ws.append([ln])
        p = uniq(out / f"{Path(f).stem}.xlsx"); wb.save(str(p)); res.append(p)
    return res

def rotate(F, o, out):
    def go(d):
        for i in pgs(o["pages"], len(d)): p = d[i]; p.set_rotation((p.rotation + int(o["angle"])) % 360)
    return each(F, out, "rotated", go)

def pagenum(F, o, out):
    def go(d):
        n, fs, st = len(d), int(o["size"]), int(o["start"])
        for k, i in enumerate(pgs(o["pages"], n)):
            p = d[i]; t = o["fmt"].format(n=st + k, total=n); w = fitz.get_text_length(t, "helv", fs); r = p.rect
            x = {"l": 36, "c": (r.width - w) / 2, "r": r.width - 36 - w}[o["pos"][1]]
            y = 30 + fs if o["pos"][0] == "t" else r.height - 30
            p.insert_text(fitz.Point(x, y) * p.derotation_matrix, t, fontsize=fs, fontname="helv", rotate=p.rotation)
    return each(F, out, "numbered", go)

def watermark(F, o, out):
    op, ang, over = float(o["opacity"]) / 100, float(o["angle"]), o["layer"] == "over"
    def go(d):
        for i in pgs(o["pages"], len(d)):
            p = d[i]; c = fitz.Point(p.rect.width / 2, p.rect.height / 2) * p.derotation_matrix
            if o["kind"] == "text":
                fs = int(o["size"]); w = fitz.get_text_length(o["text"], "helv", fs)
                p.insert_text(fitz.Point(c.x - w / 2, c.y + fs * .3), o["text"], fontsize=fs, fontname="helv",
                              color=rgb(o["color"]), fill_opacity=op, morph=(c, fitz.Matrix(ang)), overlay=over)
            else:
                im = Image.open(o["image"]).convert("RGBA"); im.putalpha(im.getchannel("A").point(lambda a: int(a * op)))
                im = im.rotate(ang, expand=True); buf = io.BytesIO(); im.save(buf, "PNG")
                w = p.rect.width * .4; h = w * im.height / im.width
                p.insert_image(fitz.Rect(c.x - w / 2, c.y - h / 2, c.x + w / 2, c.y + h / 2), stream=buf.getvalue(), overlay=over)
    return each(F, out, "watermarked", go)

def crop(F, o, out):
    l, t, r, b = (float(o[k]) / 100 for k in ("l", "t", "r", "b"))
    def go(d):
        for i in pgs(o["pages"], len(d)):
            p = d[i]; R = p.rect
            nr = fitz.Rect(R.x0 + R.width * l, R.y0 + R.height * t, R.x1 - R.width * r, R.y1 - R.height * b) * p.derotation_matrix
            nr.normalize(); p.set_cropbox(nr)
    return each(F, out, "cropped", go)

def unlock(F, o, out):
    res = []
    for f in F:
        d = fitz.open(f)
        if d.needs_pass and not d.authenticate(o["pw"]): raise ValueError(f"Wrong password for {Path(f).name}")
        res += sv(d, uniq(out / f"{Path(f).stem}_unlocked.pdf"), encryption=fitz.PDF_ENCRYPT_NONE)
    return res

def protect(F, o, out):
    if not o["pw"]: raise ValueError("Enter a password.")
    perm = fitz.PDF_PERM_ACCESSIBILITY
    if o["print"]: perm |= fitz.PDF_PERM_PRINT
    if o["copy"]: perm |= fitz.PDF_PERM_COPY
    if o["edit"]: perm |= fitz.PDF_PERM_MODIFY | fitz.PDF_PERM_ANNOTATE | fitz.PDF_PERM_FORM | fitz.PDF_PERM_ASSEMBLE
    res = []
    for f in F:
        res += sv(opn(f), uniq(out / f"{Path(f).stem}_protected.pdf"), encryption=fitz.PDF_ENCRYPT_AES_256,
                  user_pw=o["pw"], owner_pw=o["owner"] or o["pw"], permissions=perm)
    return res


# ───────────────────────── tool registry ─────────────────────────
def Fd(k, l, t="e", d="", ch=None, sh=None): return dict(k=k, l=l, t=t, d=d, ch=ch, sh=sh)
PG = lambda d="all": Fd("pages", "Pages (e.g. 1-3,5,8-)", "e", d)
def T(id, title, desc, cat, icon, kind, fn, f=(), m=1, mx=0): return dict(id=id, t=title, d=desc, c=cat, i=icon, k=kind, fn=fn, f=list(f), m=m, mx=mx)

CATS = ["Organize", "Optimize", "Convert to PDF", "Convert from PDF", "Edit", "Security"]
TOOLS = [
    T("merge", "Merge PDF", "Combine PDFs in the order you want.", "Organize", "🗂️", "pdf", merge, m=2),
    T("split", "Split PDF", "Separate one page or a whole set into independent PDFs.", "Organize", "✂️", "pdf", split, mx=1, f=[
        Fd("mode", "Split mode", "c", "ranges", [("Custom ranges", "ranges"), ("Every N pages", "every"), ("Every page", "pages")]),
        Fd("ranges", "Ranges (each comma = new file)", "e", "1-3,4-6", sh=("mode", "ranges")),
        Fd("n", "Pages per file", "e", "2", sh=("mode", "every"))]),
    T("remove", "Remove pages", "Delete the pages you don't need.", "Organize", "🗑️", "pdf", remove, f=[PG("1")]),
    T("extract", "Extract pages", "Keep only the pages you select.", "Organize", "📑", "pdf", extract, f=[PG("1")]),
    T("compress", "Compress PDF", "Reduce file size while keeping quality.", "Optimize", "🗜️", "pdf", compress, f=[
        Fd("level", "Compression level", "c", "recommended", [("Extreme (smallest)", "extreme"), ("Recommended", "recommended"), ("Low (best quality)", "low")])]),
    T("repair", "Repair PDF", "Recover data from a damaged PDF.", "Optimize", "🛠️", "pdf", repair),
    T("jpg2pdf", "JPG to PDF", "Convert images to PDF with page size & margin.", "Convert to PDF", "🖼️", "image", jpg2pdf, f=[
        Fd("size", "Page size", "c", "a4", [("A4", "a4"), ("US Letter", "letter"), ("Fit to image", "fit")]),
        Fd("margin", "Margin", "c", "0", [("None", "0"), ("Small", "20"), ("Big", "40")])]),
    T("word2pdf", "Word to PDF", "Make DOC and DOCX files easy to read as PDF.", "Convert to PDF", "📘", "word", office),
    T("ppt2pdf", "PowerPoint to PDF", "Turn PPT and PPTX slideshows into PDF.", "Convert to PDF", "📙", "ppt", office),
    T("excel2pdf", "Excel to PDF", "Make Excel spreadsheets easy to read as PDF.", "Convert to PDF", "📗", "excel", office),
    T("html2pdf", "HTML to PDF", "Convert web pages or HTML files to PDF.", "Convert to PDF", "🌐", "html", html2pdf, m=0, f=[
        Fd("url", "Or web page URL", "e", "https://")]),
    T("pdf2jpg", "PDF to JPG", "Convert each PDF page into an image.", "Convert from PDF", "📷", "pdf", pdf2jpg, f=[
        Fd("fmt", "Format", "c", "jpg", [("JPG", "jpg"), ("PNG", "png")]),
        Fd("dpi", "Quality (DPI)", "c", "150", [("Low (100)", "100"), ("Medium (150)", "150"), ("High (200)", "200"), ("Best (300)", "300")]), PG()]),
    T("pdf2word", "PDF to Word", "Convert PDF to an editable DOCX.", "Convert from PDF", "📝", "pdf", pdf2word),
    T("pdf2excel", "PDF to Excel", "Pull tables from PDF into Excel sheets.", "Convert from PDF", "📊", "pdf", pdf2excel),
    T("rotate", "Rotate PDF", "Rotate all or selected pages.", "Edit", "🔄", "pdf", rotate, f=[
        Fd("angle", "Rotation", "c", "90", [("90° clockwise", "90"), ("180°", "180"), ("90° counter-clockwise", "270")]), PG()]),
    T("pagenum", "Add page numbers", "Add page numbers with your own style.", "Edit", "🔢", "pdf", pagenum, f=[
        Fd("pos", "Position", "c", "bc", [("Bottom center", "bc"), ("Bottom right", "br"), ("Bottom left", "bl"), ("Top center", "tc"), ("Top right", "tr"), ("Top left", "tl")]),
        Fd("fmt", "Format", "c", "{n}", [("1", "{n}"), ("Page 1", "Page {n}"), ("Page 1 of N", "Page {n} of {total}")]),
        Fd("start", "First number", "e", "1"), Fd("size", "Font size", "e", "11"), PG()]),
    T("watermark", "Add watermark", "Stamp text or an image over / under your PDF.", "Edit", "💧", "pdf", watermark, f=[
        Fd("kind", "Type", "c", "text", [("Text", "text"), ("Image", "image")]),
        Fd("text", "Text", "e", "CONFIDENTIAL", sh=("kind", "text")), Fd("size", "Font size", "e", "60", sh=("kind", "text")),
        Fd("color", "Color", "col", "#cc0000", sh=("kind", "text")), Fd("image", "Image file", "img", "", sh=("kind", "image")),
        Fd("opacity", "Opacity %", "e", "30"), Fd("angle", "Rotation (°)", "e", "45"),
        Fd("layer", "Layer", "c", "over", [("Over content", "over"), ("Below content", "under")]), PG()]),
    T("crop", "Crop PDF", "Trim margins of all or selected pages.", "Edit", "⛶", "pdf", crop, f=[
        Fd("l", "Left %", "e", "0"), Fd("t", "Top %", "e", "0"), Fd("r", "Right %", "e", "0"), Fd("b", "Bottom %", "e", "0"), PG()]),
    T("unlock", "Unlock PDF", "Remove PDF password security.", "Security", "🔓", "pdf", unlock, f=[Fd("pw", "Current password", "p")]),
    T("protect", "Protect PDF", "Encrypt with a password (AES-256).", "Security", "🔒", "pdf", protect, mx=0, f=[
        Fd("pw", "Open password", "p"), Fd("owner", "Owner password (optional)", "p"),
        Fd("print", "Allow printing", "k", True), Fd("copy", "Allow copying text", "k", True), Fd("edit", "Allow editing", "k", False)]),
]
TYPES = {"pdf": [("PDF", "*.pdf")], "image": [("Images", "*.jpg *.jpeg *.png *.bmp *.webp *.tif *.tiff *.gif")],
         "word": [("Word / text", "*.doc *.docx *.rtf *.odt *.txt")], "ppt": [("PowerPoint", "*.ppt *.pptx *.odp")],
         "excel": [("Excel", "*.xls *.xlsx *.csv *.ods")], "html": [("HTML", "*.html *.htm")]}

def reveal(p):
    try: os.startfile(str(p))
    except AttributeError: subprocess.Popen(["xdg-open", str(p)])


# ───────────────────────── accounts · 90-day trial · stats ─────────────────────────
def _make_store():
    import datetime as dt, hashlib, hmac, json, os, re, secrets, types
    from pathlib import Path

    D = Path(os.environ.get("APPDATA", Path.home())) / "PDFStudio"; D.mkdir(parents=True, exist_ok=True)
    KEY = b"change-this-secret-before-you-ship"   # signs the trial start date so it can't be edited by hand
    TRIAL_DAYS = 90

    def _rj(n, default):
        try: return json.loads((D / n).read_text("utf-8"))
        except Exception: return default
    def _wj(n, v): (D / n).write_text(json.dumps(v, indent=1), "utf-8")
    def _sig(s): return hmac.new(KEY, s.encode(), hashlib.sha256).hexdigest()
    def _ok(s, g): return bool(s) and hmac.compare_digest(g or "", _sig(s))

    def _reg(val=None):                              # second copy of the trial date in the Windows registry
        try:
            import winreg
            k = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\PDFStudio")
            if val: winreg.SetValueEx(k, "cfg", 0, winreg.REG_SZ, val); return val
            return winreg.QueryValueEx(k, "cfg")[0]
        except Exception: return None

    def trial_start():
        found, bad = [], False
        f = _rj("license.json", None)
        if f is not None:
            if _ok(f.get("start"), f.get("sig")): found.append(f["start"])
            else: bad = True
        r = _reg()
        if r:
            s, _, g = r.partition("|")
            if _ok(s, g): found.append(s)
            else: bad = True
        if bad and not found: return "2000-01-01"            # tampered -> trial treated as expired
        s = min(found) if found else dt.date.today().isoformat()
        _wj("license.json", {"start": s, "sig": _sig(s)}); _reg(f"{s}|{_sig(s)}")
        return s

    def trial_left(): return TRIAL_DAYS - (dt.date.today() - dt.date.fromisoformat(trial_start())).days

    def _h(pw, salt): return hashlib.pbkdf2_hmac("sha256", pw.encode(), bytes.fromhex(salt), 200_000).hex()
    def has_users(): return bool(_rj("users.json", {}))

    def register(name, email, pw):
        email = email.strip().lower()
        if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email): raise ValueError("Enter a valid email address.")
        if len(pw) < 6: raise ValueError("Password must be at least 6 characters.")
        u = _rj("users.json", {})
        if email in u: raise ValueError("An account with this email already exists.")
        salt = secrets.token_hex(16)
        u[email] = dict(name=name.strip() or email.split("@")[0], salt=salt, hash=_h(pw, salt), created=dt.date.today().isoformat())
        _wj("users.json", u); trial_start(); return email

    def login(email, pw):
        email = email.strip().lower(); r = _rj("users.json", {}).get(email)
        if not r or not hmac.compare_digest(r["hash"], _h(pw, r["salt"])): raise ValueError("Incorrect email or password.")
        return email

    def name(email): return _rj("users.json", {}).get(email, {}).get("name", email)

    KINDS = {".pdf": "PDF", ".docx": "Word", ".xlsx": "Excel", ".pptx": "PowerPoint", ".jpg": "JPG", ".jpeg": "JPG", ".png": "PNG"}

    def record(email, tool, paths):
        s = _rj("stats.json", {}); u = s.setdefault(email, {"types": {}, "log": []})
        for p in paths:
            k = KINDS.get(Path(p).suffix.lower(), Path(p).suffix.strip(".").upper() or "FILE")
            u["types"][k] = u["types"].get(k, 0) + 1
        u["log"].insert(0, [dt.datetime.now().strftime("%d %b %Y, %H:%M"), tool, len(paths)]); u["log"] = u["log"][:50]
        _wj("stats.json", s)

    def stats(email): return _rj("stats.json", {}).get(email, {"types": {}, "log": []})

    return types.SimpleNamespace(TRIAL_DAYS=TRIAL_DAYS, trial_left=trial_left, trial_start=trial_start, has_users=has_users, register=register, login=login, name=name, record=record, stats=stats)

store = _make_store()


# ───────────────────────── UI ─────────────────────────
class Flow(tk.Frame):
    def __init__(s, m, cw):
        super().__init__(m, bg=BG); s.cw, s.ws = cw, []
        s.c = tk.Canvas(s, bg=BG, highlightthickness=0); sb = ttk.Scrollbar(s, command=s.c.yview); s.c.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y"); s.c.pack(side="left", fill="both", expand=True)
        s.inner = tk.Frame(s.c, bg=BG); s.w = s.c.create_window((0, 0), window=s.inner, anchor="nw")
        s.c.bind("<Configure>", s.lay); s.inner.bind("<Configure>", lambda e: s.c.configure(scrollregion=s.c.bbox("all")))
        s.c.bind_all("<MouseWheel>", lambda e: s.winfo_exists() and s.c.yview_scroll(int(-e.delta / 120), "units"))
    def set(s, ws): s.ws = ws; s.lay()
    def lay(s, e=None):
        w = s.c.winfo_width(); s.c.itemconfigure(s.w, width=w); cols = max(1, w // s.cw)
        for i, x in enumerate(s.ws): x.grid(row=i // cols, column=i % cols, padx=8, pady=8)


class ToolPage(tk.Frame):
    def __init__(s, app, t):
        super().__init__(app.body, bg=BG); s.app, s.t, s.files, s.v, s.rows, s.q = app, t, [], {}, [], queue.Queue()
        top = tk.Frame(s, bg="white"); top.pack(fill="x")
        tk.Button(top, text="← All tools", command=app.home, relief="flat", bg="white", fg=RED, font=FB).pack(side="left", padx=12, pady=8)
        tk.Label(top, text=f"{t['i']}  {t['t']}", font=("Segoe UI", 15, "bold"), bg="white").pack(side="left")
        tk.Label(s, text=t["d"], bg=BG, fg="#555", font=FN).pack(anchor="w", padx=20, pady=(8, 0))
        mid = tk.Frame(s, bg=BG); mid.pack(fill="both", expand=True, padx=16, pady=8)
        L = tk.LabelFrame(mid, text=" Files ", bg=BG, font=FN); L.pack(side="left", fill="both", expand=True)
        s.lb = tk.Listbox(L, selectmode="extended", font=FN, activestyle="none"); s.lb.pack(side="left", fill="both", expand=True, padx=6, pady=6)
        bt = tk.Frame(L, bg=BG); bt.pack(side="right", fill="y", padx=6)
        for txt, cmd in (("＋ Add files", s.add), ("Remove", s.rem), ("▲ Up", lambda: s.mv(-1)), ("▼ Down", lambda: s.mv(1)), ("Clear", s.clr)):
            tk.Button(bt, text=txt, command=cmd, width=12).pack(pady=3)
        R = tk.LabelFrame(mid, text=" Options ", bg=BG, font=FN, width=320); R.pack(side="right", fill="y", padx=(12, 0)); s.build(R)
        bot = tk.Frame(s, bg=BG); bot.pack(fill="x", padx=16, pady=(0, 12))
        r1 = tk.Frame(bot, bg=BG); r1.pack(fill="x")
        tk.Label(r1, text="Save to:", bg=BG, font=FN).pack(side="left")
        tk.Entry(r1, textvariable=app.out, font=FN).pack(side="left", fill="x", expand=True, padx=6)
        tk.Button(r1, text="Browse…", command=s.pick_out).pack(side="left")
        s.btn = tk.Button(bot, text=t["t"].upper() + "  ➜", bg=RED, fg="white", font=("Segoe UI", 13, "bold"), relief="flat", padx=30, pady=8, command=s.run)
        s.btn.pack(pady=10)
        s.pb = ttk.Progressbar(bot, mode="indeterminate"); s.res = tk.Frame(bot, bg=BG); s.res.pack(fill="x")

    def build(s, R):
        for r, f in enumerate(s.t["f"]):
            fr = tk.Frame(R, bg=BG); fr.grid(row=r, column=0, sticky="ew", padx=8, pady=4); s.rows.append((fr, f))
            if f["t"] == "k":
                v = tk.BooleanVar(value=f["d"]); tk.Checkbutton(fr, text=f["l"], variable=v, bg=BG, font=FN).pack(anchor="w")
            else:
                tk.Label(fr, text=f["l"], bg=BG, font=FN).pack(anchor="w")
                if f["t"] == "c":
                    v = tk.StringVar(value=next(l for l, x in f["ch"] if x == f["d"]))
                    ttk.Combobox(fr, textvariable=v, values=[l for l, _ in f["ch"]], state="readonly", width=30).pack(fill="x")
                elif f["t"] == "col":
                    v = tk.StringVar(value=f["d"]); b = tk.Button(fr, text="Pick color", bg=f["d"], fg="white")
                    b.configure(command=lambda v=v, b=b: (lambda c: (v.set(c), b.configure(bg=c)) if c else None)(colorchooser.askcolor(v.get())[1])); b.pack(fill="x")
                elif f["t"] == "img":
                    v = tk.StringVar(); e = tk.Frame(fr, bg=BG); e.pack(fill="x"); tk.Entry(e, textvariable=v, width=22).pack(side="left", fill="x", expand=True)
                    tk.Button(e, text="…", command=lambda v=v: v.set(filedialog.askopenfilename(filetypes=TYPES["image"]) or v.get())).pack(side="left")
                else:
                    v = tk.StringVar(value=f["d"]); tk.Entry(fr, textvariable=v, show="*" if f["t"] == "p" else "", width=34).pack(fill="x")
            s.v[f["k"]] = v
            if f["t"] == "c": v.trace_add("write", s.vis)
        s.vis()

    def get(s, f):
        v = s.v[f["k"]].get(); return dict(f["ch"]).get(v, v) if f["t"] == "c" else v
    def vis(s, *_):
        cur = {f["k"]: s.get(f) for _, f in s.rows}
        for fr, f in s.rows:
            if f["sh"]: (fr.grid() if cur[f["sh"][0]] == f["sh"][1] else fr.grid_remove())

    def fill(s): s.lb.delete(0, "end"); [s.lb.insert("end", f"{Path(f).name}    —    {Path(f).parent}") for f in s.files]
    def add(s):
        fs = filedialog.askopenfilenames(filetypes=TYPES[s.t["k"]])
        if fs: s.files = list(fs) if s.t["mx"] == 1 else s.files + list(fs); s.fill()
    def rem(s): [s.files.pop(i) for i in reversed(s.lb.curselection())]; s.fill()
    def clr(s): s.files = []; s.fill()
    def mv(s, d):
        sel = s.lb.curselection()
        if sel and 0 <= sel[0] + d < len(s.files):
            i = sel[0]; s.files[i], s.files[i + d] = s.files[i + d], s.files[i]; s.fill(); s.lb.selection_set(i + d)
    def pick_out(s):
        d = filedialog.askdirectory()
        if d: s.app.out.set(d)

    def run(s):
        t = s.t
        if store.trial_left() <= 0: return messagebox.showwarning("PDF Studio", "Your 90-day free trial has ended. Please contact the vendor to activate a licence.")
        if len(s.files) < t["m"]: return messagebox.showwarning("PDF Studio", f"Please add at least {t['m']} file(s).")
        files = s.files[:t["mx"]] if t["mx"] else list(s.files); o = {f["k"]: s.get(f) for _, f in s.rows}
        out = Path(s.app.out.get()); out.mkdir(parents=True, exist_ok=True)
        s.btn.configure(state="disabled", text="Working…"); s.pb.pack(fill="x"); s.pb.start(12)
        for w in s.res.winfo_children(): w.destroy()
        def work():
            try: s.q.put((t["fn"](files, o, out), None))
            except Exception as e: s.q.put((None, e))
        threading.Thread(target=work, daemon=True).start(); s.poll()

    def poll(s):
        try: res, err = s.q.get_nowait()
        except queue.Empty: return s.after(150, s.poll)
        s.pb.stop(); s.pb.pack_forget(); s.btn.configure(state="normal", text=s.t["t"].upper() + "  ➜")
        if err: return messagebox.showerror("PDF Studio", str(err))
        store.record(s.app.user, s.t["t"], res)
        tk.Label(s.res, text=f"✅ Done! {len(res)} file(s) created.", fg="#1a7f37", bg=BG, font=FB).pack(side="left", padx=6)
        tk.Button(s.res, text="Open file", command=lambda: reveal(res[0])).pack(side="left", padx=4)
        tk.Button(s.res, text="Open folder", command=lambda: reveal(Path(res[0]).parent)).pack(side="left")


ACC, DARK = "#4f46e5", "#111827"
PAL = {"PDF": "#ef4444", "Word": "#2563eb", "Excel": "#16a34a", "JPG": "#f59e0b", "PNG": "#8b5cf6", "PowerPoint": "#ea580c"}
BORDER = dict(highlightbackground="#e5e7eb", highlightthickness=1)


class App(tk.Tk):
    def __init__(s):
        super().__init__(); s.title("PDF Studio"); s.geometry("1200x780"); s.minsize(1000, 660)
        s.out = tk.StringVar(value=str(Path.home() / "Documents" / "PDF Studio Output")); s.user = None; s.cat = "All"
        s.login_screen("login" if store.has_users() else "reg")

    def clear_all(s): [c.destroy() for c in s.winfo_children()]
    def clr_main(s): [c.destroy() for c in s.main.winfo_children()]

    # ── authentication ──
    def login_screen(s, mode):
        s.clear_all(); s.configure(bg=DARK)
        card = tk.Frame(s, bg="white", padx=44, pady=34); card.place(relx=.5, rely=.5, anchor="center")
        tk.Label(card, text="❤ PDF Studio", font=("Segoe UI", 24, "bold"), fg=ACC, bg="white").pack()
        tk.Label(card, text="Sign in to your workspace" if mode == "login" else "Create your account · 90-day free trial",
                 bg="white", fg="#6b7280", font=FN).pack(pady=(2, 18))
        v = {k: tk.StringVar() for k in ("name", "email", "pw")}
        for lab, k, pw in [("Full name", "name", False)] * (mode == "reg") + [("Email", "email", False), ("Password", "pw", True)]:
            tk.Label(card, text=lab, bg="white", anchor="w", fg="#374151", font=FN).pack(fill="x")
            tk.Entry(card, textvariable=v[k], show="*" if pw else "", width=34, font=("Segoe UI", 11), relief="solid", bd=1).pack(fill="x", pady=(2, 12), ipady=6)
        def go(*_):
            try: e = store.register(v["name"].get(), v["email"].get(), v["pw"].get()) if mode == "reg" else store.login(v["email"].get(), v["pw"].get())
            except Exception as ex: return messagebox.showerror("PDF Studio", str(ex))
            s.user = e; s.shell()
        tk.Button(card, text="Create account" if mode == "reg" else "Sign in", command=go, bg=ACC, fg="white", relief="flat",
                  font=("Segoe UI", 11, "bold"), pady=9, cursor="hand2").pack(fill="x", pady=6)
        s.bind("<Return>", go)
        tk.Button(card, text="Already have an account? Sign in" if mode == "reg" else "New here? Create an account", relief="flat", bg="white", fg=ACC,
                  cursor="hand2", command=lambda: s.login_screen("login" if mode == "reg" else "reg")).pack(pady=(6, 0))

    def logout(s): s.user = None; s.login_screen("login")

    # ── main shell ──
    def shell(s):
        s.clear_all(); s.configure(bg=BG); s.unbind("<Return>")
        sb = tk.Frame(s, bg=DARK, width=250); sb.pack(side="left", fill="y"); sb.pack_propagate(False)
        tk.Label(sb, text="❤ PDF Studio", font=("Segoe UI", 17, "bold"), fg="white", bg=DARK).pack(anchor="w", padx=20, pady=(22, 18))
        s.nav = {}
        for key, lab in [("dash", "📊  Dashboard"), ("All", "🧰  All tools")] + [(c, "        " + c) for c in CATS]:
            b = tk.Button(sb, text=lab, anchor="w", relief="flat", bd=0, bg=DARK, fg="#cbd5e1", activebackground="#1f2937", activeforeground="white",
                          font=("Segoe UI", 10), padx=18, pady=8, cursor="hand2", command=lambda k=key: s.nav_to(k)); b.pack(fill="x"); s.nav[key] = b
        left = store.trial_left(); col = "#22c55e" if left > 14 else "#f59e0b" if left > 0 else "#ef4444"
        bx = tk.Frame(sb, bg=DARK); bx.pack(side="bottom", fill="x", padx=18, pady=16)
        tk.Label(bx, text=store.name(s.user), fg="white", bg=DARK, font=FB, anchor="w").pack(fill="x")
        tk.Label(bx, text=s.user, fg="#94a3b8", bg=DARK, font=("Segoe UI", 8), anchor="w").pack(fill="x")
        tk.Label(bx, text=f"Free trial: {left} days left" if left > 0 else "Trial expired", fg=col, bg=DARK, font=FB, anchor="w").pack(fill="x", pady=(12, 4))
        bar = tk.Frame(bx, bg="#374151", height=6); bar.pack(fill="x")
        tk.Frame(bar, bg=col).place(relx=0, rely=0, relheight=1, relwidth=max(0, min(1, left / store.TRIAL_DAYS)))
        tk.Button(bx, text="Sign out", command=s.logout, relief="flat", bg="#1f2937", fg="white", cursor="hand2", pady=5).pack(fill="x", pady=(12, 0))
        s.main = s.body = tk.Frame(s, bg=BG); s.main.pack(side="left", fill="both", expand=True)
        s.nav_to("dash")

    def nav_to(s, k):
        for n, b in s.nav.items(): b.configure(bg=ACC if n == k else DARK, fg="white" if n == k else "#cbd5e1")
        if k == "dash": s.dashboard()
        else: s.cat = k; s.tools_view()

    def home(s): s.nav_to(s.cat)

    def head(s, title, sub):
        tk.Label(s.main, text=title, font=("Segoe UI", 21, "bold"), bg=BG, fg=DARK).pack(anchor="w", padx=28, pady=(24, 0))
        tk.Label(s.main, text=sub, bg=BG, fg="#6b7280", font=FN).pack(anchor="w", padx=28)

    # ── dashboard ──
    def dashboard(s):
        s.clr_main(); m = s.main; st = store.stats(s.user); ty = st["types"]; tot = sum(ty.values()); left = store.trial_left()
        s.head(f"Welcome back, {store.name(s.user).split()[0]} 👋", "Here's an overview of the files you've converted with PDF Studio.")
        row = tk.Frame(m, bg=BG); row.pack(fill="x", padx=22, pady=16)
        top = max(ty, key=ty.get) if ty else "—"
        for lab, val, col in (("Files converted", tot, ACC), ("Most produced type", top, "#0ea5e9"), ("File types", len(ty), "#16a34a"), ("Trial days left", max(left, 0), "#f59e0b")):
            c = tk.Frame(row, bg="white", padx=20, pady=16, **BORDER); c.pack(side="left", fill="x", expand=True, padx=6)
            tk.Label(c, text=str(val), font=("Segoe UI", 26, "bold"), fg=col, bg="white").pack(anchor="w")
            tk.Label(c, text=lab, fg="#6b7280", bg="white", font=FN).pack(anchor="w")
        low = tk.Frame(m, bg=BG); low.pack(fill="both", expand=True, padx=22, pady=(0, 20))
        ch = tk.Frame(low, bg="white", **BORDER); ch.pack(side="left", fill="both", expand=True, padx=6)
        tk.Label(ch, text="Converted files by type", font=("Segoe UI", 12, "bold"), bg="white").pack(anchor="w", padx=18, pady=(16, 6))
        cv = tk.Canvas(ch, bg="white", highlightthickness=0); cv.pack(fill="both", expand=True, padx=18, pady=(0, 16))
        def draw(e=None):
            cv.delete("all"); w = cv.winfo_width()
            if not ty: return cv.create_text(w / 2, 70, text="No conversions yet — run a tool and your stats appear here.", fill="#9ca3af", font=FN)
            mx = max(ty.values())
            for i, (k, n) in enumerate(sorted(ty.items(), key=lambda x: -x[1])[:8]):
                y = 8 + i * 40; bw = max(6, (w - 170) * n / mx)
                cv.create_text(0, y + 13, text=k, anchor="w", font=FN, fill="#374151")
                cv.create_rectangle(95, y, 95 + bw, y + 26, fill=PAL.get(k, ACC), width=0)
                cv.create_text(103 + bw, y + 13, text=str(n), anchor="w", font=FB, fill=DARK)
        cv.bind("<Configure>", draw)
        rc = tk.Frame(low, bg="white", width=340, **BORDER); rc.pack(side="right", fill="y", padx=6); rc.pack_propagate(False)
        tk.Label(rc, text="Recent activity", font=("Segoe UI", 12, "bold"), bg="white").pack(anchor="w", padx=18, pady=(16, 6))
        for t, tool, n in st["log"][:8]:
            tk.Label(rc, text=f"{tool} · {n} file(s)\n{t}", justify="left", anchor="w", bg="white", fg="#374151", font=FN).pack(fill="x", padx=18, pady=5)
        if not st["log"]: tk.Label(rc, text="Nothing yet.", bg="white", fg="#9ca3af").pack(padx=18, anchor="w")

    # ── tools ──
    def tools_view(s):
        s.clr_main(); s.head("All tools" if s.cat == "All" else s.cat, "Pick a tool to get started. Everything runs locally on your PC.")
        fl = Flow(s.main, 270); fl.pack(fill="both", expand=True, padx=18, pady=12)
        fl.set([s.card(fl.inner, t) for t in TOOLS if s.cat in ("All", t["c"])])

    def card(s, parent, t):
        f = tk.Frame(parent, bg="white", width=250, height=140, cursor="hand2", **BORDER); f.pack_propagate(False)
        tk.Label(f, text=t["i"], font=("Segoe UI Emoji", 22), bg="white", anchor="w").pack(fill="x", padx=14, pady=(12, 0))
        tk.Label(f, text=t["t"], font=("Segoe UI", 11, "bold"), bg="white", fg=DARK, anchor="w").pack(fill="x", padx=14)
        tk.Label(f, text=t["d"], font=("Segoe UI", 9), fg="#6b7280", bg="white", wraplength=220, justify="left", anchor="nw").pack(fill="x", padx=14, pady=(2, 0))
        for w in (f, *f.winfo_children()):
            w.bind("<Button-1>", lambda e, t=t: s.open(t))
            w.bind("<Enter>", lambda e, f=f: f.configure(highlightbackground=ACC)); w.bind("<Leave>", lambda e, f=f: f.configure(highlightbackground="#e5e7eb"))
        return f

    def open(s, t): s.clr_main(); ToolPage(s, t).pack(fill="both", expand=True)


if __name__ == "__main__":
    try:
        import ctypes; ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception: pass
    App().mainloop()
