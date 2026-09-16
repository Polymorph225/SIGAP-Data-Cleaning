import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────
#  PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Data Cleaning SIGAP Bojonegoro",
    page_icon="🧹",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
#  CUSTOM CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
/* ═══════════════════════════════════════════════════════════
   Tema terang, senada dengan SIGAP-Bojonegoro.
   Token warna dan bayangan sengaja dibuat sama persis supaya
   kedua aplikasi terasa satu keluarga.
   ═══════════════════════════════════════════════════════════ */
:root {
    --sigap-blue:      #1e5eff;
    --sigap-blue-deep: #1039a8;
    --sigap-cyan:      #06b6d4;
    --sigap-ink:       #0f172a;
    --sigap-muted:     #56637a;
    --sigap-line:      rgba(148,163,184,.30);
    --sigap-surface:   rgba(255,255,255,.74);

    --lift-1: 0 1px 2px rgba(15,23,42,.05), 0 2px 6px rgba(15,23,42,.06);
    --lift-2: 0 2px 4px rgba(15,23,42,.04), 0 10px 24px rgba(30,94,255,.12);
    --lift-3: 0 12px 28px rgba(30,94,255,.18), 0 28px 56px rgba(15,23,42,.12);
}

/* ── Latar: gradien lembut + noda cahaya ─────────────────── */
.stApp {
    background:
        radial-gradient(900px 520px at 12% -8%,  rgba(30,94,255,.13), transparent 60%),
        radial-gradient(760px 460px at 92% 4%,   rgba(6,182,212,.12), transparent 62%),
        linear-gradient(180deg, #f5f8ff 0%, #eef3fb 46%, #f7f9fc 100%);
    background-attachment: fixed;
}
.block-container { padding: 1.25rem 2rem 3.5rem 2rem; max-width: 1500px; }

h1, h2, h3 { color: var(--sigap-ink); letter-spacing: -.015em; }
h1 { font-weight: 800 !important; }
h2, h3 { font-weight: 700 !important; }

/* ── Sidebar ─────────────────────────────────────────────── */
section[data-testid="stSidebar"] {
    background: linear-gradient(185deg, #ffffff 0%, #f3f7ff 100%);
    border-right: 1px solid var(--sigap-line);
    box-shadow: 6px 0 26px rgba(15,23,42,.06);
}
section[data-testid="stSidebar"] * { color: var(--sigap-ink); }
section[data-testid="stSidebar"] .stCaption,
section[data-testid="stSidebar"] small { color: var(--sigap-muted) !important; }

/* ── Kepala halaman ──────────────────────────────────────── */
.main-header {
    position: relative;
    overflow: hidden;
    border-radius: 20px;
    padding: 1.6rem 2rem;
    margin-bottom: 1.4rem;
    background:
        radial-gradient(620px 280px at 84% 44%, rgba(6,182,212,.30), transparent 66%),
        linear-gradient(126deg, #0b2f8f 0%, #1e5eff 46%, #2aa9d9 100%);
    box-shadow: 0 14px 32px rgba(16,57,168,.28), 0 34px 64px rgba(15,23,42,.16);
}
.main-header::after {
    content: ""; position: absolute; inset: 0; border-radius: inherit; pointer-events: none;
    background: linear-gradient(168deg, rgba(255,255,255,.30), rgba(255,255,255,0) 38%);
}
.main-header h1 {
    color: #ffffff !important;
    font-size: 1.85rem; font-weight: 800; margin: 0 0 .35rem;
    letter-spacing: -.02em; position: relative; z-index: 1;
}
.main-header p {
    color: rgba(233,242,255,.94);
    font-size: .95rem; margin: 0; position: relative; z-index: 1;
}

/* ── Kartu metrik: kaca + terangkat ──────────────────────── */
.metric-row { display: flex; gap: .9rem; flex-wrap: wrap; margin: .2rem 0 1.1rem; }
.metric-card {
    position: relative;
    flex: 1 1 180px;
    padding: 1rem 1.15rem;
    border-radius: 16px;
    background: var(--sigap-surface);
    backdrop-filter: blur(14px) saturate(150%);
    -webkit-backdrop-filter: blur(14px) saturate(150%);
    border: 1px solid var(--sigap-line);
    box-shadow: var(--lift-2);
    transition: transform .28s cubic-bezier(.22,1,.36,1), box-shadow .28s ease;
}
.metric-card::before {
    content: ""; position: absolute; inset: 0; border-radius: inherit;
    padding: 1px;
    background: linear-gradient(160deg, rgba(255,255,255,.95), rgba(255,255,255,0) 42%);
    -webkit-mask: linear-gradient(#000 0 0) content-box, linear-gradient(#000 0 0);
    -webkit-mask-composite: xor; mask-composite: exclude;
    pointer-events: none;
}
.metric-card:hover {
    transform: perspective(900px) translateY(-4px) rotateX(4deg);
    box-shadow: var(--lift-3);
}
.metric-card .label {
    color: var(--sigap-muted);
    font-size: .72rem; font-weight: 700;
    letter-spacing: .06em; text-transform: uppercase;
}
.metric-card .value {
    font-size: 1.85rem; font-weight: 800; line-height: 1.25; margin-top: .25rem;
    background: linear-gradient(120deg, var(--sigap-blue-deep), var(--sigap-cyan));
    -webkit-background-clip: text; background-clip: text;
    -webkit-text-fill-color: transparent;
}
.metric-card .value.good {
    background: linear-gradient(120deg, #0f766e, #10b981);
    -webkit-background-clip: text; background-clip: text;
    -webkit-text-fill-color: transparent;
}
.metric-card .value.bad {
    background: linear-gradient(120deg, #b45309, #f59e0b);
    -webkit-background-clip: text; background-clip: text;
    -webkit-text-fill-color: transparent;
}

/* ── Judul bagian ───────────────────────────────────────── */
.section-title {
    display: inline-block;
    color: var(--sigap-blue-deep);
    font-size: .8rem; font-weight: 800;
    letter-spacing: .06em; text-transform: uppercase;
    padding: .34rem .8rem;
    margin: .2rem 0 .7rem;
    border-radius: 999px;
    background: linear-gradient(145deg, #e9f0ff, #d8e6ff);
    border: 1px solid rgba(255,255,255,.7);
    box-shadow: var(--lift-1);
}

/* ── Baris catatan pembersihan ───────────────────────────── */
.log-item {
    display: flex; align-items: flex-start; gap: .6rem;
    padding: .55rem .7rem;
    border-radius: 10px;
    background: rgba(255,255,255,.62);
    border: 1px solid var(--sigap-line);
    color: var(--sigap-ink);
    font-size: .86rem;
    margin-bottom: .45rem;
}
.log-icon { flex-shrink: 0; font-size: 1rem; line-height: 1.35; }

/* ── Lencana ─────────────────────────────────────────────── */
.badge {
    display: inline-block;
    padding: .26rem .8rem;
    border-radius: 999px;
    font-size: .74rem; font-weight: 700;
    margin-right: .4rem;
    border: 1px solid rgba(255,255,255,.6);
    box-shadow: var(--lift-1);
}
.badge-blue   { background:linear-gradient(145deg,#e6efff,#cfe0ff); color:#1746b8; }
.badge-yellow { background:linear-gradient(145deg,#fff5da,#ffe9b4); color:#8a5a06; }
.badge-red    { background:linear-gradient(145deg,#ffe3e3,#ffc9c9); color:#a11a1a; }
.badge-green  { background:linear-gradient(145deg,#e4fbec,#c9f3da); color:#12693c; }

/* ── Petunjuk unggah ─────────────────────────────────────── */
.upload-hint {
    text-align: center;
    padding: 2.4rem 1.5rem;
    border-radius: 18px;
    background: linear-gradient(150deg, rgba(30,94,255,.08), rgba(6,182,212,.06));
    border: 1px dashed rgba(30,94,255,.34);
    box-shadow: var(--lift-1);
    margin-bottom: 1.2rem;
}

/* ── Panel bawaan Streamlit ──────────────────────────────── */
div[data-testid="stExpander"],
div[data-testid="stAlert"],
div[data-testid="stDataFrame"],
div[data-testid="stTable"] {
    border-radius: 14px !important;
    border: 1px solid var(--sigap-line) !important;
    box-shadow: var(--lift-1);
    overflow: hidden;
}
div[data-testid="stExpander"] { background: var(--sigap-surface); }

/* ── Tombol: timbul, menekan saat diklik ─────────────────── */
.stButton > button, .stDownloadButton > button, .stLinkButton > a {
    border-radius: 12px;
    font-weight: 650;
    border: 1px solid rgba(30,94,255,.28);
    box-shadow: var(--lift-1);
    transition: transform .16s ease, box-shadow .16s ease;
}
.stButton > button:hover, .stDownloadButton > button:hover, .stLinkButton > a:hover {
    transform: translateY(-2px);
    box-shadow: var(--lift-2);
}
.stButton > button:active, .stDownloadButton > button:active {
    transform: translateY(0);
    box-shadow: inset 0 2px 5px rgba(15,23,42,.16);
}
.stButton > button[kind="primary"],
.stDownloadButton > button[kind="primary"] {
    background: linear-gradient(135deg, var(--sigap-blue), var(--sigap-cyan));
    border: none; color: #fff;
}

hr { border-color: var(--sigap-line); }

@media (prefers-reduced-motion: reduce) {
    * { transition: none !important; animation: none !important; }
}
@media (max-width: 640px) {
    .metric-card:hover { transform: none; }
    .block-container { padding: 1rem 1rem 2.5rem 1rem; }
}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────

def detect_issues(df: pd.DataFrame) -> dict:
    """Detect data quality issues and return a report dict."""
    issues = {}

    # Trailing comma/space patterns
    trailing_comma_cols = []
    for col in df.select_dtypes(include="object").columns:
        if df[col].dropna().astype(str).str.contains(r'\s*,\s*$', regex=True).any():
            trailing_comma_cols.append(col)
    if trailing_comma_cols:
        issues["trailing_comma"] = trailing_comma_cols

    # Duplicates
    dup_count = df.duplicated().sum()
    if dup_count:
        issues["duplicates"] = int(dup_count)

    # Missing values
    nulls = df.isnull().sum()
    null_cols = nulls[nulls > 0].to_dict()
    if null_cols:
        issues["nulls"] = null_cols

    # Placeholder values
    placeholder_cols = []
    for col in df.select_dtypes(include="object").columns:
        if df[col].dropna().astype(str).str.fullmatch(r'[-=]+').any():
            placeholder_cols.append(col)
    if placeholder_cols:
        issues["placeholders"] = placeholder_cols

    # Date columns
    date_cols = []
    for col in df.columns:
        if "tgl" in col.lower() or "tanggal" in col.lower() or "date" in col.lower() or "lahir" in col.lower():
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                date_cols.append(col)
    if date_cols:
        issues["date_cols"] = date_cols

    return issues


def clean_dataframe(df: pd.DataFrame, options: dict) -> tuple[pd.DataFrame, list]:
    """Apply cleaning steps based on options. Returns (cleaned_df, log)."""
    df = df.copy()
    log = []

    # 1. Remove duplicates
    if options.get("remove_duplicates"):
        before = len(df)
        df.drop_duplicates(inplace=True)
        removed = before - len(df)
        if removed:
            log.append(("🗑️", f"Hapus {removed} baris duplikat"))
        else:
            log.append(("✅", "Tidak ada baris duplikat"))

    # 2. Strip trailing commas/spaces from all object cols
    if options.get("strip_trailing_comma"):
        affected = []
        for col in df.select_dtypes(include="object").columns:
            mask = df[col].dropna().astype(str).str.contains(r'\s*,\s*$', regex=True)
            if mask.any():
                df[col] = df[col].astype(str).str.replace(r'\s*,\s*$', '', regex=True).str.strip()
                affected.append(col)
        if affected:
            log.append(("✂️", f"Hapus trailing koma pada: {', '.join(affected)}"))

    # 3. Strip extra internal spaces
    if options.get("strip_spaces"):
        for col in options.get("strip_spaces_cols", []):
            if col in df.columns:
                df[col] = df[col].astype(str).str.replace(r'\s+', '', regex=True)
        log.append(("🔤", "Hapus spasi berlebih di dalam NIK / No Penjamin"))

    # 4. Format dates
    if options.get("format_dates"):
        fmt = options.get("date_format", "%d/%m/%Y")
        for col in df.columns:
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = pd.to_datetime(df[col]).dt.strftime(fmt)
                log.append(("📅", f"Format tanggal kolom '{col}' → {fmt}"))

    # 5. Standardize placeholder Keterangan
    if options.get("standardize_keterangan") and "Keterangan" in df.columns:
        placeholder_val = options.get("placeholder_replacement", "Tidak Ada Keterangan")
        df["Keterangan"] = df["Keterangan"].astype(str).str.strip()
        df["Keterangan"] = df["Keterangan"].replace(["-", "=", "nan", ""], placeholder_val)
        df["Keterangan"] = df["Keterangan"].apply(
            lambda x: placeholder_val if re.fullmatch(r'[-=\s]+', str(x)) else x
        )
        if options.get("uppercase_keterangan"):
            df["Keterangan"] = df["Keterangan"].str.upper()
        log.append(("📝", f"Standarisasi Keterangan: '-', '=' → '{placeholder_val}'"))

    # 6. Uppercase Nama
    if options.get("uppercase_nama") and "Nama" in df.columns:
        df["Nama"] = df["Nama"].str.upper().str.strip()
        log.append(("🔠", "Nama diubah ke UPPERCASE"))

    # 7. Title case Desa
    if options.get("titlecase_desa") and "Desa" in df.columns:
        df["Desa"] = df["Desa"].str.strip().str.title()
        log.append(("🏘️", "Desa diubah ke Title Case"))

    # 8. Fill nulls
    fill_map = options.get("fill_nulls", {})
    for col, fill_val in fill_map.items():
        if col in df.columns and fill_val:
            n = int(df[col].isnull().sum())
            if n:
                df[col] = df[col].fillna(fill_val)
                log.append(("📋", f"Isi {n} nilai kosong di '{col}' dengan '{fill_val}'"))

    # 9. Rename columns
    rename_map = options.get("rename_cols", {})
    if rename_map:
        df.rename(columns=rename_map, inplace=True)
        for old, new in rename_map.items():
            log.append(("✏️", f"Rename kolom '{old}' → '{new}'"))

    # 10. Drop selected columns
    drop_cols = [c for c in options.get("drop_cols", []) if c in df.columns]
    if drop_cols:
        df.drop(columns=drop_cols, inplace=True)
        log.append(("❌", f"Hapus kolom: {', '.join(drop_cols)}"))

    # 11. Sort
    sort_col = options.get("sort_by")
    if sort_col and sort_col in df.columns:
        df.sort_values(sort_col, inplace=True)
        df.reset_index(drop=True, inplace=True)
        log.append(("🔃", f"Urutkan berdasarkan '{sort_col}'"))

    return df, log


def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Data") -> bytes:
    """Export cleaned DataFrame to a nicely formatted Excel file."""
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name[:31]

    h_fill  = PatternFill('solid', start_color='1F4E79')
    h_font  = Font(bold=True, color='FFFFFF', name='Calibri', size=10)
    h_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    alt_fill = PatternFill('solid', start_color='EBF3FB')
    d_font   = Font(name='Calibri', size=9)
    d_align  = Alignment(vertical='center')
    thin = Side(style='thin', color='B0C4DE')
    bdr  = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Header row
    for ci, col in enumerate(df.columns, 1):
        c = ws.cell(row=1, column=ci, value=col)
        c.font = h_font; c.fill = h_fill
        c.alignment = h_align; c.border = bdr

    # Data rows
    df_str = df.fillna('-')
    for ri, row in enumerate(df_str.itertuples(index=False), 2):
        fill = alt_fill if ri % 2 == 0 else PatternFill('solid', start_color='FFFFFF')
        for ci, val in enumerate(row, 1):
            c = ws.cell(row=ri, column=ci, value=val)
            c.font = d_font; c.fill = fill
            c.alignment = d_align; c.border = bdr

    # Auto column widths (capped)
    for ci, col in enumerate(df.columns, 1):
        max_len = max(
            len(str(col)),
            df_str.iloc[:, ci-1].astype(str).str.len().max() if len(df_str) else 0
        )
        ws.column_dimensions[get_column_letter(ci)].width = min(max_len + 3, 45)

    ws.freeze_panes = 'A2'
    ws.row_dimensions[1].height = 32

    # Summary sheet
    ws2 = wb.create_sheet('Info')
    summary = [
        ['File diekspor oleh RME Data Cleaner'],
        ['Total baris:', len(df)],
        ['Total kolom:', len(df.columns)],
    ]
    for ri, row in enumerate(summary, 1):
        for ci, val in enumerate(row, 1):
            ws2.cell(row=ri, column=ci, value=val)
    ws2.column_dimensions['A'].width = 30
    ws2.column_dimensions['B'].width = 15

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()



# ─────────────────────────────────────────────
#  BANNER 3D — sama seperti SIGAP-Bojonegoro
# ─────────────────────────────────────────────
HERO_HEIGHT = 230

HERO_HTML = """
<div id="hero">
  <canvas id="globe"></canvas>
  <div id="copy">
    <div id="eyebrow">PEMERINTAH KABUPATEN BOJONEGORO &middot; UPT PUSKESMAS PURWOSARI</div>
    <h1>Data Cleaning SIGAP</h1>
    <p>Membersihkan dan menggabungkan berkas RME bulanan menjadi satu berkas siap unggah ke SIGAP&#8209;Bojonegoro</p>
  </div>
</div>

<style>
  html, body { margin:0; padding:0; background:transparent; overflow:hidden; }
  #hero {
    position: relative;
    height: 250px;
    border-radius: 20px;
    overflow: hidden;
    background:
      radial-gradient(680px 300px at 82% 48%, rgba(6,182,212,.30), transparent 66%),
      linear-gradient(126deg, #0b2f8f 0%, #1e5eff 46%, #2aa9d9 100%);
    box-shadow: 0 14px 32px rgba(16,57,168,.28), 0 34px 64px rgba(15,23,42,.16);
    font-family: "Source Sans Pro", -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
  }
  /* kilau tipis di tepi atas, biar permukaannya terasa melengkung */
  #hero::after {
    content:""; position:absolute; inset:0; border-radius:inherit; pointer-events:none;
    background: linear-gradient(168deg, rgba(255,255,255,.30), rgba(255,255,255,0) 38%);
  }
  #globe { position:absolute; inset:0; width:100%; height:100%; display:block; }
  #copy {
    position: relative; z-index: 2;
    padding: 2.1rem 2.3rem;
    max-width: 58%;
    color: #fff;
  }
  #eyebrow {
    font-size: .68rem; font-weight: 700; letter-spacing: .14em;
    color: rgba(255,255,255,.80); margin-bottom: .55rem;
  }
  #copy h1 {
    margin: 0 0 .45rem 0;
    color: #fff;
    font-size: clamp(1.7rem, 3.6vw, 2.6rem);
    font-weight: 800; letter-spacing: -.02em; line-height: 1.08;
    text-shadow: 0 2px 18px rgba(4,22,74,.45);
  }
  #copy p {
    margin: 0; font-size: .95rem; line-height: 1.5;
    color: rgba(255,255,255,.90); max-width: 30rem;
    text-shadow: 0 1px 10px rgba(4,22,74,.35);
  }
  @media (max-width: 720px) {
    #copy { max-width: 100%; padding: 1.5rem 1.5rem; }
    #copy p { font-size: .86rem; }
  }
</style>

<script src="https://cdnjs.cloudflare.com/ajax/libs/three.js/r128/three.min.js"></script>
<script>
(function () {
  var cv = document.getElementById('globe');
  // three.js tidak termuat (CDN diblokir) -> banner tetap tampil, hanya tanpa globe
  if (typeof THREE === 'undefined' || !cv) { if (cv) cv.style.display = 'none'; return; }

  var reduce = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;

  var scene  = new THREE.Scene();
  var camera = new THREE.PerspectiveCamera(42, 1, 0.1, 100);
  camera.position.z = 5.2;

  var renderer = new THREE.WebGLRenderer({ canvas: cv, alpha: true, antialias: true });
  renderer.setClearColor(0x000000, 0);
  renderer.setPixelRatio(Math.min(window.devicePixelRatio || 1, 2));

  var world = new THREE.Group();
  // digeser ke kanan supaya tidak menabrak teks
  world.position.x = 1.55;
  scene.add(world);

  var R = 1.62;

  // rangka bola
  var wire = new THREE.Mesh(
    new THREE.IcosahedronGeometry(R, 2),
    new THREE.MeshBasicMaterial({ color: 0xbfe4ff, wireframe: true, transparent: true, opacity: 0.44 })
  );
  world.add(wire);

  // bola inti semu, memberi kesan padat
  var core = new THREE.Mesh(
    new THREE.SphereGeometry(R * 0.985, 32, 32),
    new THREE.MeshBasicMaterial({ color: 0x0a2a7a, transparent: true, opacity: 0.42 })
  );
  world.add(core);

  // titik-titik "desa" tersebar di permukaan (distribusi spiral Fibonacci)
  var N = 170, pos = new Float32Array(N * 3), gold = Math.PI * (3 - Math.sqrt(5));
  for (var i = 0; i < N; i++) {
    var y = 1 - (i / (N - 1)) * 2;
    var r = Math.sqrt(Math.max(0, 1 - y * y));
    var th = gold * i;
    pos[i*3]   = Math.cos(th) * r * R * 1.012;
    pos[i*3+1] = y * R * 1.012;
    pos[i*3+2] = Math.sin(th) * r * R * 1.012;
  }
  var pg = new THREE.BufferGeometry();
  pg.setAttribute('position', new THREE.BufferAttribute(pos, 3));
  world.add(new THREE.Points(pg, new THREE.PointsMaterial({
    color: 0xaef6ff, size: 0.056, transparent: true, opacity: 1.0, sizeAttenuation: true
  })));

  // dua cincin orbit miring
  [[0.62, 2.18], [-0.45, 2.52]].forEach(function (o) {
    var ring = new THREE.Mesh(
      new THREE.TorusGeometry(o[1], 0.006, 8, 128),
      new THREE.MeshBasicMaterial({ color: 0x9fe9ff, transparent: true, opacity: 0.42 })
    );
    ring.rotation.x = Math.PI / 2 + o[0];
    ring.rotation.y = o[0] * 0.5;
    world.add(ring);
  });

  function resize() {
    var w = cv.clientWidth, h = cv.clientHeight;
    if (!w || !h) return;
    renderer.setSize(w, h, false);
    camera.aspect = w / h;
    // di layar sempit globe dikecilkan & digeser supaya teks tetap terbaca
    var narrow = w < 720;
    world.position.x = narrow ? 0.85 : 1.55;
    world.scale.setScalar(narrow ? 0.72 : 1);
    camera.updateProjectionMatrix();
  }
  resize();
  window.addEventListener('resize', resize);

  // parallax halus mengikuti kursor
  var tx = 0, ty = 0;
  document.addEventListener('mousemove', function (e) {
    tx = (e.clientX / window.innerWidth  - 0.5) * 0.34;
    ty = (e.clientY / window.innerHeight - 0.5) * 0.22;
  });

  var visible = true;
  document.addEventListener('visibilitychange', function () { visible = !document.hidden; });

  var t = 0;
  function frame() {
    requestAnimationFrame(frame);
    if (!visible) return;                 // berhenti menggambar saat tab tidak aktif
    t += reduce ? 0 : 0.0024;             // hormati setelan "kurangi animasi"
    world.rotation.y = t * 2.6;
    world.rotation.x = -0.24 + Math.sin(t * 1.7) * 0.05;
    world.rotation.y += (tx - world.rotation.y % (Math.PI * 2)) * 0;
    camera.position.x = tx;
    camera.position.y = -ty;
    camera.lookAt(world.position.x * 0.45, 0, 0);
    renderer.render(scene, camera);
  }
  frame();
})();
</script>
"""

def render_hero_3d():
    """Banner header dengan globe 3D berputar, kembaran banner SIGAP-Bojonegoro.

    Bila three.js gagal dimuat — misalnya jaringan puskesmas memblokir CDN —
    banner tetap tampil rapi dengan gradiennya saja, teksnya tidak pernah hilang.
    """
    if hasattr(st, "iframe"):
        st.iframe(HERO_HTML, height=HERO_HEIGHT)
    else:
        components.html(HERO_HTML, height=HERO_HEIGHT, scrolling=False)


# ─────────────────────────────────────────────
#  PEMBACAAN & PENGGABUNGAN BERKAS
# ─────────────────────────────────────────────

# Pemetaan kolom ekspor RME Sintesa Emas ke kolom masukan SIGAP-Bojonegoro.
KOLOM_SIGAP = {
    "Tgl":        "tanggal_kunjungan",
    "No RM":      "no_rm",
    "Usia":       "umur",
    "L/LP":       "jenis_kelamin",
    "Unit":       "poli",
    "Diagnosis":  "diagnosa",
    "Cara Bayar": "pembiayaan",
    "Desa":       "desa",
}


def _baca_teks_berpemisah(data: bytes) -> pd.DataFrame:
    """Ekspor RME kerap berekstensi .xls padahal isinya teks berpemisah tab.

    Fungsi ini menebak pemisahnya dari baris judul, lalu membaca seluruh kolom
    sebagai teks supaya nomor rekam medis dan NIK tidak berubah jadi angka.
    """
    teks = None
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            teks = data.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if teks is None:
        raise ValueError("Penyandian karakter berkas tidak dikenali.")

    baris_judul = teks.split("\n", 1)[0]
    pemisah = max(("\t", ";", "|", ","), key=baris_judul.count)
    if baris_judul.count(pemisah) == 0:
        raise ValueError("Tidak ditemukan pemisah kolom pada baris judul.")

    return pd.read_csv(
        io.StringIO(teks), sep=pemisah, dtype=str,
        quotechar='"', engine="python", keep_default_na=True,
    )


def baca_berkas(nama: str, data: bytes) -> pd.DataFrame:
    """Baca satu berkas apa pun bentuknya: csv, xlsx, xls asli, atau xls palsu."""
    n = nama.lower()
    if n.endswith(".csv"):
        return _baca_teks_berpemisah(data)
    if n.endswith(".xlsx"):
        return pd.read_excel(io.BytesIO(data), dtype=str, engine="openpyxl")
    # .xls — bisa Excel lama yang sebenarnya, bisa pula teks berpemisah
    try:
        return pd.read_excel(io.BytesIO(data), dtype=str, engine="xlrd")
    except Exception:
        return _baca_teks_berpemisah(data)


def _ke_tanggal(seri: pd.Series) -> pd.Series:
    """Ubah kolom tanggal menjadi datetime tanpa salah tebak hari/bulan.

    `dayfirst=True` berbahaya untuk tanggal ISO: pandas menebak formatnya dari
    nilai pertama, sehingga "2026-04-01" bisa terbaca 1 April atau 4 Januari,
    dan sisanya yang tanggalnya di atas 12 gagal diurai. Di sini setiap format
    yang mungkin dicoba satu per satu, lalu yang paling banyak berhasil dipakai.
    """
    teks = seri.astype(str).str.strip()
    formats = ["%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%d/%m/%Y", "%d-%m-%Y",
               "%Y/%m/%d", "%d/%m/%y", "%m/%d/%Y"]
    terbaik, jml_terbaik = None, -1
    for fmt in formats:
        t = pd.to_datetime(teks, errors="coerce", format=fmt)
        n = int(t.notna().sum())
        if n > jml_terbaik:
            terbaik, jml_terbaik = t, n
        if n == len(teks):
            return t
    # Tidak ada format baku yang cocok sepenuhnya — serahkan ke penebak pandas.
    bebas = pd.to_datetime(teks, errors="coerce")
    return bebas if int(bebas.notna().sum()) > jml_terbaik else terbaik


def cari_kolom_tanggal(df: pd.DataFrame):
    """Cari kolom tanggal kunjungan — bukan tanggal lahir."""
    for kandidat in ("Tgl", "Tanggal", "tanggal_kunjungan", "Tgl Kunjungan", "Tanggal Kunjungan"):
        if kandidat in df.columns:
            return kandidat
    for c in df.columns:
        nama = str(c)
        if re.search(r"tgl|tanggal|date", nama, re.I) and not re.search(r"lahir|birth|daftar", nama, re.I):
            return c
    return None


@st.cache_data(show_spinner=False)
def gabung_berkas(berkas: list, buang_kembar: bool = True):
    """Gabungkan beberapa berkas menjadi satu, urut menurut tanggal kunjungan.

    `berkas` berupa daftar pasangan (nama, isi byte) supaya hasilnya bisa
    di-cache — objek unggahan Streamlit sendiri tidak bisa di-hash.
    Mengembalikan (df_gabungan, ringkasan_per_berkas, jumlah_kembar, nama_kolom_tanggal).
    """
    potongan, ringkasan = [], []

    for nama, data in berkas:
        df = baca_berkas(nama, data)
        df.columns = [str(c).strip() for c in df.columns]
        kol_tgl = cari_kolom_tanggal(df)
        tgl = (_ke_tanggal(df[kol_tgl]) if kol_tgl
               else pd.Series(pd.NaT, index=df.index))
        ringkasan.append({
            "Berkas": nama,
            "Baris": len(df),
            "Kolom": len(df.columns),
            "Tanggal awal": tgl.min().strftime("%d/%m/%Y") if tgl.notna().any() else "—",
            "Tanggal akhir": tgl.max().strftime("%d/%m/%Y") if tgl.notna().any() else "—",
            "Tanggal tak terbaca": int(tgl.isna().sum()),
        })
        df = df.copy()
        df["_urut_tanggal"] = tgl
        df["_asal_berkas"] = nama
        potongan.append(df)

    gab = pd.concat(potongan, ignore_index=True, sort=False)

    kolom_asli = [c for c in gab.columns if not c.startswith("_")]
    kembar = int(gab.duplicated(subset=kolom_asli).sum())
    if buang_kembar and kembar:
        gab = gab.drop_duplicates(subset=kolom_asli).reset_index(drop=True)

    # Urutkan menurut tanggal; baris tanpa tanggal ditaruh di akhir agar terlihat.
    gab = gab.sort_values("_urut_tanggal", kind="stable", na_position="last").reset_index(drop=True)

    kol_tgl_akhir = cari_kolom_tanggal(gab.drop(columns=["_urut_tanggal", "_asal_berkas"]))
    return gab, ringkasan, kembar, kol_tgl_akhir


def ke_format_sigap(df: pd.DataFrame):
    """Ubah data gabungan menjadi 8 kolom masukan SIGAP-Bojonegoro.

    Kolom identitas — nama, NIK, alamat, nomor penjamin — sengaja tidak ikut.
    Mengembalikan (df_sigap, daftar_kolom_yang_tidak_ditemukan).
    """
    ada = {asal: baru for asal, baru in KOLOM_SIGAP.items() if asal in df.columns}
    kurang = [asal for asal in KOLOM_SIGAP if asal not in df.columns]
    if not ada:
        return None, kurang

    out = df[list(ada.keys())].rename(columns=ada).copy()

    if "tanggal_kunjungan" in out.columns:
        out["tanggal_kunjungan"] = _ke_tanggal(out["tanggal_kunjungan"]).dt.strftime("%Y-%m-%d")
    if "umur" in out.columns:
        # "69 Thn 11 Bln 28 Hari" -> 69
        out["umur"] = out["umur"].astype(str).str.extract(r"(\d+)")[0]
    if "no_rm" in out.columns:
        out["no_rm"] = (out["no_rm"].astype(str)
                        .str.replace(r"[\s,]+$", "", regex=True).str.strip())

    urutan = [baru for baru in KOLOM_SIGAP.values() if baru in out.columns]
    return out[urutan], kurang


def ke_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")


# ─────────────────────────────────────────────
#  SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style="padding: 1rem 0 1.5rem;">
        <div style="font-size:1.3rem; font-weight:800; color:#0f172a; letter-spacing:-0.03em;">🧹 Data Cleaning SIGAP</div>
        <div style="font-size:0.75rem; color:#56637a; margin-top:0.2rem;">Penyiap data untuk SIGAP-Bojonegoro · v2.0</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("**📂 Unggah Berkas**")
    uploaded_files = st.file_uploader(
        "Pilih satu atau beberapa berkas bulanan",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )
    st.caption("Beberapa berkas bulanan bisa diunggah sekaligus; hasilnya digabung dan diurutkan menurut tanggal kunjungan.")

    st.markdown("---")
    st.markdown("**🔗 Penggabungan**")
    opt_buang_kembar = st.checkbox("Buang baris kembar antarberkas", value=True)
    opt_urut_tanggal = st.checkbox("Urutkan menurut tanggal kunjungan", value=True)
    opt_kolom_asal   = st.checkbox("Tambah kolom Asal Berkas", value=False)

    st.markdown("---")
    st.markdown("**⚙️ Opsi Pembersihan**")

    opt_remove_dup   = st.checkbox("Hapus baris duplikat", value=True)
    opt_strip_comma  = st.checkbox("Hapus trailing koma/spasi", value=True)
    opt_strip_spaces = st.checkbox("Hapus spasi dalam NIK & No Penjamin", value=True)
    opt_format_dates = st.checkbox("Format tanggal ke dd/mm/yyyy", value=True)
    opt_std_ket      = st.checkbox("Standarisasi kolom Keterangan", value=True)
    opt_upper_ket    = st.checkbox("↳ UPPERCASE Keterangan", value=True)
    opt_upper_nama   = st.checkbox("UPPERCASE kolom Nama", value=True)
    opt_title_desa   = st.checkbox("Title Case kolom Desa", value=True)

    st.markdown("---")
    st.markdown("**🔄 Nilai Pengganti**")
    fill_keterangan = st.text_input("Keterangan kosong/simbol", value="Tidak Ada Keterangan")
    fill_pekerjaan  = st.text_input("Pekerjaan kosong", value="Tidak Diketahui")
    fill_rm_lama    = st.text_input("RM Lama kosong", value="-")
    fill_desa       = st.text_input("Desa kosong", value="-")
    fill_no_penjamin = st.text_input("No Penjamin kosong", value="-")

    st.markdown("---")
    st.markdown("**📤 Keluaran**")
    st.caption("Setiap hasil tersedia dalam dua bentuk: arsip lengkap dan berkas siap unggah ke SIGAP, masing-masing sebagai .xlsx maupun .csv.")
    sheet_name = st.text_input("Nama sheet (Excel)", value="Data Bersih")


# ─────────────────────────────────────────────
#  MAIN CONTENT
# ─────────────────────────────────────────────
render_hero_3d()

if not uploaded_files:
    st.markdown("""
    <div class="upload-hint">
        <div style="font-size:2.5rem; margin-bottom:0.75rem;">📂</div>
        <div style="color:#1039a8; font-weight:700; font-size:1rem; margin-bottom:0.4rem;">
            Unggah berkas RME di sidebar kiri — boleh beberapa bulan sekaligus
        </div>
        <div style="font-size:0.82rem; color:#56637a;">
            Format yang didukung: .xlsx · .xls · .csv<br>
            Termasuk berkas .xls dari RME yang sebenarnya berisi teks berpemisah tab
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="section-title">📋 Fitur Pembersihan</div>
    """, unsafe_allow_html=True)

    cols = st.columns(3)
    features = [
        ("🔗", "Gabung Banyak Berkas", "Beberapa berkas bulanan menjadi satu, urut tanggal"),
        ("📤", "Dua Bentuk Unduhan", "Arsip lengkap, atau 8 kolom siap unggah ke SIGAP"),
        ("🗑️", "Hapus Duplikat", "Deteksi dan hapus baris yang persis sama"),
        ("✂️", "Trailing Koma", "Hapus karakter ' ,' di akhir sel (No RM, NIK, dll)"),
        ("📅", "Format Tanggal", "Ubah ke format dd/mm/yyyy secara konsisten"),
        ("📝", "Standarisasi Keterangan", "Ubah '-' dan '=' menjadi teks bermakna"),
        ("🔠", "Normalisasi Teks", "UPPERCASE nama, Title Case desa"),
        ("📋", "Isi Nilai Kosong", "Isi kolom kosong dengan nilai default"),
    ]
    for i, (icon, title, desc) in enumerate(features):
        with cols[i % 3]:
            st.markdown(f"""
            <div style="background:rgba(255,255,255,.74); border:1px solid rgba(148,163,184,.30);
                        border-radius:14px; padding:1rem; margin-bottom:0.75rem; text-align:center;
                        box-shadow:0 2px 4px rgba(15,23,42,.04), 0 10px 24px rgba(30,94,255,.12);">
                <div style="font-size:1.8rem;">{icon}</div>
                <div style="color:#0f172a; font-weight:700; font-size:0.9rem; margin:0.4rem 0 0.3rem;">{title}</div>
                <div style="color:#56637a; font-size:0.78rem;">{desc}</div>
            </div>
            """, unsafe_allow_html=True)

    st.stop()


# ─────────────────────────────────────────────
#  LOAD DATA
# ─────────────────────────────────────────────
berkas_masuk = [(f.name, f.getvalue()) for f in uploaded_files]

try:
    with st.spinner(f"Membaca dan menggabungkan {len(berkas_masuk)} berkas..."):
        df_raw, ringkasan, jml_kembar, kol_tgl = gabung_berkas(berkas_masuk, opt_buang_kembar)
except Exception as e:
    st.error(f"❌ Gagal membaca berkas: {e}")
    st.stop()

asal_berkas = df_raw["_asal_berkas"].copy()
urut_tanggal = df_raw["_urut_tanggal"].copy()
df_raw = df_raw.drop(columns=["_urut_tanggal", "_asal_berkas"])
if opt_kolom_asal:
    df_raw["Asal Berkas"] = asal_berkas.values

# ── Ringkasan penggabungan ──────────────────────────────────
if len(berkas_masuk) > 1 or jml_kembar:
    with st.expander(f"🔗 Ringkasan penggabungan — {len(berkas_masuk)} berkas, {len(df_raw):,} baris".replace(",", "."), expanded=True):
        st.dataframe(pd.DataFrame(ringkasan), use_container_width=True, hide_index=True)
        pesan = []
        if kol_tgl:
            pesan.append(f"Diurutkan menurut kolom **{kol_tgl}**.")
        else:
            pesan.append("⚠️ Kolom tanggal kunjungan tidak ditemukan, urutan mengikuti urutan unggah.")
        if jml_kembar:
            pesan.append(
                f"**{jml_kembar:,}** baris kembar ditemukan dan **dibuang**.".replace(",", ".")
                if opt_buang_kembar else
                f"**{jml_kembar:,}** baris kembar ditemukan dan **dipertahankan**.".replace(",", ".")
            )
        else:
            pesan.append("Tidak ada baris kembar antarberkas.")
        tak_terbaca = int(urut_tanggal.isna().sum())
        if tak_terbaca:
            pesan.append(f"⚠️ **{tak_terbaca}** baris tanggalnya tidak terbaca dan ditaruh di urutan paling akhir.")
        st.markdown(" ".join(pesan))


# ─────────────────────────────────────────────
#  DETECT ISSUES
# ─────────────────────────────────────────────
issues = detect_issues(df_raw)
total_nulls = sum(issues.get("nulls", {}).values())
dup_count   = issues.get("duplicates", 0)

# Metric cards
st.markdown(f"""
<div class="metric-row">
    <div class="metric-card">
        <div class="label">Total Baris</div>
        <div class="value">{len(df_raw):,}</div>
    </div>
    <div class="metric-card">
        <div class="label">Total Kolom</div>
        <div class="value">{len(df_raw.columns)}</div>
    </div>
    <div class="metric-card">
        <div class="label">Duplikat</div>
        <div class="value {'bad' if dup_count else 'good'}">{dup_count}</div>
    </div>
    <div class="metric-card">
        <div class="label">Nilai Kosong</div>
        <div class="value {'bad' if total_nulls else 'good'}">{total_nulls:,}</div>
    </div>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
#  TWO COLUMNS: Issues + Preview
# ─────────────────────────────────────────────
col_left, col_right = st.columns([1, 2])

with col_left:
    st.markdown('<div class="section-title">🔍 Masalah Terdeteksi</div>', unsafe_allow_html=True)

    if not issues:
        st.markdown('<div class="log-item"><span class="log-icon">✅</span> Data sudah bersih!</div>', unsafe_allow_html=True)
    else:
        if "duplicates" in issues:
            st.markdown(f"""
            <div class="log-item">
                <span class="log-icon">🔴</span>
                <span><b style="color:#f87171">{issues['duplicates']} baris duplikat</b></span>
            </div>""", unsafe_allow_html=True)

        if "trailing_comma" in issues:
            cols_str = ", ".join(issues["trailing_comma"][:4])
            st.markdown(f"""
            <div class="log-item">
                <span class="log-icon">🟡</span>
                <span>Trailing koma:<br><span class="badge badge-yellow">{cols_str}</span></span>
            </div>""", unsafe_allow_html=True)

        if "nulls" in issues:
            null_items = "".join([
                f'<span class="badge badge-red" style="margin:2px;">{c}: {v}</span> '
                for c, v in list(issues["nulls"].items())[:8]
            ])
            st.markdown(f"""
            <div class="log-item">
                <span class="log-icon">🔴</span>
                <span>Nilai kosong:<br>{null_items}</span>
            </div>""", unsafe_allow_html=True)

        if "placeholders" in issues:
            ph_str = ", ".join(issues["placeholders"][:4])
            st.markdown(f"""
            <div class="log-item">
                <span class="log-icon">🟡</span>
                <span>Nilai placeholder (-/=):<br>
                <span class="badge badge-yellow">{ph_str}</span></span>
            </div>""", unsafe_allow_html=True)

        if "date_cols" in issues:
            st.markdown(f"""
            <div class="log-item">
                <span class="log-icon">🔵</span>
                <span>Kolom tanggal: <span class="badge badge-blue">{', '.join(issues['date_cols'])}</span></span>
            </div>""", unsafe_allow_html=True)


    # Column selector for dropping
    st.markdown('<div class="section-title">🗂️ Pilih Kolom</div>', unsafe_allow_html=True)
    cols_to_drop = st.multiselect(
        "Kolom yang ingin dihapus",
        options=list(df_raw.columns),
        default=[],
        placeholder="Pilih kolom..."
    )
    sort_col = st.selectbox(
        "Urutkan berdasarkan",
        options=["(tidak diurutkan)"] + list(df_raw.columns),
        index=0
    )

with col_right:
    st.markdown('<div class="section-title">👁️ Preview Data Asli (10 baris pertama)</div>', unsafe_allow_html=True)
    st.dataframe(df_raw.head(10), use_container_width=True, height=280)


# ─────────────────────────────────────────────
#  CLEAN BUTTON
# ─────────────────────────────────────────────
st.markdown("---")
btn_col, _ = st.columns([1, 3])
with btn_col:
    run_clean = st.button("🚀 Bersihkan Data Sekarang", use_container_width=True)

if run_clean or "df_clean" in st.session_state:

    if run_clean:
        # Build options dict
        strip_space_cols = []
        for col in df_raw.select_dtypes(include="object").columns:
            if "nik" in col.lower() or "penjamin" in col.lower():
                strip_space_cols.append(col)

        fill_nulls_map = {}
        if fill_pekerjaan and "Pekerjaan" in df_raw.columns:
            fill_nulls_map["Pekerjaan"] = fill_pekerjaan
        if fill_rm_lama and "RM Lama" in df_raw.columns:
            fill_nulls_map["RM Lama"] = fill_rm_lama
        if fill_desa and "Desa" in df_raw.columns:
            fill_nulls_map["Desa"] = fill_desa
        if fill_no_penjamin and "No Penjamin" in df_raw.columns:
            fill_nulls_map["No Penjamin"] = fill_no_penjamin

        rename_map = {}
        if "kategori" in df_raw.columns:
            rename_map["kategori"] = "Kategori"

        options = {
            "remove_duplicates":     opt_remove_dup,
            "strip_trailing_comma":  opt_strip_comma,
            "strip_spaces":          opt_strip_spaces,
            "strip_spaces_cols":     strip_space_cols,
            "format_dates":          opt_format_dates,
            "date_format":           "%d/%m/%Y",
            "standardize_keterangan":opt_std_ket,
            "placeholder_replacement": fill_keterangan,
            "uppercase_keterangan":  opt_upper_ket,
            "uppercase_nama":        opt_upper_nama,
            "titlecase_desa":        opt_title_desa,
            "fill_nulls":            fill_nulls_map,
            "rename_cols":           rename_map,
            "drop_cols":             cols_to_drop,
            "sort_by":               sort_col if sort_col != "(tidak diurutkan)" else None,
        }

        with st.spinner("Membersihkan data..."):
            df_clean, clean_log = clean_dataframe(df_raw, options)
        st.session_state["df_clean"] = df_clean
        st.session_state["clean_log"] = clean_log

    df_clean = st.session_state["df_clean"]
    clean_log = st.session_state.get("clean_log", [])

    # ── Results ────────────────────────────────
    removed_rows = len(df_raw) - len(df_clean)
    st.markdown(f"""
    <div class="metric-row">
        <div class="metric-card">
            <div class="label">Baris Tersisa</div>
            <div class="value good">{len(df_clean):,}</div>
        </div>
        <div class="metric-card">
            <div class="label">Baris Dihapus</div>
            <div class="value {'bad' if removed_rows else 'good'}">{removed_rows}</div>
        </div>
        <div class="metric-card">
            <div class="label">Kolom Tersisa</div>
            <div class="value">{len(df_clean.columns)}</div>
        </div>
        <div class="metric-card">
            <div class="label">Nilai Kosong</div>
            <div class="value good">{int(df_clean.isnull().sum().sum())}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    res_left, res_right = st.columns([1, 2])

    with res_left:
        st.markdown('<div class="section-title">✅ Log Pembersihan</div>', unsafe_allow_html=True)
        for icon, msg in clean_log:
            st.markdown(f"""
            <div class="log-item">
                <span class="log-icon">{icon}</span>
                <span>{msg}</span>
            </div>""", unsafe_allow_html=True)

    with res_right:
        st.markdown('<div class="section-title">✨ Preview Data Bersih</div>', unsafe_allow_html=True)
        st.dataframe(df_clean.head(15), use_container_width=True, height=320)

    # ── Unduhan ────────────────────────────────
    st.markdown("---")

    if opt_urut_tanggal:
        kol_tgl_bersih = cari_kolom_tanggal(df_clean)
        if kol_tgl_bersih:
            _t = _ke_tanggal(df_clean[kol_tgl_bersih])
            df_clean = (df_clean.assign(_t=_t)
                        .sort_values("_t", kind="stable", na_position="last")
                        .drop(columns="_t").reset_index(drop=True))

    # Nama berkas mengikuti rentang tanggal isinya, bukan nama berkas asal.
    _t = (_ke_tanggal(df_clean[cari_kolom_tanggal(df_clean)])
          if cari_kolom_tanggal(df_clean) else pd.Series([], dtype="datetime64[ns]"))
    if len(_t) and _t.notna().any():
        base_name = f"RME_{_t.min():%Y-%m-%d}_sd_{_t.max():%Y-%m-%d}"
    else:
        base_name = "RME_gabungan"

    df_sigap, kolom_kurang = ke_format_sigap(df_clean)

    st.markdown('<div class="section-title">📤 Unduh Hasil</div>', unsafe_allow_html=True)
    u1, u2 = st.columns(2)

    with u1:
        st.markdown(f"**📚 Arsip lengkap** — {len(df_clean.columns)} kolom, {len(df_clean):,} baris".replace(",", "."))
        st.caption("Seluruh kolom apa adanya, termasuk identitas pasien. Simpan di tempat yang aman.")
        st.download_button(
            "⬇️ Arsip lengkap (.xlsx)",
            data=to_excel_bytes(df_clean, sheet_name),
            file_name=f"{base_name}_lengkap.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.download_button(
            "⬇️ Arsip lengkap (.csv)",
            data=ke_csv_bytes(df_clean.fillna("-")),
            file_name=f"{base_name}_lengkap.csv",
            mime="text/csv",
            use_container_width=True,
        )

    with u2:
        if df_sigap is None:
            st.warning("Kolom yang dibutuhkan SIGAP tidak ditemukan pada data ini.")
        else:
            st.markdown(f"**🏥 Siap unggah ke SIGAP** — {len(df_sigap.columns)} kolom, {len(df_sigap):,} baris".replace(",", "."))
            st.caption("Tanpa nama, NIK, alamat, dan nomor penjamin — aman dibawa ke luar puskesmas.")
            st.download_button(
                "⬇️ Siap SIGAP (.csv)",
                data=ke_csv_bytes(df_sigap),
                file_name=f"{base_name}_sigap.csv",
                mime="text/csv",
                use_container_width=True,
                type="primary",
            )
            st.download_button(
                "⬇️ Siap SIGAP (.xlsx)",
                data=to_excel_bytes(df_sigap, "Data SIGAP"),
                file_name=f"{base_name}_sigap.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            if kolom_kurang:
                st.caption("Kolom asal yang tidak ditemukan: " + ", ".join(kolom_kurang))
            with st.expander("Pratinjau data siap SIGAP"):
                st.dataframe(df_sigap.head(15), use_container_width=True, hide_index=True)

    # Null checker after clean
    remaining_nulls = df_clean.isnull().sum()
    remaining_nulls = remaining_nulls[remaining_nulls > 0]
    if not remaining_nulls.empty:
        with st.expander("⚠️ Nilai kosong yang masih tersisa"):
            st.dataframe(
                remaining_nulls.reset_index().rename(columns={"index": "Kolom", 0: "Jumlah Kosong"}),
                use_container_width=True
            )
