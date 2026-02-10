import streamlit as st
import pandas as pd
import sqlite3
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
import xlsxwriter

# ==========================================
# 1. DATABASE MANAGEMENT (SQLITE)
# ==========================================
DB_NAME = "sekolah.db"

def init_db():
    """Inisialisasi Database dan Tabel jika belum ada"""
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    
    # Tabel Config (Info Sekolah)
    c.execute('''CREATE TABLE IF NOT EXISTS config (key TEXT PRIMARY KEY, value TEXT)''')
    
    # Tabel Master
    c.execute('''CREATE TABLE IF NOT EXISTS master_guru (nama TEXT PRIMARY KEY)''')
    c.execute('''CREATE TABLE IF NOT EXISTS master_mapel (nama TEXT PRIMARY KEY, kkm INTEGER DEFAULT 75)''')
    c.execute('''CREATE TABLE IF NOT EXISTS master_kelas (nama TEXT PRIMARY KEY, wali_kelas TEXT)''')
    
    # Tabel Siswa
    c.execute('''CREATE TABLE IF NOT EXISTS siswa (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nama TEXT, nisn TEXT, nipd TEXT, jk TEXT, kelas TEXT,
        UNIQUE(nisn)
    )''')
    
    # Tabel Penugasan Guru (Jadwal)
    c.execute('''CREATE TABLE IF NOT EXISTS penugasan (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        guru TEXT, mapel TEXT, kelas TEXT,
        UNIQUE(guru, mapel, kelas)
    )''')
    
    # Tabel Nilai
    c.execute('''CREATE TABLE IF NOT EXISTS nilai (
        siswa_id INTEGER, mapel TEXT, nilai INTEGER,
        PRIMARY KEY (siswa_id, mapel)
    )''')
    
    # Tabel Non-Akademik (Absen & Kepribadian)
    c.execute('''CREATE TABLE IF NOT EXISTS non_akademik (
        siswa_id INTEGER PRIMARY KEY,
        rapi TEXT, disiplin TEXT, jujur TEXT,
        sakit INTEGER, izin INTEGER, alpha INTEGER
    )''')
    
    # Default Info Sekolah jika kosong
    c.execute("SELECT count(*) FROM config")
    if c.fetchone()[0] == 0:
        defaults = {
            "nama_sekolah": "SMA ISLAM AL-GHOZALI",
            "alamat": "Jl. Permata No. 19 Desa Curug",
            "kepsek": "Antoni Firdaus M.Pd.",
            "semester": "Genap", "tahun_ajar": "2024/2025",
            "kota": "Gunungsindur", "tgl_raport": "20 Maret 2025"
        }
        for k, v in defaults.items():
            c.execute("INSERT INTO config VALUES (?, ?)", (k, v))
            
    conn.commit()
    conn.close()

# --- FUNGSI CRUD HELPER ---
def run_query(query, params=(), fetch=False):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    try:
        c.execute(query, params)
        if fetch:
            data = c.fetchall()
            return data
        conn.commit()
    except Exception as e:
        st.error(f"Database Error: {e}")
    finally:
        conn.close()

def get_config():
    data = run_query("SELECT key, value FROM config", fetch=True)
    return {row[0]: row[1] for row in data}

def update_config(key, value):
    run_query("INSERT OR REPLACE INTO config (key, value) VALUES (?, ?)", (key, value))

# ==========================================
# 2. KONFIGURASI STREAMLIT
# ==========================================
st.set_page_config(page_title="Sistem Raport Database", layout="wide", page_icon="🏫")
init_db() # Jalankan init DB di awal

# Helper Docx
def set_cell_bg(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr(); shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), color_hex); tcPr.append(shd)

def terbilang(n):
    angka = ["", "Satu", "Dua", "Tiga", "Empat", "Lima", "Enam", "Tujuh", "Delapan", "Sembilan", "Sepuluh", "Sebelas"]
    if n < 0 or n > 100: return ""
    elif n < 12: return angka[n]
    elif n < 20: return angka[n-10] + " Belas"
    elif n < 100: return angka[n//10] + " Puluh " + angka[n%10]
    elif n == 100: return "Seratus"
    return ""

if 'login_status' not in st.session_state: st.session_state['login_status'] = False

# ==========================================
# 3. WORD GENERATOR (AMBIL DARI DB)
# ==========================================
def generate_docx_db(siswa_id, rank, total_siswa):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.5); section.bottom_margin = Inches(0.5); section.left_margin = Inches(0.5); section.right_margin = Inches(0.5)

    # Ambil Data dari DB
    conf = get_config()
    siswa = run_query("SELECT nama, nisn, nipd, kelas FROM siswa WHERE id=?", (siswa_id,), fetch=True)[0]
    nama_siswa, nisn, nipd, kelas = siswa
    
    wali_data = run_query("SELECT wali_kelas FROM master_kelas WHERE nama=?", (kelas,), fetch=True)
    nama_wali = wali_data[0][0] if wali_data and wali_data[0][0] else "(....................)"
    
    non_akademik = run_query("SELECT rapi, disiplin, jujur, sakit, izin, alpha FROM non_akademik WHERE siswa_id=?", (siswa_id,), fetch=True)
    if non_akademik: rapi, disiplin, jujur, sakit, izin, alpha = non_akademik[0]
    else: rapi, disiplin, jujur, sakit, izin, alpha = "-", "-", "-", 0, 0, 0

    # Header
    p = doc.add_paragraph(conf['nama_sekolah']); p.alignment=1; p.runs[0].bold=True; p.runs[0].font.size=Pt(16)
    doc.add_paragraph(f"LAPORAN HASIL BELAJAR - {conf['tahun_ajar']}").alignment=1
    doc.add_paragraph(conf['alamat']).alignment=1
    doc.add_paragraph("-" * 80).alignment=1
    
    # Identitas
    ti = doc.add_table(3,4); ti.autofit=False; ti.columns[0].width=Inches(1.5)
    ti.cell(0,0).text="Nama"; ti.cell(0,1).text=f": {nama_siswa}"; ti.cell(0,2).text="NIPD"; ti.cell(0,3).text=f": {nipd}"
    ti.cell(1,0).text="Kelas"; ti.cell(1,1).text=f": {kelas}"; ti.cell(1,2).text="NISN"; ti.cell(1,3).text=f": {nisn}"
    doc.add_paragraph()

    # Nilai
    tn = doc.add_table(2,6); tn.style='Table Grid'
    h0=tn.rows[0].cells; h1=tn.rows[1].cells
    h0[0].merge(h1[0]).text="NO"; h0[1].merge(h1[1]).text="Mata Pelajaran"; h0[2].merge(h1[2]).text="KKM"
    h0[3].merge(h0[4]).text="Nilai"; h1[3].text="Angka"; h1[4].text="Huruf"; h0[5].merge(h1[5]).text="Predikat"
    for c in h0+h1: set_cell_bg(c, "E0F7FA"); c.paragraphs[0].alignment=1; c.paragraphs[0].runs[0].bold=True

    mapels = run_query("SELECT nama, kkm FROM master_mapel")
    tot=0; cnt=0
    for idx, (m_nama, m_kkm) in enumerate(mapels):
        r = tn.add_row().cells
        nilai_data = run_query("SELECT nilai FROM nilai WHERE siswa_id=? AND mapel=?", (siswa_id, m_nama), fetch=True)
        val = nilai_data[0][0] if nilai_data else 0
        
        r[0].text=str(idx+1); r[1].text=m_nama; r[2].text=str(m_kkm); r[3].text=str(val); r[4].text=terbilang(val)
        r[5].text="B" if val>=m_kkm else "C" if val>0 else "-"
        for c in r: c.paragraphs[0].alignment=1
        r[1].paragraphs[0].alignment=0
        tot+=val; 
        if val>0: cnt+=1

    rs = tn.add_row().cells; rs[0].merge(rs[2]).text="Jumlah"; rs[3].text=str(tot)
    rs[0].paragraphs[0].alignment=1; rs[3].paragraphs[0].alignment=1
    ra = tn.add_row().cells; ra[0].merge(ra[2]).text="Rata - rata"; ra[3].text=f"{tot/cnt:.2f}" if cnt else "0"
    ra[0].paragraphs[0].alignment=1; ra[3].paragraphs[0].alignment=1
    doc.add_paragraph()

    # Non Akademik & TTD
    tc = doc.add_table(1,2); tc.style='Table Grid'
    tk = tc.cell(0,0).add_table(4,2); tk.cell(0,0).text="Kepribadian"; tk.cell(1,0).text="Kerapihan"; tk.cell(1,1).text=str(rapi)
    tk.cell(2,0).text="Kedisiplinan"; tk.cell(2,1).text=str(disiplin); tk.cell(3,0).text="Kejujuran"; tk.cell(3,1).text=str(jujur)
    ta = tc.cell(0,1).add_table(4,2); ta.cell(0,0).text="Absensi"; ta.cell(1,0).text="Sakit"; ta.cell(1,1).text=str(sakit)
    ta.cell(2,0).text="Izin"; ta.cell(2,1).text=str(izin); ta.cell(3,0).text="Alpha"; ta.cell(3,1).text=str(alpha)

    doc.add_paragraph(f"\nPeringkat Kelas: {rank} dari {total_siswa} siswa")
    ttd = doc.add_table(1,3); ttd.alignment=1
    ttd.cell(0,0).text="\nOrang Tua\n\n\n(..........)"
    ttd.cell(0,1).text=f"\nKepala Sekolah\n\n\n({conf['kepsek']})"
    ttd.cell(0,2).text=f"\n{conf['kota']}, {conf['tgl_raport']}\nWali Kelas\n\n\n({nama_wali})"
    for c in ttd.rows[0].cells: c.paragraphs[0].alignment=1

    bio = io.BytesIO(); doc.save(bio); bio.seek(0)
    return bio

# ==========================================
# 4. HALAMAN ADMIN (CRUD LENGKAP)
# ==========================================
def admin_page():
    st.sidebar.title("Panel Admin")
    menu = st.sidebar.radio("Menu", ["🏠 Dashboard", "👨‍🎓 Data Siswa", "⚙️ Data Master", "👨‍🏫 Penugasan & Wali", "📊 Monitoring", "⚙️ Info Sekolah"])
    st.title("Administrator Database")

    # --- CRUD SISWA ---
    if menu == "👨‍🎓 Data Siswa":
        t1, t2, t3 = st.tabs(["📋 Copy-Paste Excel", "Manual Input", "🗂️ Data & Hapus"])
        
        # 1. Copy Paste
        with t1:
            st.info("Format: **KELAS | Nama | NIPD | JK | NISN**")
            raw = st.text_area("Paste Data Siswa", height=200)
            if st.button("Simpan Data Paste"):
                cnt=0
                for line in raw.strip().split('\n'):
                    p = line.split('\t') if '\t' in line else line.split(',')
                    p = [x.strip() for x in p]
                    if len(p)>=2:
                        k=p[0]; n=p[1]
                        nipd=p[2] if len(p)>2 else "-"; jk=p[3] if len(p)>3 else "-"; nisn=p[4] if len(p)>4 else "-"
                        if k.upper() != "KELAS":
                            # Auto add kelas if not exist
                            run_query("INSERT OR IGNORE INTO master_kelas (nama) VALUES (?)", (k,))
                            run_query("INSERT OR REPLACE INTO siswa (nama, nisn, nipd, jk, kelas) VALUES (?,?,?,?,?)", (n, nisn, nipd, jk, k))
                            cnt+=1
                st.success(f"{cnt} Siswa berhasil disimpan ke Database!")

        # 2. Manual
        with t2:
            with st.form("add_s"):
                c1,c2=st.columns(2)
                nm=c1.text_input("Nama"); ni=c2.text_input("NISN")
                np=c1.text_input("NIPD"); jk=c2.selectbox("JK",["L","P"])
                kls_list = [r[0] for r in run_query("SELECT nama FROM master_kelas", fetch=True)]
                k=st.selectbox("Kelas", kls_list if kls_list else ["Belum Ada Kelas"])
                if st.form_submit_button("Tambah Siswa"):
                    run_query("INSERT INTO siswa (nama, nisn, nipd, jk, kelas) VALUES (?,?,?,?,?)", (nm, ni, np, jk, k))
                    st.success("Siswa ditambahkan")

        # 3. View & Delete
        with t3:
            df = pd.read_sql("SELECT * FROM siswa", sqlite3.connect(DB_NAME))
            st.dataframe(df)
            
            # Delete feature
            del_id = st.number_input("Masukkan ID Siswa untuk dihapus", min_value=0)
            if st.button("Hapus Siswa"):
                run_query("DELETE FROM siswa WHERE id=?", (del_id,))
                run_query("DELETE FROM nilai WHERE siswa_id=?", (del_id,)) # Hapus nilainya juga
                st.success(f"Siswa ID {del_id} terhapus.")
                st.rerun()

    # --- DATA MASTER ---
    elif menu == "⚙️ Data Master":
        c1, c2, c3 = st.columns(3)
        with c1:
            st.write("#### Guru")
            raw_g = st.text_area("Paste Guru (Baris baru)", height=150)
            if st.button("Update Guru"):
                for l in raw_g.split('\n'):
                    if l.strip(): run_query("INSERT OR IGNORE INTO master_guru (nama) VALUES (?)", (l.strip(),))
                st.success("Tersimpan")
            # Show list
            st.write([r[0] for r in run_query("SELECT nama FROM master_guru", fetch=True)])

        with c2:
            st.write("#### Mapel & KKM")
            raw_m = st.text_area("Paste Mapel", height=150)
            kkm_def = st.number_input("KKM Default", 60, 100, 75)
            if st.button("Update Mapel"):
                for l in raw_m.split('\n'):
                    if l.strip(): run_query("INSERT OR IGNORE INTO master_mapel (nama, kkm) VALUES (?,?)", (l.strip(), kkm_def))
                st.success("Tersimpan")
            # Show list
            st.write(pd.read_sql("SELECT * FROM master_mapel", sqlite3.connect(DB_NAME)))

        with c3:
            st.write("#### Kelas")
            raw_k = st.text_area("Paste Kelas", height=150)
            if st.button("Update Kelas"):
                for l in raw_k.split('\n'):
                    if l.strip(): run_query("INSERT OR IGNORE INTO master_kelas (nama) VALUES (?)", (l.strip(),))
                st.success("Tersimpan")
            st.write([r[0] for r in run_query("SELECT nama FROM master_kelas", fetch=True)])

    # --- PENUGASAN & WALI ---
    elif menu == "👨‍🏫 Penugasan & Wali":
        t1, t2 = st.tabs(["Wali Kelas", "Penugasan Mapel"])
        
        with t1:
            kls = [r[0] for r in run_query("SELECT nama FROM master_kelas", fetch=True)]
            gru = [r[0] for r in run_query("SELECT nama FROM master_guru", fetch=True)]
            
            c1, c2 = st.columns(2)
            k_sel = c1.selectbox("Pilih Kelas", kls)
            g_sel = c2.selectbox("Pilih Wali", gru)
            if st.button("Set Wali Kelas"):
                run_query("UPDATE master_kelas SET wali_kelas=? WHERE nama=?", (g_sel, k_sel))
                st.success(f"Wali kelas {k_sel} diset ke {g_sel}")
            
            st.write("Daftar Wali Kelas:")
            st.dataframe(pd.read_sql("SELECT nama as Kelas, wali_kelas FROM master_kelas", sqlite3.connect(DB_NAME)))

        with t2:
            mpl = [r[0] for r in run_query("SELECT nama FROM master_mapel", fetch=True)]
            
            with st.form("assign"):
                g = st.selectbox("Guru", gru)
                m = st.selectbox("Mapel", mpl)
                ks = st.multiselect("Kelas Ajar", kls)
                if st.form_submit_button("Simpan Penugasan"):
                    for k in ks:
                        run_query("INSERT OR REPLACE INTO penugasan (guru, mapel, kelas) VALUES (?,?,?)", (g, m, k))
                    st.success("Penugasan tersimpan")
            
            st.write("Tabel Penugasan:")
            st.dataframe(pd.read_sql("SELECT * FROM penugasan", sqlite3.connect(DB_NAME)))
            
            del_id = st.number_input("ID Penugasan Hapus", 0)
            if st.button("Hapus Penugasan"):
                run_query("DELETE FROM penugasan WHERE id=?", (del_id,))
                st.success("Terhapus"); st.rerun()

    # --- MONITORING ---
    elif menu == "📊 Monitoring":
        st.subheader("Monitoring Input Nilai")
        # Logic Matrix
        mapels = [r[0] for r in run_query("SELECT nama FROM master_mapel", fetch=True)]
        kelas_list = [r[0] for r in run_query("SELECT nama FROM master_kelas", fetch=True)]
        
        data = []
        for m in mapels:
            row = {"Mapel": m}
            for k in kelas_list:
                # Cek Penugasan
                tugas = run_query("SELECT guru FROM penugasan WHERE mapel=? AND kelas=?", (m, k), fetch=True)
                guru_nama = tugas[0][0] if tugas else ""
                
                # Cek Nilai (Minimal 1 siswa)
                if guru_nama:
                    siswa_ids = run_query("SELECT id FROM siswa WHERE kelas=?", (k,), fetch=True)
                    has_nilai = False
                    if siswa_ids:
                        chk = run_query("SELECT count(*) FROM nilai WHERE mapel=? AND siswa_id=?", (m, siswa_ids[0][0]), fetch=True)
                        if chk[0][0] > 0: has_nilai = True
                    
                    status = f"✅ {guru_nama}" if has_nilai else f"❌ {guru_nama}"
                else:
                    status = "⚠️ Kosong"
                row[k] = status
            data.append(row)
        st.dataframe(pd.DataFrame(data), use_container_width=True)

    elif menu == "⚙️ Info Sekolah":
        conf = get_config()
        with st.form("sch"):
            n=st.text_input("Nama",conf['nama_sekolah']); a=st.text_input("Alamat",conf['alamat'])
            k=st.text_input("Kepsek",conf['kepsek'])
            c1,c2=st.columns(2); ci=c1.text_input("Kota",conf['kota']); tg=c2.text_input("Tgl",conf['tgl_raport'])
            if st.form_submit_button("Simpan"):
                update_config("nama_sekolah", n); update_config("alamat", a)
                update_config("kepsek", k); update_config("kota", ci); update_config("tgl_raport", tg)
                st.success("Tersimpan")

    elif menu == "🏠 Dashboard":
        jml_siswa = run_query("SELECT count(*) FROM siswa", fetch=True)[0][0]
        st.metric("Total Siswa", jml_siswa)

# ==========================================
# 5. HALAMAN GURU
# ==========================================
def guru_page():
    guru = st.session_state['active_user']
    
    # Ambil mapel & kelas yg ditugaskan ke guru ini
    tugas = run_query("SELECT mapel, kelas FROM penugasan WHERE guru=?", (guru,), fetch=True)
    if not tugas: st.warning("Anda belum memiliki penugasan jadwal."); return
    
    # Selectbox filter
    list_mapel = list(set([t[0] for t in tugas]))
    p_mapel = st.selectbox("Pilih Mapel", list_mapel)
    
    list_kelas = [t[1] for t in tugas if t[0] == p_mapel]
    p_kelas = st.selectbox("Pilih Kelas", list_kelas)
    
    st.title(f"Input Nilai: {p_mapel} - {p_kelas}")
    
    # Ambil KKM
    kkm_data = run_query("SELECT kkm FROM master_mapel WHERE nama=?", (p_mapel,), fetch=True)
    kkm = kkm_data[0][0] if kkm_data else 75
    st.info(f"KKM: {kkm}")
    
    # Ambil Siswa
    siswa = run_query("SELECT id, nama FROM siswa WHERE kelas=? ORDER BY nama", (p_kelas,), fetch=True)
    
    t1, t2, t3 = st.tabs(["Manual", "Upload", "Copy-Paste"])
    
    with t1:
        with st.form("input_manual"):
            input_vals = {}
            for sid, snama in siswa:
                # Get current val
                cur = run_query("SELECT nilai FROM nilai WHERE siswa_id=? AND mapel=?", (sid, p_mapel), fetch=True)
                val = cur[0][0] if cur else 0
                input_vals[sid] = st.number_input(f"{snama}", 0, 100, val)
            
            if st.form_submit_button("Simpan"):
                for sid, v in input_vals.items():
                    run_query("INSERT OR REPLACE INTO nilai (siswa_id, mapel, nilai) VALUES (?,?,?)", (sid, p_mapel, v))
                st.success("Tersimpan ke Database!")

    with t3:
        st.info("Copy kolom **Nama** dan **Nilai** dari Excel")
        raw = st.text_area("Paste", height=200)
        if st.button("Proses Paste"):
            cnt=0
            for line in raw.split('\n'):
                p = line.split('\t') if '\t' in line else line.split(',')
                if len(p)>=2:
                    nm=p[0].strip()
                    try: val=int(float(p[1].strip()))
                    except: val=0
                    # Cari ID siswa by nama (fuzzy/exact match logic sederhana)
                    # Di sini pakai exact match lowercase
                    found_sid = next((s[0] for s in siswa if s[1].lower() == nm.lower()), None)
                    if found_sid:
                        run_query("INSERT OR REPLACE INTO nilai (siswa_id, mapel, nilai) VALUES (?,?,?)", (found_sid, p_mapel, val))
                        cnt+=1
            st.success(f"{cnt} nilai tersimpan")

# ==========================================
# 6. HALAMAN WALI KELAS
# ==========================================
def wali_page():
    wali = st.session_state['active_user']
    # Cari kelas binaan
    kelas_data = run_query("SELECT nama FROM master_kelas WHERE wali_kelas=?", (wali,), fetch=True)
    if not kelas_data: st.warning("Anda tidak terdaftar sebagai Wali Kelas."); return
    
    kelas = st.selectbox("Pilih Kelas Binaan", [k[0] for k in kelas_data])
    
    siswa = run_query("SELECT id, nama, nisn FROM siswa WHERE kelas=? ORDER BY nama", (kelas,), fetch=True)
    mapels = [r[0] for r in run_query("SELECT nama FROM master_mapel", fetch=True)]
    
    # Hitung Ranking Logic
    rank_data = []
    for sid, _, _ in siswa:
        total = run_query("SELECT sum(nilai) FROM nilai WHERE siswa_id=?", (sid,), fetch=True)[0][0]
        rank_data.append({"id": sid, "total": total if total else 0})
    rank_data.sort(key=lambda x: x['total'], reverse=True)
    rank_map = {x['id']: i+1 for i, x in enumerate(rank_data)}
    
    t1, t2, t3 = st.tabs(["Non-Akademik", "Leger", "Raport"])
    
    with t1:
        with st.form("non"):
            vals = {}
            for sid, snama, _ in siswa:
                cur = run_query("SELECT * FROM non_akademik WHERE siswa_id=?", (sid,), fetch=True)
                if cur: _, rapi, disp, jujur, s, i, a = cur[0]
                else: rapi, disp, jujur, s, i, a = "-","-","-",0,0,0
                
                with st.expander(snama):
                    c1,c2=st.columns(2)
                    with c1:
                        kr=st.selectbox(f"Rapi {sid}", ["-","AA","BB"], index=["-","AA","BB"].index(rapi) if rapi in ["-","AA","BB"] else 0)
                        kd=st.selectbox(f"Disp {sid}", ["-","AA","BB"], index=["-","AA","BB"].index(disp) if disp in ["-","AA","BB"] else 0)
                        kj=st.selectbox(f"Jujur {sid}", ["-","AA","BB"], index=["-","AA","BB"].index(jujur) if jujur in ["-","AA","BB"] else 0)
                    with c2:
                        sa=st.number_input(f"Sakit {sid}",0,100,s)
                        iz=st.number_input(f"Izin {sid}",0,100,i)
                        al=st.number_input(f"Alpa {sid}",0,100,a)
                    vals[sid] = (kr, kd, kj, sa, iz, al)
            if st.form_submit_button("Simpan"):
                for sid, d in vals.items():
                    run_query("INSERT OR REPLACE INTO non_akademik VALUES (?,?,?,?,?,?,?)", (sid, *d))
                st.success("Tersimpan")

    with t2:
        rows = []
        shorts = {m: m[:4].upper() for m in mapels}
        for idx, (sid, snama, _) in enumerate(siswa):
            r = {"No": idx+1, "Nama": snama}
            tot = 0; cnt = 0
            for m in mapels:
                val = run_query("SELECT nilai FROM nilai WHERE siswa_id=? AND mapel=?", (sid, m), fetch=True)
                v = val[0][0] if val else 0
                r[shorts[m]] = v; tot += v; 
                if v>0: cnt+=1
            r["Total"] = tot; r["Rata"] = f"{tot/cnt:.2f}" if cnt else "0"
            r["Rank"] = rank_map.get(sid)
            rows.append(r)
        
        df = pd.DataFrame(rows)
        st.dataframe(df)
        
        buf = io.BytesIO()
        with pd.ExcelWriter(buf) as w: df.to_excel(w, index=False)
        st.download_button("Download Excel", buf, "Leger.xlsx")

    with t3:
        for sid, snama, _ in siswa:
            c1,c2 = st.columns([4,1])
            c1.write(f"{snama} (Rank {rank_map.get(sid)})")
            docx = generate_docx_db(sid, rank_map.get(sid), len(siswa))
            c2.download_button("Unduh", docx, f"Raport_{snama}.docx", key=f"dl_{sid}")

# ==========================================
# 7. LOGIN SCREEN
# ==========================================
def login_screen():
    st.markdown("<h1 style='text-align:center'>SISTEM RAPORT DATABASE</h1>", unsafe_allow_html=True)
    t1,t2,t3 = st.tabs(["ADMIN", "WALI KELAS", "GURU"])
    
    with t1:
        if st.button("Masuk Admin") and st.text_input("Password", type="password") == "admin":
            st.session_state['login_status'] = True; st.session_state['user_role'] = 'admin'; st.rerun()
            
    with t2:
        walis = [r[0] for r in run_query("SELECT DISTINCT wali_kelas FROM master_kelas WHERE wali_kelas IS NOT NULL", fetch=True)]
        w = st.selectbox("Nama Wali", walis)
        if st.button("Masuk Wali"):
            st.session_state['login_status']=True; st.session_state['user_role']='wali'; st.session_state['active_user']=w; st.rerun()
            
    with t3:
        gurus = [r[0] for r in run_query("SELECT nama FROM master_guru", fetch=True)]
        g = st.selectbox("Nama Guru", gurus)
        if st.button("Masuk Guru"):
            st.session_state['login_status']=True; st.session_state['user_role']='guru'; st.session_state['active_user']=g; st.rerun()

if not st.session_state['login_status']:
    login_screen()
else:
    with st.sidebar:
        if st.button("Keluar"): 
            st.session_state['login_status'] = False; st.rerun()
    
    role = st.session_state['user_role']
    if role == 'admin': admin_page()
    elif role == 'guru': guru_page()
    elif role == 'wali': wali_page()