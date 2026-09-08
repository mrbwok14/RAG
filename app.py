import io
import json
import os
import re
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import mysql.connector
import pandas as pd
import streamlit as st
from langchain_community.document_loaders import PyPDFLoader
from langchain_community.embeddings import HuggingFaceEmbeddings
from langchain_community.vectorstores import Chroma
from langchain_text_splitters import RecursiveCharacterTextSplitter

# Opsional untuk LLM Cloud agar bisa diakses di Streamlit Cloud
from langchain_groq import ChatGroq
from langchain_community.llms import Ollama

# 1. Konfigurasi Halaman (Wide Mode)
st.set_page_config(
    page_title="EduBoard AI: UKRIDA Local RAG Ground Truth",
    page_icon="🎬",
    layout="wide",
)

# Custom CSS dengan Tema Warna Khas UKRIDA & Styling Tabel
st.markdown(
    """
    <style>
        .main-header { 
            font-size: 26px; 
            font-weight: bold; 
            color: #003366; 
        }
        .sub-text { 
            color: #2B6CB0; 
            font-size: 14px; 
            margin-bottom: 20px; 
        }
        .card-box { 
            background-color: #F0F4F8; 
            padding: 15px; 
            border-radius: 8px; 
            border-left: 4px solid #003366; 
            border-top: 1px solid #E2E8F0; 
            border-right: 1px solid #E2E8F0; 
            border-bottom: 1px solid #E2E8F0; 
            margin-bottom: 10px; 
        }
        .stButton>button {
            background-color: #003366;
            color: white;
            border-radius: 6px;
            font-weight: 600;
        }
        .stButton>button:hover {
            background-color: #D69E2E;
            color: #003366;
        }
    </style>
""",
    unsafe_allow_html=True,
)


# --- FUNGSI DATABASE MYSQL (KONEKSI RAILWAY) ---
def get_db_connection():
  if "mysql" in st.secrets:
    db_config = st.secrets["mysql"]
    host = db_config.get("host", "localhost")
    user = db_config.get("user", "root")
    password = db_config.get("password", "")
    database = db_config.get("database", "railway")
    port = int(db_config.get("port", 3306))
  else:
    host = "localhost"
    user = "root"
    password = ""
    database = "railway"
    port = 3306

  return mysql.connector.connect(
      host=host,
      user=user,
      password=password,
      database=database,
      port=port,
      connect_timeout=10,
  )


def check_db_connection():
  try:
    conn = get_db_connection()
    conn.close()
    return True
  except mysql.connector.Error:
    return False


def init_db_tables():
  try:
    if check_db_connection():
      conn = get_db_connection()
      cursor = conn.cursor()

      cursor.execute("""
                CREATE TABLE IF NOT EXISTS users (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    username VARCHAR(100) UNIQUE NOT NULL,
                    password VARCHAR(255) NOT NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)

      # Menambahkan kolom alignment_evaluation jika belum ada di tabel storyboards
      cursor.execute("""
                CREATE TABLE IF NOT EXISTS storyboards (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    username VARCHAR(100),
                    program_studi VARCHAR(255),
                    nama_mata_kuliah VARCHAR(255),
                    project_name VARCHAR(255),
                    learning_objectives TEXT,
                    target_audience VARCHAR(100),
                    visual_style VARCHAR(100),
                    rps_filename VARCHAR(255) DEFAULT '-',
                    alignment_evaluation TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)

      cursor.execute("""
                CREATE TABLE IF NOT EXISTS storyboard_scenes (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    storyboard_id INT,
                    judul_scene VARCHAR(255),
                    learning_objective_tag VARCHAR(50),
                    visualisasi TEXT,
                    instruksi_visual TEXT,
                    animasi TEXT,
                    on_screen_text TEXT,
                    voice_over_text TEXT,
                    backsound VARCHAR(100),
                    durasi VARCHAR(50),
                    saran_info TEXT,
                    FOREIGN KEY (storyboard_id) REFERENCES storyboards(id) ON DELETE CASCADE
                )
            """)

      conn.commit()
      cursor.close()
      conn.close()
  except Exception:
    pass


init_db_tables()


def parse_llm_json_response(response_text):
  try:
    if not response_text:
      return None

    text = response_text.strip()

    if "```" in text:
      match_code = re.findall(r"```(?:json)?\s*(.*?)\s*```", text, re.DOTALL)
      if match_code:
        text = match_code[0].strip()
      else:
        text = re.sub(r"^```[a-zA-Z]*\n?", "", text)
        text = re.sub(r"\n?```$", "", text)
        text = text.strip()

    try:
      return json.loads(text)
    except Exception:
      pass

    match_array = re.search(r"\[\s*\{.*\}\s*\]", text, re.DOTALL)
    if match_array:
      try:
        return json.loads(match_array.group(0))
      except Exception:
        pass

    start_idx = text.find("[")
    end_idx = text.rfind("]")
    if start_idx != -1 and end_idx != -1 and end_idx > start_idx:
      potential_json = text[start_idx : end_idx + 1]
      try:
        return json.loads(potential_json)
      except Exception:
        pass

    return None
  except Exception:
    return None


def generate_pptx(df, program_studi, nama_mata_kuliah, project_title):
  prs = Presentation()

  for index, row in df.iterrows():
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    rows_count = 5
    cols_count = 4
    left = Inches(0.8)
    top = Inches(0.8)
    width = Inches(11.7)
    height = Inches(6.0)

    table_shape = slide.shapes.add_table(
        rows_count, cols_count, left, top, width, height
    )
    table = table_shape.table

    table.columns[0].width = Inches(3.2)
    table.columns[1].width = Inches(3.2)
    table.columns[2].width = Inches(3.0)
    table.columns[3].width = Inches(2.3)

    lo_tag = row.get("Learning Objective", "LO1")
    table.cell(0, 0).text = f"Program Studi: {program_studi}\nTarget: {lo_tag}"
    table.cell(0, 1).text = ":"

    cell_judul_start = table.cell(0, 2)
    cell_judul_end = table.cell(0, 3)
    cell_judul_start.merge(cell_judul_end)
    cell_judul_start.text = f"Judul Scene:\n{row.get('Judul Scene', '')}"

    table.cell(1, 0).text = f"Nama Mata Kuliah\n{nama_mata_kuliah}"
    table.cell(1, 1).text = ":"
    table.cell(1, 2).text = ""
    table.cell(1, 3).text = ""

    table.cell(2, 0).text = "Visualisasi"
    table.cell(2, 1).text = ""
    table.cell(2, 2).text = "Instruksi untuk Visualisasi"
    table.cell(2, 3).text = "Animasi"

    cell_top_left = table.cell(2, 0)
    cell_bottom_right = table.cell(3, 1)
    cell_top_left.merge(cell_bottom_right)
    cell_top_left.text = f"Visualisasi:\n{row.get('Visualisasi', '')}"

    table.cell(2, 2).text = str(row.get("Instruksi untuk Visual", ""))
    table.cell(2, 3).text = str(row.get("Animasi", ""))

    table.cell(3, 2).text = ""
    table.cell(3, 3).text = ""

    table.cell(4, 0).text = f"On-Screen Text:\n{row.get('On-Screen Text', '')}"
    table.cell(4, 1).text = f"Voice Over Text:\n{row.get('Voice Over Text', '')}"
    table.cell(4, 2).text = f"Backsound:\n{row.get('Backsound', '')}"
    table.cell(4, 3).text = f"Durasi:\n{row.get('Durasi', '')}"

    if row.get("Saran dan Info"):
      saran_box = slide.shapes.add_textbox(
          Inches(0.8), Inches(6.9), Inches(11.7), Inches(0.5)
      )
      tf_saran = saran_box.text_frame
      tf_saran.word_wrap = True
      p_saran = tf_saran.paragraphs[0]
      p_saran.text = f"Saran / Info: {row.get('Saran dan Info', '')}"
      p_saran.font.size = Pt(11)
      p_saran.font.color.rgb = RGBColor(51, 51, 51)

    for r_idx, row_cells in enumerate(table.rows):
      for c_idx, cell in enumerate(row_cells.cells):
        cell.fill.solid()
        if r_idx in [0, 1, 2] and c_idx in [0, 2, 3]:
          cell.fill.fore_color.rgb = RGBColor(0, 51, 102)
          for p in cell.text_frame.paragraphs:
            p.font.bold = True
            p.font.color.rgb = RGBColor(255, 255, 255)
            p.font.size = Pt(12)
        else:
          cell.fill.fore_color.rgb = RGBColor(240, 244, 248)
          for p in cell.text_frame.paragraphs:
            p.font.color.rgb = RGBColor(51, 51, 51)
            p.font.size = Pt(11)

  ppt_stream = io.BytesIO()
  prs.save(ppt_stream)
  ppt_stream.seek(0)
  return ppt_stream


# --- MANAJEMEN SESSION AUTH ---
if "logged_in" not in st.session_state:
  st.session_state.logged_in = False
if "username" not in st.session_state:
  st.session_state.username = ""
if "is_admin" not in st.session_state:
  st.session_state.is_admin = False

# ==========================================
# KONDISI 1: JIKA BELUM LOGIN
# ==========================================
if not st.session_state.logged_in:
  st.markdown(
      '<div style="text-align: center; margin-top: 50px;" class="main-header">'
      " UKRIDA EduBoard Ground Truth AI 🎓🎬</div>",
      unsafe_allow_html=True,
  )
  st.markdown(
      '<div style="text-align: center;" class="sub-text">Automated'
      " Storyboard Generator based on Ground Truth & Local RAG</div>",
      unsafe_allow_html=True,
  )

  col1, col2, col3 = st.columns([1, 1.2, 1])

  with col2:
    if not check_db_connection():
      st.error(
          "⚠️ **Database Railway Tidak Terhubung!** Pastikan koneksi jaringan"
          " dan kredensial MySQL terpasang."
      )

    tab_login, tab_register = st.tabs(["🔐 Login", "📝 Register"])

    with tab_login:
      st.subheader("Masuk ke Akun Anda")
      login_user = st.text_input("Username", key="login_user")
      login_pass = st.text_input(
          "Password", type="password", key="login_pass"
      )

      if st.button("Login", use_container_width=True, type="primary"):
        if not login_user or not login_pass:
          st.warning("Isi username dan password terlebih dahulu.")
        else:
          try:
            conn = get_db_connection()
            cursor = conn.cursor(dictionary=True)
            cursor.execute(
                "SELECT * FROM users WHERE username = %s AND password = %s",
                (login_user, login_pass),
            )
            user = cursor.fetchone()
            cursor.close()
            conn.close()

            if user:
              st.session_state.logged_in = True
              st.session_state.username = user["username"]
              if user["username"].lower() == "admin":
                st.session_state.is_admin = True
              else:
                st.session_state.is_admin = False

              st.session_state.storyboard_df = pd.DataFrame(columns=[
                  "Judul Scene",
                  "Learning Objective",
                  "Visualisasi",
                  "Instruksi untuk Visual",
                  "Animasi",
                  "On-Screen Text",
                  "Voice Over Text",
                  "Backsound",
                  "Durasi",
                  "Saran dan Info",
              ])
              st.session_state.uploaded_rps_name = "-"
              st.session_state.retriever = None

              st.success("Login berhasil! Memuat halaman...")
              st.rerun()
            else:
              st.error("Username atau password salah.")
          except Exception as e:
            st.error(f"Terjadi kesalahan database: {e}")

    with tab_register:
      st.subheader("Buat Akun Baru")
      reg_user = st.text_input("Pilih Username", key="reg_user")
      reg_pass = st.text_input(
          "Pilih Password", type="password", key="reg_pass"
      )
      reg_pass_confirm = st.text_input(
          "Konfirmasi Password", type="password", key="reg_pass_confirm"
      )

      if st.button("Daftar Akun", use_container_width=True):
        if not reg_user or not reg_pass:
          st.warning("Username dan password tidak boleh kosong.")
        elif reg_pass != reg_pass_confirm:
          st.error("Konfirmasi password tidak cocok.")
        else:
          try:
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute(
                "SELECT * FROM users WHERE username = %s", (reg_user,)
            )
            if cursor.fetchone():
              st.error("Username sudah terpakai.")
            else:
              cursor.execute(
                  "INSERT INTO users (username, password) VALUES (%s, %s)",
                  (reg_user, reg_pass),
              )
              conn.commit()
              st.success("✅ Registrasi berhasil! Silakan ke tab Login.")
            cursor.close()
            conn.close()
          except Exception as e:
            st.error(f"Gagal mendaftarkan akun: {e}")

# ==========================================
# KONDISI 2: JIKA LOGIN SEBAGAI SUPER ADMIN
# ==========================================
elif st.session_state.is_admin:
  with st.sidebar:
    st.markdown("### 🛡️ Super Admin Panel")
    st.write(f"👤 Admin: **{st.session_state.username}**")
    if check_db_connection():
      st.caption("🟢 MySQL Railway: Connected")
    else:
      st.caption("🔴 MySQL Railway: Disconnected")

    if st.button("🚪 Logout", use_container_width=True):
      st.session_state.logged_in = False
      st.session_state.username = ""
      st.session_state.is_admin = False
      st.session_state.storyboard_df = pd.DataFrame(columns=[
          "Judul Scene",
          "Learning Objective",
          "Visualisasi",
          "Instruksi untuk Visual",
          "Animasi",
          "On-Screen Text",
          "Voice Over Text",
          "Backsound",
          "Durasi",
          "Saran dan Info",
      ])
      st.session_state.uploaded_rps_name = "-"
      st.session_state.retriever = None
      st.rerun()
    st.divider()

  st.markdown(
      '<div class="main-header">🛡️ Super Admin Dashboard - UKRIDA EduBoard</div>',
      unsafe_allow_html=True,
  )
  st.markdown(
      '<div class="sub-text">Kelola data pengguna dan riwayat storyboard'
      " berbasis Ground Truth.</div>",
      unsafe_allow_html=True,
  )

  tab_admin_users, tab_admin_projects = st.tabs([
      "👥 Kelola Data User",
      "📂 Kelola Data Storyboard",
  ])

  with tab_admin_users:
    st.subheader("Daftar Pengguna Sistem")
    if check_db_connection():
      try:
        conn = get_db_connection()
        df_users = pd.read_sql(
            "SELECT id, username, created_at FROM users", conn
        )
        conn.close()

        if not df_users.empty:
          st.dataframe(df_users, use_container_width=True, hide_index=True)

          st.markdown("### Hapus User")
          del_user_id = st.number_input(
              "Masukkan ID User yang ingin dihapus:",
              min_value=1,
              step=1,
              key="del_user",
          )
          if st.button("🗑️ Hapus User Ini", type="primary"):
            try:
              conn = get_db_connection()
              cursor = conn.cursor()
              cursor.execute("DELETE FROM users WHERE id = %s", (del_user_id,))
              conn.commit()
              cursor.close()
              conn.close()
              st.success(f"User dengan ID {del_user_id} berhasil dihapus.")
              st.rerun()
            except Exception as e:
              st.error(f"Gagal menghapus user: {e}")
        else:
          st.info("Belum ada data user terdaftar.")
      except Exception as e:
        st.error(f"Gagal memuat data user: {e}")

  with tab_admin_projects:
    st.subheader("Daftar Proyek Storyboard Tersimpan & Evaluasi RAG")
    if check_db_connection():
      try:
        conn = get_db_connection()
        df_projects = pd.read_sql(
            "SELECT id, username, program_studi, nama_mata_kuliah, project_name,"
            " rps_filename, alignment_evaluation, created_at FROM storyboards"
            " ORDER BY id DESC",
            conn,
        )
        conn.close()

        if not df_projects.empty:
          # Tampilkan tabel proyek tanpa kolom teks panjang evaluation agar rapi
          df_display = df_projects.drop(columns=["alignment_evaluation"])

          event_admin = st.dataframe(
              df_display,
              use_container_width=True,
              hide_index=True,
              on_select="rerun",
              selection_mode="single-row",
          )

          # --- KOMPONEN: LEARNING OBJECTIVE-GROUNDED RAG EVALUATION BERDASARKAN PROYEK YANG DIKLIK ---
          selected_admin_rows = event_admin.selection.rows
          if selected_admin_rows:
            selected_admin_idx = selected_admin_rows[0]
            selected_row_data = df_projects.iloc[selected_admin_idx]
            proj_title = selected_row_data["project_name"]
            eval_text = selected_row_data["alignment_evaluation"]

            st.markdown("---")
            st.markdown(
                f"### 🎯 Evaluasi RAG untuk Proyek: *{proj_title}*"
            )
            if eval_text and str(eval_text).strip() != "None":
              st.info(eval_text)
            else:
              st.warning(
                  "⚠️ Tidak ada data Evaluasi RAG tersimpan untuk proyek ini."
              )

          st.divider()
          st.markdown("### Hapus Proyek Storyboard")
          del_proj_id = st.number_input(
              "Masukkan ID Proyek Storyboard yang ingin dihapus:",
              min_value=1,
              step=1,
              key="del_proj",
          )
          if st.button("🗑️ Hapus Proyek Ini", type="primary"):
            try:
              conn = get_db_connection()
              cursor = conn.cursor()
              cursor.execute(
                  "DELETE FROM storyboards WHERE id = %s", (del_proj_id,)
              )
              conn.commit()
              cursor.close()
              conn.close()
              st.success(
                  f"Proyek Storyboard ID {del_proj_id} berhasil dihapus."
              )
              st.rerun()
            except Exception as e:
              st.error(f"Gagal menghapus proyek: {e}")
        else:
          st.info("Belum ada proyek storyboard tersimpan di database.")
      except Exception as e:
        st.error(f"Gagal memuat data storyboard: {e}")

# ==========================================
# KONDISI 3: JIKA LOGIN SEBAGAI PENGGUNA BIASA
# ==========================================
else:
  with st.sidebar:
    st.markdown("### 🏛️ UKRIDA EduBoard")
    st.write(f"👤 User: **{st.session_state.username}**")

    if check_db_connection():
      st.caption("🟢 MySQL Railway: Connected")
    else:
      st.caption("🔴 MySQL Railway: Disconnected")

    st.divider()
    st.markdown("### ⚙️ Konfigurasi LLM")
    llm_provider = st.selectbox(
        "Pilih Penyedia LLM",
        ["Ollama (Lokal PC)", "Groq API (Cloud / Streamlit Cloud)"],
    )

    if llm_provider == "Groq API (Cloud / Streamlit Cloud)":
      groq_api_key = st.text_input(
          "Masukkan Groq API Key", type="password", value=""
      )
      groq_model = st.selectbox(
          "Model Groq", ["openai/gpt-oss-120b", "openai/gpt-oss-20b"]
      )
    else:
      local_model_name = st.selectbox(
          "Model Ollama", ["llama3", "mistral", "phi3"]
      )

    if st.button("🚪 Logout", use_container_width=True):
      st.session_state.logged_in = False
      st.session_state.username = ""
      st.session_state.is_admin = False
      st.session_state.storyboard_df = pd.DataFrame(columns=[
          "Judul Scene",
          "Learning Objective",
          "Visualisasi",
          "Instruksi untuk Visual",
          "Animasi",
          "On-Screen Text",
          "Voice Over Text",
          "Backsound",
          "Durasi",
          "Saran dan Info",
      ])
      st.session_state.uploaded_rps_name = "-"
      st.session_state.retriever = None
      st.rerun()
    st.divider()

  st.markdown(
      '<div class="main-header">UKRIDA EduBoard AI (Ground Truth RAG'
      ' Mode)</div>',
      unsafe_allow_html=True,
  )
  st.markdown(
      '<div class="sub-text">Sistem RAG lokal berbasis Standar Emas & Ground'
      " Truth Dataset menggunakan Ollama / Cloud LLM</div>",
      unsafe_allow_html=True,
  )

  @st.cache_resource
  def initialize_local_vector_store(file_path):
    loader = PyPDFLoader(file_path)
    documents = loader.load()

    text_splitter = RecursiveCharacterTextSplitter(
        chunk_size=1000, chunk_overlap=200, separators=["\n\n", "\n", " ", ""]
    )
    docs_split = text_splitter.split_documents(documents)

    embeddings = HuggingFaceEmbeddings(model_name="all-MiniLM-L6-v2")
    vectorstore = Chroma.from_documents(docs_split, embeddings)
    return vectorstore.as_retriever(
        search_type="similarity", search_kwargs={"k": 4}
    )

  if "retriever" not in st.session_state:
    st.session_state.retriever = None
  if "rag_contexts" not in st.session_state:
    st.session_state.rag_contexts = []
  if "uploaded_rps_name" not in st.session_state:
    st.session_state.uploaded_rps_name = "-"

  col_left, col_middle, col_right = st.columns([1.1, 1.2, 2.7], gap="medium")

  # ==========================================
  # KOLOM 1: INPUT & SETTINGS
  # ==========================================
  with col_left:
    st.subheader("Project Input & Ground Truth Settings")

    program_studi = st.text_input("Program Studi", value="")
    nama_mata_kuliah = st.text_input("Nama Mata Kuliah", value="")
    project_name = st.text_input(
        "Project Name",
        value="",
    )

    learning_objectives = st.text_area(
        "Learning Objectives",
        value="",
        height=120,
    )

    target_audience = st.selectbox(
        "Target Audience",
        [
            "-- Pilih Target Audience --",
            "Mahasiswa S1/S2 (Advanced)",
            "Beginner Students",
            "General",
        ],
    )
    visual_style = st.selectbox(
        "Visual Style",
        [
            "-- Pilih Visual Style --",
            "Animasi Diagram Blok & Rumus Matematis",
            "Biology Lab",
            "Cinematic",
        ],
    )

    st.markdown("### RAG Knowledge Base (RPS)")
    uploaded_file = st.file_uploader(
        "Upload Dokumen RPS / Kurikulum (PDF)", type=["pdf"]
    )
    process_btn = st.button("Proses & Indeks Dokumen RPS")

    if process_btn:
      if uploaded_file is None:
        st.error("⚠️ Upload file PDF terlebih dahulu.")
      else:
        with st.spinner("Memproses indeks dokumen RPS secara lokal..."):
          temp_path = "temp_curriculum.pdf"
          with open(temp_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
          try:
            st.session_state.retriever = initialize_local_vector_store(
                temp_path
            )
            st.session_state.uploaded_rps_name = uploaded_file.name
            st.success(f"✅ RPS '{uploaded_file.name}' berhasil diindeks!")
          except Exception as e:
            st.error(f"Gagal memproses dokumen lokal: {e}")

    generate_btn = st.button(
        "Generate Storyboard (Ground Truth Style)",
        type="primary",
        use_container_width=True,
    )

  # ==========================================
  # KOLOM 2: RAG CONTENT VIEWER & CONTEXT
  # ==========================================
  with col_middle:
    st.subheader("RAG Content Viewer & Context")
    search_kb = st.text_input(
        "Search your knowledge base", placeholder="🔍 Search..."
    )

    if generate_btn:
      if not program_studi or not nama_mata_kuliah or not learning_objectives:
        st.error(
            "⚠️ Harap isi Program Studi, Nama Mata Kuliah, dan Learning"
            " Objectives terlebih dahulu!"
        )
      elif st.session_state.retriever is None:
        st.error(
            "⚠️ Harap upload dan proses dokumen RPS terlebih dahulu di sidebar."
        )
      else:
        with st.spinner(
            "🔄 Menjalankan RAG dengan Standar Ground Truth..."
        ):
          try:
            relevant_docs = st.session_state.retriever.invoke(
                learning_objectives
            )
            st.session_state.rag_contexts = relevant_docs
            context_text = "\n\n".join(
                [doc.page_content for doc in relevant_docs]
            )

            if llm_provider == "Groq API (Cloud / Streamlit Cloud)":
              if not groq_api_key:
                st.error("⚠️ Masukkan Groq API Key terlebih dahulu di sidebar.")
                st.stop()
              llm = ChatGroq(groq_api_key=groq_api_key, model_name=groq_model)
            else:
              llm = Ollama(model=local_model_name, temperature=0.7)

            prompt = (
                "SYSTEM: Anda adalah API generator JSON murni. Tugas Anda adalah"
                " mengembalikan HANYA valid JSON array tanpa teks pengantar,"
                " tanpa penjelasan, dan tanpa markdown block.\n\n"
                "Berdasarkan dokumen kurikulum referensi berikut, buatlah"
                " rancangan storyboard pembelajaran yang komprehensif.\n"
                "Sertakan mapping Learning Objective ke setiap scene menggunakan format tag 'LO1', 'LO2', 'LO3', dst.\n\n"
                "DOKUMEN KURIKULUM REFERENSI:\n"
                f"{context_text}\n\n"
                f"Program Studi: {program_studi}\n"
                f"Mata Kuliah: {nama_mata_kuliah}\n"
                f"Target Learning Objective: {learning_objectives}\n"
                f"Visual Style: {visual_style}\n"
                f"Target Audience: {target_audience}\n\n"
                "ATURAN MUTLAK:\n"
                "1. Output HARUS berupa JSON array (list of objects) yang valid.\n"
                "2. Jangan tulis kata-kata pengantar. Langsung mulai dengan karakter '[' dan akhiri dengan ']'.\n"
                "3. Gunakan kunci JSON persis seperti ini:\n"
                "[\n"
                "  {\n"
                '    "judul_scene": "01 - Pengenalan Konsep",\n'
                '    "learning_objective": "LO1",\n'
                '    "visualisasi": "Ilustrasi diagram blok...",\n'
                '    "instruksi_visual": "Zoom in ke bagian matriks",\n'
                '    "animasi": "Fade in / slide from left",\n'
                '    "on_screen_text": "Rumus Utama CNN",\n'
                '    "voice_over_text": "Pernahkah kalian berpikir...",\n'
                '    "backsound": "Ambient Tech - Calm",\n'
                '    "durasi": "00:15",\n'
                '    "saran_info": "Pastikan visual matriks jelas"\n'
                "  }\n"
                "]"
            )

            if llm_provider == "Groq API (Cloud / Streamlit Cloud)":
              response = llm.invoke(prompt)
              resp_text = (
                  response.content
                  if hasattr(response, "content")
                  else str(response)
              )
            else:
              response = llm.invoke(prompt)
              resp_text = str(response)

            scenes_json = parse_llm_json_response(resp_text)

            # Evaluasi Alignment Learning Objective untuk disimpan ke DB
            alignment_prompt = (
                "SYSTEM: Anda adalah evaluator RAG dan kurikulum pendidikan.\n"
                "Berdasarkan dokumen kurikulum referensi dan Learning Objective yang ditentukan, "
                "berikan evaluasi singkat berupa 'Learning Objective Alignment Evaluation' "
                "apakah rancangan scene storyboard ini sudah selaras, serta berikan skor kecocokan (0-100%).\n\n"
                f"Target Learning Objective: {learning_objectives}\n"
                f"Dokumen Referensi: {context_text[:1000]}"
            )
            alignment_response = llm.invoke(alignment_prompt)
            alignment_text = (
                alignment_response.content
                if hasattr(alignment_response, "content")
                else str(alignment_response)
            )

            if scenes_json and isinstance(scenes_json, list):
              formatted_scenes = []
              for item in scenes_json:
                formatted_scenes.append({
                    "Judul Scene": str(item.get("judul_scene", "Scene 01")),
                    "Learning Objective": str(
                        item.get("learning_objective", "LO1")
                    ),
                    "Visualisasi": str(item.get("visualisasi", "")),
                    "Instruksi untuk Visual": str(
                        item.get("instruksi_visual", "")
                    ),
                    "Animasi": str(item.get("animasi", "")),
                    "On-Screen Text": str(item.get("on_screen_text", "")),
                    "Voice Over Text": str(item.get("voice_over_text", "")),
                    "Backsound": str(item.get("backsound", "Audio Calm")),
                    "Durasi": str(item.get("durasi", "00:15")),
                    "Saran dan Info": str(item.get("saran_info", "")),
                })

              new_df = pd.DataFrame(formatted_scenes)
              st.session_state.storyboard_df = new_df

              if check_db_connection():
                conn = get_db_connection()
                cursor = conn.cursor()

                query_main = "INSERT INTO storyboards (username, program_studi, nama_mata_kuliah, project_name, learning_objectives, target_audience, visual_style, rps_filename, alignment_evaluation) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)"
                cursor.execute(
                    query_main,
                    (
                        st.session_state.username,
                        program_studi,
                        nama_mata_kuliah,
                        project_name,
                        learning_objectives,
                        target_audience,
                        visual_style,
                        st.session_state.uploaded_rps_name,
                        alignment_text,
                    ),
                )
                storyboard_id = cursor.lastrowid

                query_scene = "INSERT INTO storyboard_scenes (storyboard_id, judul_scene, learning_objective_tag, visualisasi, instruksi_visual, animasi, on_screen_text, voice_over_text, backsound, durasi, saran_info) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"
                for row in formatted_scenes:
                  cursor.execute(
                      query_scene,
                      (
                          storyboard_id,
                          row["Judul Scene"],
                          row["Learning Objective"],
                          row["Visualisasi"],
                          row["Instruksi untuk Visual"],
                          row["Animasi"],
                          row["On-Screen Text"],
                          row["Voice Over Text"],
                          row["Backsound"],
                          row["Durasi"],
                          row["Saran dan Info"],
                      ),
                  )

                conn.commit()
                cursor.close()
                conn.close()

                st.success(
                    "✨ Storyboard & referensi RPS berhasil direkam ke"
                    f" Railway MySQL! (ID: {storyboard_id})"
                )
              else:
                st.warning(
                    "⚠️ Berhasil digenerate, namun koneksi MySQL terputus."
                )
            else:
              st.error(
                  "Gagal memproses format JSON dari model LLM. Periksa respons"
                  " teks."
              )

          except Exception as e:
            st.error(f"Terjadi kesalahan saat pemrosesan LLM: {e}")

    if st.session_state.rag_contexts:
      for i, doc in enumerate(st.session_state.rag_contexts):
        st.markdown(
            f"""
                <div class="card-box">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <b>Konten Referensi #{i+1}</b>
                        <span style="background-color: #D69E2E; color: #003366; padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: bold;">Active</span>
                    </div>
                    <p style="font-size: 13px; color: #334155; margin-top: 5px;">
                        {doc.page_content[:180]}...
                    </p>
                    <div style="font-size: 11px; color: #64748B;">
                        <span>Source RPS: {st.session_state.uploaded_rps_name}</span>
                    </div>
                </div>
            """,
            unsafe_allow_html=True,
        )
    else:
      st.info(
          "Belum ada konteks aktif. Upload dokumen RPS PDF dan klik 'Generate"
          " Storyboard'."
      )

  # ==========================================
  # KOLOM 3: STORYBOARD EDITOR & MANAGEMENT
  # ==========================================
  with col_right:
    st.subheader("Storyboard Editor & Management")

    if "storyboard_df" not in st.session_state:
      st.session_state.storyboard_df = pd.DataFrame(columns=[
          "Judul Scene",
          "Learning Objective",
          "Visualisasi",
          "Instruksi untuk Visual",
          "Animasi",
          "On-Screen Text",
          "Voice Over Text",
          "Backsound",
          "Durasi",
          "Saran dan Info",
      ])

    st.markdown(
        f"📋 **Prodi:** {program_studi if program_studi else '-'} | **Mata"
        f" Kuliah:** {nama_mata_kuliah if nama_mata_kuliah else '-'} | 📄"
        f" **RPS:** `{st.session_state.uploaded_rps_name}`"
    )
    st.divider()

    updated_rows = []

    if st.session_state.storyboard_df.empty:
      st.info(
          "Belum ada scene yang dimuat atau digenerate. Klik '➕ Add New Scene'"
          " atau 'Generate Storyboard' untuk memulai."
      )
    else:
      for idx, row in st.session_state.storyboard_df.iterrows():
        with st.expander(
            f"🎬 {row.get('Judul Scene', f'Scene {idx+1}')} [Target:"
            f" {row.get('Learning Objective', 'LO1')}] (Durasi:"
            f" {row.get('Durasi', '00:15')})",
            expanded=False,
        ):
          c1, c2 = st.columns(2)
          with c1:
            new_judul = st.text_input(
                "Judul Scene", value=str(row["Judul Scene"]), key=f"jdl_{idx}"
            )
            new_lo = st.selectbox(
                "Learning Objective Target",
                ["LO1", "LO2", "LO3", "LO4", "LO5"],
                index=(
                    ["LO1", "LO2", "LO3", "LO4", "LO5"].index(
                        str(row["Learning Objective"])
                    )
                    if str(row["Learning Objective"])
                    in ["LO1", "LO2", "LO3", "LO4", "LO5"]
                    else 0
                ),
                key=f"lo_{idx}",
            )
            new_dur = st.text_input(
                "Durasi", value=str(row["Durasi"]), key=f"dur_{idx}"
            )
            new_anim = st.text_input(
                "Animasi", value=str(row["Animasi"]), key=f"anim_{idx}"
            )
            new_ost = st.text_input(
                "On-Screen Text",
                value=str(row["On-Screen Text"]),
                key=f"ost_{idx}",
            )
          with c2:
            new_bs = st.text_input(
                "Backsound", value=str(row["Backsound"]), key=f"bs_{idx}"
            )
            new_vis = st.text_area(
                "Visualisasi",
                value=str(row["Visualisasi"]),
                key=f"vis_{idx}",
                height=75,
            )
            new_inst = st.text_area(
                "Instruksi untuk Visualisasi",
                value=str(row["Instruksi untuk Visual"]),
                key=f"inst_{idx}",
                height=75,
            )
            new_vot = st.text_area(
                "Voice Over Text",
                value=str(row["Voice Over Text"]),
                key=f"vot_{idx}",
                height=75,
            )
            new_saran = st.text_area(
                "Saran : Info",
                value=str(row["Saran dan Info"]),
                key=f"saran_{idx}",
                height=75,
            )

          b_col1, b_col2 = st.columns(2)
          with b_col1:
            is_saved = st.button(
                "💾 Simpan Baris Ini",
                key=f"save_row_{idx}",
                use_container_width=True,
            )
          with b_col2:
            is_deleted = st.button(
                "🗑️ Hapus Baris Ini",
                key=f"del_row_{idx}",
                use_container_width=True,
            )

          if is_deleted:
            continue

          if is_saved:
            st.success(f"Scene '{new_judul}' berhasil diperbarui di memori!")

          updated_rows.append({
              "Judul Scene": new_judul,
              "Learning Objective": new_lo,
              "Visualisasi": new_vis,
              "Instruksi untuk Visual": new_inst,
              "Animasi": new_anim,
              "On-Screen Text": new_ost,
              "Voice Over Text": new_vot,
              "Backsound": new_bs,
              "Durasi": new_dur,
              "Saran dan Info": new_saran,
          })
        st.markdown("---")

      st.session_state.storyboard_df = pd.DataFrame(updated_rows)

    col_btn1, col_btn2 = st.columns([1, 1])
    with col_btn1:
      if st.button("➕ Add New Scene"):
        next_num = len(st.session_state.storyboard_df) + 1
        new_row_df = pd.DataFrame([{
            "Judul Scene": f"Scene {str(next_num).zfill(2)} - Judul Baru",
            "Learning Objective": "LO1",
            "Visualisasi": "Deskripsi visual...",
            "Instruksi untuk Visual": "Instruksi khusus...",
            "Animasi": "Fade In",
            "On-Screen Text": "Teks Layar...",
            "Voice Over Text": '"Naskah suara..."',
            "Backsound": "Audio...",
            "Durasi": "00:15",
            "Saran dan Info": "Catatan tambahan...",
        }])
        st.session_state.storyboard_df = pd.concat(
            [st.session_state.storyboard_df, new_row_df], ignore_index=True
        )
        st.rerun()

    st.divider()

    if st.session_state.storyboard_df.empty:
      st.warning("⚠️ Belum ada scene untuk diexport.")
    else:
      pptx_data = generate_pptx(
          st.session_state.storyboard_df,
          program_studi,
          nama_mata_kuliah,
          project_name,
      )
      st.download_button(
          label="📥 Export to PPT",
          data=pptx_data,
          file_name=(
              f"{project_name.replace(' ', '_')}_Storyboard.pptx"
              if project_name
              else "Storyboard.pptx"
          ),
          mime=(
              "application/vnd.openxmlformats-officedocument.presentationml.presentation"
          ),
          use_container_width=True,
      )

    # ==========================================
    # TABEL DAFTAR RIWAYAT PROYEK USER
    # ==========================================
    st.markdown("---")
    with st.expander("📂 Riwayat Proyek Storyboard Anda", expanded=False):
      if check_db_connection():
        try:
          conn = get_db_connection()
          query_history = (
              "SELECT id, program_studi, nama_mata_kuliah, project_name,"
              " rps_filename, created_at FROM storyboards WHERE username = %s"
              " ORDER BY id DESC"
          )
          df_history = pd.read_sql(
              query_history, conn, params=(st.session_state.username,)
          )
          conn.close()

          if not df_history.empty:
            df_display = df_history.drop(columns=["id"])

            event = st.dataframe(
                df_display,
                use_container_width=True,
                hide_index=True,
                on_select="rerun",
                selection_mode="single-row",
            )

            selected_rows = event.selection.rows
            if selected_rows:
              selected_index = selected_rows[0]
              selected_proj_id = int(df_history.iloc[selected_index]["id"])
              selected_proj_name = df_history.iloc[selected_index][
                  "project_name"
              ]
              selected_rps = df_history.iloc[selected_index]["rps_filename"]

              if st.button(
                  f"📂 Muat Proyek '{selected_proj_name}' (ID:"
                  f" {selected_proj_id}) ke Editor"
              ):
                conn = get_db_connection()
                query_load_scenes = """
                                    SELECT judul_scene, learning_objective_tag, visualisasi, instruksi_visual, animasi, 
                                           on_screen_text, voice_over_text, backsound, durasi, saran_info 
                                    FROM storyboard_scenes WHERE storyboard_id = %s
                                """
                df_scenes_loaded = pd.read_sql(
                    query_load_scenes, conn, params=(selected_proj_id,)
                )
                conn.close()

                if not df_scenes_loaded.empty:
                  df_scenes_loaded.rename(
                      columns={
                          "judul_scene": "Judul Scene",
                          "learning_objective_tag": "Learning Objective",
                          "visualisasi": "Visualisasi",
                          "instruksi_visual": "Instruksi untuk Visual",
                          "animasi": "Animasi",
                          "on_screen_text": "On-Screen Text",
                          "voice_over_text": "Voice Over Text",
                          "backsound": "Backsound",
                          "durasi": "Durasi",
                          "saran_info": "Saran dan Info",
                      },
                      inplace=True,
                  )

                  st.session_state.storyboard_df = df_scenes_loaded
                  st.session_state.uploaded_rps_name = selected_rps
                  st.success(
                      f"✨ Proyek '{selected_proj_name}' (Ref RPS:"
                      f" {selected_rps}) berhasil dimuat ke editor!"
                  )
                  st.rerun()
                else:
                  st.warning(
                      "⚠️ Tidak ada data scene yang ditemukan untuk proyek"
                      " ini."
                  )
          else:
            st.info("Belum ada riwayat proyek tersimpan untuk akun Anda.")
        except Exception as e:
          st.info(f"Riwayat proyek belum tersedia: {e}")
