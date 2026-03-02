
# =========================================================
# SICOT — Sistema de Cruzamento Operacional Telefônico e Telemático
# Versão 1.0.3 (build: 2026-03-02)
# =========================================================

import streamlit as st
import os
import base64

APP_NAME = "SICOT"
APP_SUBTITLE = "Sistema de Cruzamento Operacional Telefônico e Telemático"
APP_VERSION = "Versão 1.0.3"

DEV_NAME = "Marcos Da Silva Machado"
DEV_CONTACT = "agente.vcz@gmail.com"


def _find_first_existing(paths):
    for p in paths:
        try:
            if p and os.path.exists(p):
                return p
        except Exception:
            continue
    return None


def apply_watermark(
    preferred_path="banner_operacional.png",
    overlay_alpha=0.90,  # 0.90 => imagem ~10% visível
    blur_px=0,
):
    candidates = [
        preferred_path,
        "banner_operacional.png.png",
        "ChatGPT Image 26 de fev. de 2026.png",
        r"C:\SICOT_2\banner_operacional.png",
        r"C:\SICOT\banner_operacional.png",
    ]
    img_path = _find_first_existing(candidates)
    if not img_path:
        st.session_state["_sicot_watermark_missing"] = True
        return

    with open(img_path, "rb") as f:
        encoded = base64.b64encode(f.read()).decode()

    blur_css = (
        f"backdrop-filter: blur({int(blur_px)}px); -webkit-backdrop-filter: blur({int(blur_px)}px);"
        if blur_px else ""
    )

    st.markdown(
        f"""
        <style>
        .stApp, .stAppViewContainer {{
            background-image: url("data:image/png;base64,{encoded}");
            background-size: cover;
            background-position: center;
            background-attachment: fixed;
        }}

        .stApp::before, .stAppViewContainer::before {{
            content: "";
            position: fixed;
            inset: 0;
            {blur_css}
            background: rgba(17, 24, 39, {overlay_alpha});
            z-index: 0;
            pointer-events: none;
        }}

        .block-container {{
            position: relative;
            z-index: 1;
        }}

        header[data-testid="stHeader"] {{
            background: transparent;
        }}
        </style>
        """,
        unsafe_allow_html=True
    )


def render_header_and_sidebar():
    col1, col2 = st.columns([1, 4], vertical_alignment="center")
    with col1:
        if os.path.exists("sicot_logo.png"):
            st.image("sicot_logo.png", width=170)
    with col2:
        st.markdown(f"# {APP_NAME}")
        st.markdown(f"### {APP_SUBTITLE}")
        st.caption(f"{APP_VERSION}  |  Análise • Cruzamento • Ranking • Exportação")
    st.divider()

    with st.sidebar:
        st.markdown("## SICOT")
        st.write("Ferramenta destinada à análise operacional de interceptações telefônicas e telemáticas.")
        st.divider()

        st.markdown("### Versão")
        st.write(APP_VERSION)
        st.divider()

        st.markdown("### Desenvolvedor")
        st.write(DEV_NAME)
        st.write(DEV_CONTACT)
        st.divider()

        if st.session_state.get("_sicot_watermark_missing"):
            st.warning(
                "Marca d'água não aplicada: arquivo do banner não encontrado.\n\n"
                "Coloque **banner_operacional.png** na mesma pasta do .py (ou ajuste o caminho).",
                icon="⚠️"
            )

        st.caption("Uso destinado a atividades institucionais.")


def configure_page():
    st.set_page_config(page_title=APP_NAME, page_icon="📊", layout="wide")


def main():
    configure_page()
    apply_watermark(preferred_path="banner_operacional.png", overlay_alpha=0.90, blur_px=0)
    render_header_and_sidebar()
    run_sicot()


# =========================================================
# SISTEMA (código original, modularizado)
# =========================================================

# app.py
# Streamlit - Histórico Chamadas VIVO (XLSX) + Histórico WhatsApp (XLSX)
# + Cadastros (VIVO TXT / CLARO PDF / TIM PDF)
#
# WHATSAPP (sem alvo): cruza TODOS os telefones encontrados em Remetente + Destinatário(s)
# com os cadastros e exibe todos os identificados (sem direção enviada/recebida).
#
# EXTRA:
# - Exporta TXT com números NÃO identificados (Chamadas + WhatsApp), 1 por linha (ex.: 14981386337)
# - TAB "Estatísticas": top números mais frequentes (Chamadas e WhatsApp) com NOME (se identificado)
# - Gráfico pizza: Identificados vs Não identificados (Chamadas e WhatsApp)
#
# Sem warnings:
# - Streamlit: usa width="stretch" (sem use_container_width)
# - openpyxl: suprime warning "no default style" (arquivos de terceiros)
#
# Requisitos:
#   pip install streamlit pandas openpyxl pymupdf matplotlib
#
# Executar:
#   streamlit run app.py

import io
import re
import time
import warnings
from typing import Any, Dict, Optional, Tuple, List

import pandas as pd
import streamlit as st
import fitz  # PyMuPDF
import openpyxl
import matplotlib.pyplot as plt

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


# =========================================================
# UTILITÁRIOS
# =========================================================

def only_digits(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return re.sub(r"\D+", "", str(value))


def clean_field(value: Any) -> Optional[str]:
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    s = str(value).strip()
    s = s.rstrip("*").strip()
    return s or None


def format_doc_br(value: Any) -> Optional[str]:
    raw = clean_field(value)
    d = only_digits(raw)
    if len(d) == 11:  # CPF
        return f"{d[0:3]}.{d[3:6]}.{d[6:9]}-{d[9:11]}"
    if len(d) == 14:  # CNPJ
        return f"{d[0:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:14]}"
    return raw


def normalize_phone_br(value: Any) -> str:
    """
    Normalização única (cadastros + históricos):
    - remove tudo que não é dígito
    - remove '55' se vier com DDI (ex.: 55DDDNÚMERO)
    - aceita somente 10 (fixo c/ DDD) ou 11 (celular c/ DDD)
    - caso contrário retorna "" (descarta)
    """
    d = only_digits(value)
    if d.startswith("55") and len(d) >= 12:
        d = d[2:]
    if len(d) in (10, 11):
        return d
    return ""


def pdf_bytes_to_text(pdf_bytes: bytes) -> str:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    raw = "\n".join(page.get_text("text") for page in doc)
    doc.close()
    return raw


def dict_base_to_df(base: Dict[str, Tuple[float, Dict[str, Optional[str]]]]) -> pd.DataFrame:
    if not base:
        return pd.DataFrame(columns=["Telefone", "Nome", "CPF", "CIDADE", "OPERADORA", "Arquivo"])
    return pd.DataFrame([v[1] for v in base.values()])


# =========================================================
# PARSER VIVO TXT (robusto)
# =========================================================

RE_TXT_PHONE_LINE = re.compile(r"N[ÚU]MERO DA LINHA:.*\((\d{2})\)\s*([0-9\- ]+)")
RE_TXT_NAME = re.compile(r"CLIENTE:\.*\s*(.+)$")
RE_TXT_CPF = re.compile(r"CPF:\.*\s*([0-9\.\-]+)")
RE_TXT_CITY = re.compile(r"MUNIC[ÍI]PIO:\.*\s*(.+)$")


def parse_txt(text: str, file_name: str, mtime: float, operadora: str) -> Dict[str, Tuple[float, Dict[str, Optional[str]]]]:
    lines = text.splitlines()
    out: Dict[str, Tuple[float, Dict[str, Optional[str]]]] = {}

    i = 0
    while i < len(lines):
        m = RE_TXT_PHONE_LINE.search(lines[i])
        if not m:
            i += 1
            continue

        ddd = m.group(1)
        phone_raw = m.group(2)
        telefone = normalize_phone_br(ddd + phone_raw)

        window = lines[i : min(i + 60, len(lines))]
        nome = cpf = cidade = None

        for w in window:
            if nome is None:
                mn = RE_TXT_NAME.search(w)
                if mn:
                    nome = clean_field(mn.group(1))
            if cpf is None:
                mc = RE_TXT_CPF.search(w)
                if mc:
                    cpf = clean_field(mc.group(1))
            if cidade is None:
                mci = RE_TXT_CITY.search(w)
                if mci:
                    cidade = clean_field(mci.group(1))
            if nome and cpf and cidade:
                break

        cpf = format_doc_br(cpf)

        if telefone:
            out[telefone] = (
                mtime,
                {
                    "Telefone": telefone,
                    "Nome": nome,
                    "CPF": cpf,
                    "CIDADE": cidade,
                    "OPERADORA": operadora,
                    "Arquivo": file_name,
                },
            )
        i += 10

    return out


def build_base_vivo_from_uploads(uploaded_files) -> Dict[str, Tuple[float, Dict[str, Optional[str]]]]:
    base: Dict[str, Tuple[float, Dict[str, Optional[str]]]] = {}
    if not uploaded_files:
        return base

    for f in uploaded_files:
        raw = f.getvalue()
        text = raw.decode("utf-8", errors="replace")
        if text.count("�") > 10:
            text = raw.decode("cp1252", errors="replace")

        mtime = time.time()
        parsed = parse_txt(text, f.name, mtime, "VIVO")

        for tel, (ts, rec) in parsed.items():
            old = base.get(tel)
            if old is None or ts > old[0]:
                base[tel] = (ts, rec)

    return base


# =========================================================
# PARSER CLARO PDF
# =========================================================

RE_PDF_CLARO_ROW = re.compile(
    r"MSISDNs\s+"
    r"(?P<nome>.+?)\s+"
    r"(?P<data>\d{2}/\d{2}/\d{4})\s*[–-]\s*\d{2}:\d{2}:\s*\d{2}\s+"
    r"(?P<msisdn>55\d{11})\s+"
    r"(?P<contato>\d{10,11})\s+"
    r"(?P<doc>\d{11,14})\s+"
    r"(?P<tipo>[A-Z]+)\s+"
    r"(?P<end>.+)$"
)


def extract_city_from_address(addr: Optional[str]) -> Optional[str]:
    if not addr:
        return None
    m = re.search(r",\s*([^,]+?)\s*-\s*[A-Z]{2}\b", addr)
    if m:
        return m.group(1).strip()
    m = re.search(r",\s*([^,]+?)\s*[A-Z]{2}\b", addr)
    if m:
        return m.group(1).strip()
    return None


def build_base_claro_from_uploads(uploaded_files) -> Dict[str, Tuple[float, Dict[str, Optional[str]]]]:
    base: Dict[str, Tuple[float, Dict[str, Optional[str]]]] = {}
    if not uploaded_files:
        return base

    for f in uploaded_files:
        mtime = time.time()
        raw_text = pdf_bytes_to_text(f.getvalue())
        norm = re.sub(r"\s+", " ", raw_text).strip()

        matches = list(RE_PDF_CLARO_ROW.finditer(norm))
        if not matches:
            continue

        m = matches[-1]
        nome = clean_field(m.group("nome"))
        msisdn = normalize_phone_br(m.group("msisdn"))
        contato = normalize_phone_br(m.group("contato"))
        docnum = format_doc_br(m.group("doc"))
        endereco = clean_field(m.group("end"))

        telefone_key = contato if contato else msisdn
        cidade = extract_city_from_address(endereco)

        if telefone_key:
            old = base.get(telefone_key)
            if old is None or mtime > old[0]:
                base[telefone_key] = (
                    mtime,
                    {
                        "Telefone": telefone_key,
                        "Nome": nome,
                        "CPF": docnum,
                        "CIDADE": cidade,
                        "OPERADORA": "CLARO",
                        "Arquivo": f.name,
                    },
                )

    return base


# =========================================================
# PARSER TIM PDF
# =========================================================

RE_TIM_LINE = re.compile(r"N[ÚU]MERO DA LINHA:\s*([0-9\.\-\s]+)")
RE_TIM_NAME = re.compile(r"NOME:\s*(.+?)\s+(?:a a|$)")
RE_TIM_DOC = re.compile(r"CPF/CNPJ:\s*([0-9\.\-\/]+)")
RE_TIM_CITY_FAT = re.compile(r"CIDADE/UF\s*-\s*CEP\s*FATURA:\s*([A-ZÀ-Üa-zà-ü\s]+\/[A-Z]{2})")
RE_TIM_CITY_RES = re.compile(r"CIDADE/UF\s*-\s*CEP\s*RESIDENCIAL:\s*([A-ZÀ-Üa-zà-ü\s]+\/[A-Z]{2})")


def build_base_tim_from_uploads(uploaded_files) -> Dict[str, Tuple[float, Dict[str, Optional[str]]]]:
    base: Dict[str, Tuple[float, Dict[str, Optional[str]]]] = {}
    if not uploaded_files:
        return base

    for f in uploaded_files:
        mtime = time.time()
        raw_text = pdf_bytes_to_text(f.getvalue())
        norm = re.sub(r"\s+", " ", raw_text).strip()

        mline = RE_TIM_LINE.search(norm)
        if not mline:
            continue

        telefone = normalize_phone_br(mline.group(1))

        mname = RE_TIM_NAME.search(norm)
        mdoc = RE_TIM_DOC.search(norm)
        mfat = RE_TIM_CITY_FAT.search(norm)
        mres = RE_TIM_CITY_RES.search(norm)

        nome = clean_field(mname.group(1)) if mname else None
        cpf = format_doc_br(mdoc.group(1)) if mdoc else None

        cidade_uf = None
        if mfat and clean_field(mfat.group(1)) not in (None, ".", ""):
            cidade_uf = mfat.group(1).strip()
        elif mres and clean_field(mres.group(1)) not in (None, ".", ""):
            cidade_uf = mres.group(1).strip()

        cidade = None
        if cidade_uf and "/" in cidade_uf:
            cidade = cidade_uf.split("/")[0].strip()

        if telefone:
            old = base.get(telefone)
            if old is None or mtime > old[0]:
                base[telefone] = (
                    mtime,
                    {
                        "Telefone": telefone,
                        "Nome": nome,
                        "CPF": cpf,
                        "CIDADE": cidade,
                        "OPERADORA": "TIM",
                        "Arquivo": f.name,
                    },
                )

    return base


# =========================================================
# HISTÓRICO CHAMADAS VIVO XLSX (linha alvo e período em B4)
# =========================================================

RE_B4 = re.compile(
    r"Linha:(?P<linha>\d+)\s+Per[ií]odo:(?P<ini>\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2})\s+a\s+(?P<fim>\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2})",
    re.IGNORECASE,
)


def extrair_linha_e_periodo_b4(xlsx_bytes: bytes, sheet_name: str) -> Tuple[str, Optional[pd.Timestamp], Optional[pd.Timestamp]]:
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    if sheet_name not in wb.sheetnames:
        return "", None, None
    sh = wb[sheet_name]
    b4 = str(sh["B4"].value or "")
    m = RE_B4.search(b4)
    if not m:
        return "", None, None

    linha = m.group("linha").strip()
    ini = pd.to_datetime(m.group("ini"), format="%d/%m/%Y %H:%M:%S", errors="coerce")
    fim = pd.to_datetime(m.group("fim"), format="%d/%m/%Y %H:%M:%S", errors="coerce")
    return linha, ini, fim


def carregar_historico_chamadas_vivo(xlsx_bytes: bytes, sheet_name: str, header_linha_1based: int) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    linha_alvo, dt_ini, dt_fim = extrair_linha_e_periodo_b4(xlsx_bytes, sheet_name)

    header = max(0, int(header_linha_1based) - 1)
    df = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=sheet_name, header=header, dtype=str)

    if "Chamador" in df.columns:
        df["Chamador"] = df["Chamador"].apply(normalize_phone_br)
    if "Chamado" in df.columns:
        df["Chamado"] = df["Chamado"].apply(normalize_phone_br)

    if "Data" in df.columns and "Hora" in df.columns:
        df["data_hora_inicio"] = pd.to_datetime(
            df["Data"].fillna("").astype(str) + " " + df["Hora"].fillna("").astype(str),
            format="%d/%m/%Y %H:%M:%S",
            errors="coerce",
        )
    else:
        df["data_hora_inicio"] = pd.NaT

    def direcao(row) -> str:
        c1 = row.get("Chamador", "") or ""
        c2 = row.get("Chamado", "") or ""
        if not linha_alvo:
            return "DESCONHECIDA"
        if c1 == linha_alvo:
            return "ORIGINADA"
        if c2 == linha_alvo:
            return "RECEBIDA"
        return "TERCEIROS"

    def outro(row) -> str:
        c1 = row.get("Chamador", "") or ""
        c2 = row.get("Chamado", "") or ""
        if not linha_alvo:
            return ""
        if c1 == linha_alvo:
            return c2
        if c2 == linha_alvo:
            return c1
        return ""

    df["direcao"] = df.apply(direcao, axis=1)
    df["outro_numero"] = df.apply(outro, axis=1).apply(normalize_phone_br)

    if dt_ini is not None and dt_fim is not None and "data_hora_inicio" in df.columns:
        df = df[(df["data_hora_inicio"] >= dt_ini) & (df["data_hora_inicio"] <= dt_fim)]

    meta = {"linha_alvo": linha_alvo, "periodo_inicio": dt_ini, "periodo_fim": dt_fim}
    return df, meta


# =========================================================
# HISTÓRICO WHATSAPP XLSX (SEM ALVO)
# =========================================================

PHONE_TOKEN_RE = re.compile(r"(?:\+?55\s*)?\(?\d{2}\)?\s*\d{4,5}[-\s]?\d{4}")


def extrair_telefones_de_texto(texto: str) -> List[str]:
    if not texto:
        return []
    achados = PHONE_TOKEN_RE.findall(texto)
    nums = [normalize_phone_br(a) for a in achados]
    return [n for n in nums if n]


def parse_destinatarios(valor: Any) -> List[str]:
    s = str(valor or "")
    nums = extrair_telefones_de_texto(s)

    if not nums:
        n = normalize_phone_br(s)
        if n:
            nums = [n]

    out, seen = [], set()
    for n in nums:
        if n not in seen:
            out.append(n)
            seen.add(n)
    return out


def carregar_historico_whatsapp(xlsx_bytes: bytes, sheet_name: str) -> pd.DataFrame:
    df = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=sheet_name, dtype=str)

    if "Data" in df.columns and "Hora" in df.columns:
        df["data_hora"] = pd.to_datetime(
            df["Data"].fillna("").astype(str) + " " + df["Hora"].fillna("").astype(str),
            format="%d/%m/%Y %H:%M:%S",
            errors="coerce",
        )
    else:
        df["data_hora"] = pd.NaT

    df["remetente_norm"] = df.get("Remetente", "").astype(str).apply(normalize_phone_br)
    df["dest_list"] = df.get("Destinatário(s)", "").apply(parse_destinatarios)

    df = df.explode("dest_list").rename(columns={"dest_list": "destinatario_norm"}).copy()
    df["destinatario_norm"] = df["destinatario_norm"].fillna("").astype(str)

    return df


# =========================================================
# CRUZAMENTOS / EXPORTS / ESTATÍSTICAS
# =========================================================

def cruzar_whatsapp_sem_alvo(df_msgs: pd.DataFrame, df_cad: pd.DataFrame) -> pd.DataFrame:
    nums = set()

    if "remetente_norm" in df_msgs.columns:
        nums |= set([n for n in df_msgs["remetente_norm"].fillna("").tolist() if n])
    if "destinatario_norm" in df_msgs.columns:
        nums |= set([n for n in df_msgs["destinatario_norm"].fillna("").tolist() if n])

    df_hist = pd.DataFrame({"Telefone": sorted(nums)})
    df_hist["Telefone"] = df_hist["Telefone"].apply(normalize_phone_br)
    df_hist = df_hist[df_hist["Telefone"] != ""].drop_duplicates()

    base = df_cad.copy()
    base["Telefone"] = base["Telefone"].apply(normalize_phone_br)
    base = base[base["Telefone"] != ""].drop_duplicates("Telefone")

    df_ident = df_hist.merge(
        base[["Telefone", "Nome", "CPF", "CIDADE", "OPERADORA", "Arquivo"]],
        on="Telefone",
        how="left",
    )

    return df_ident[df_ident["Nome"].notna() | df_ident["CPF"].notna() | df_ident["CIDADE"].notna()].copy()


def gerar_txt_nao_identificados(df_cad: pd.DataFrame, df_calls: Optional[pd.DataFrame], df_msgs: Optional[pd.DataFrame]) -> Tuple[bytes, int]:
    known = set(df_cad["Telefone"].fillna("").astype(str).apply(normalize_phone_br).tolist())
    known = {n for n in known if n}

    unknown = set()

    if df_calls is not None and "outro_numero" in df_calls.columns:
        nums_calls = df_calls["outro_numero"].fillna("").astype(str).apply(normalize_phone_br).tolist()
        unknown |= {n for n in nums_calls if n and n not in known}

    if df_msgs is not None:
        if "remetente_norm" in df_msgs.columns:
            nums_r = df_msgs["remetente_norm"].fillna("").astype(str).apply(normalize_phone_br).tolist()
            unknown |= {n for n in nums_r if n and n not in known}
        if "destinatario_norm" in df_msgs.columns:
            nums_d = df_msgs["destinatario_norm"].fillna("").astype(str).apply(normalize_phone_br).tolist()
            unknown |= {n for n in nums_d if n and n not in known}

    unknown_sorted = sorted(unknown)
    txt = ("\n".join(unknown_sorted) + ("\n" if unknown_sorted else "")).encode("utf-8")
    return txt, len(unknown_sorted)


def montar_top_frequencias(df_calls: Optional[pd.DataFrame], df_msgs: Optional[pd.DataFrame], df_cad: Optional[pd.DataFrame], top_n: int) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Retorna:
      - top_calls: Nome + Telefone + Frequencia + Identificado
      - top_wa:    Nome + Telefone + Frequencia + Identificado
    """
    cad_map = None
    known = set()

    if df_cad is not None and "Telefone" in df_cad.columns:
        tmp = df_cad.copy()
        tmp["Telefone"] = tmp["Telefone"].fillna("").astype(str).apply(normalize_phone_br)
        tmp = tmp[tmp["Telefone"] != ""].drop_duplicates("Telefone")
        cad_map = tmp[["Telefone", "Nome"]].copy()
        cad_map["Nome"] = cad_map["Nome"].fillna("").astype(str)
        known = set(cad_map["Telefone"].tolist())

    # Chamadas: conta outro_numero
    if df_calls is not None and "outro_numero" in df_calls.columns:
        s = df_calls["outro_numero"].fillna("").astype(str).apply(normalize_phone_br)
        s = s[s != ""]
        vc = s.value_counts().head(top_n)
        top_calls = pd.DataFrame({"Telefone": vc.index.tolist(), "Frequencia": vc.values.tolist()})
    else:
        top_calls = pd.DataFrame(columns=["Telefone", "Frequencia"])

    # WhatsApp: conta remetente_norm + destinatario_norm
    nums = []
    if df_msgs is not None:
        if "remetente_norm" in df_msgs.columns:
            nums.extend(df_msgs["remetente_norm"].fillna("").astype(str).apply(normalize_phone_br).tolist())
        if "destinatario_norm" in df_msgs.columns:
            nums.extend(df_msgs["destinatario_norm"].fillna("").astype(str).apply(normalize_phone_br).tolist())

    if nums:
        s2 = pd.Series([n for n in nums if n], dtype="object")
        vc2 = s2.value_counts().head(top_n)
        top_wa = pd.DataFrame({"Telefone": vc2.index.tolist(), "Frequencia": vc2.values.tolist()})
    else:
        top_wa = pd.DataFrame(columns=["Telefone", "Frequencia"])

    # Marca identificado + inclui NOME (se existir no cadastro)
    if not top_calls.empty:
        top_calls["Identificado"] = top_calls["Telefone"].apply(lambda x: x in known)
        if cad_map is not None:
            top_calls = top_calls.merge(cad_map, on="Telefone", how="left")
        else:
            top_calls["Nome"] = ""
        top_calls["Nome"] = top_calls["Nome"].fillna("").astype(str)
        top_calls = top_calls[["Nome", "Telefone", "Frequencia", "Identificado"]]

    if not top_wa.empty:
        top_wa["Identificado"] = top_wa["Telefone"].apply(lambda x: x in known)
        if cad_map is not None:
            top_wa = top_wa.merge(cad_map, on="Telefone", how="left")
        else:
            top_wa["Nome"] = ""
        top_wa["Nome"] = top_wa["Nome"].fillna("").astype(str)
        top_wa = top_wa[["Nome", "Telefone", "Frequencia", "Identificado"]]

    return top_calls, top_wa


def build_known_set(df_cad: Optional[pd.DataFrame]) -> set:
    if df_cad is None or "Telefone" not in df_cad.columns:
        return set()
    known = set(df_cad["Telefone"].fillna("").astype(str).apply(normalize_phone_br).tolist())
    return {n for n in known if n}


# =========================================================

def run_sicot():
    # STREAMLIT UI
    # =========================================================

    st.title("Históricos (Chamadas VIVO + WhatsApp) + Cadastros (Vivo/Claro/Tim) + Exportação Excel/TXT")

    if "df_calls" not in st.session_state:
        st.session_state["df_calls"] = None
    if "meta_calls" not in st.session_state:
        st.session_state["meta_calls"] = {}
    if "df_msgs" not in st.session_state:
        st.session_state["df_msgs"] = None
    if "df_cad" not in st.session_state:
        st.session_state["df_cad"] = None
    if "df_ident_calls" not in st.session_state:
        st.session_state["df_ident_calls"] = None
    if "df_ident_wa" not in st.session_state:
        st.session_state["df_ident_wa"] = None
    if "xlsx_bytes" not in st.session_state:
        st.session_state["xlsx_bytes"] = b""
    if "unknown_txt_bytes" not in st.session_state:
        st.session_state["unknown_txt_bytes"] = b""
    if "unknown_count" not in st.session_state:
        st.session_state["unknown_count"] = 0

    tabs = st.tabs(["Operação", "Estatísticas"])


    # =========================================================
    # TAB 1 - OPERAÇÃO
    # =========================================================
    with tabs[0]:
        st.subheader("1) Histórico de Chamadas (VIVO - XLSX)")
        hist_xlsx = st.file_uploader("Anexar XLSX do histórico de chamadas", type=["xlsx"], key="hist_xlsx")

        c1, c2 = st.columns(2)
        with c1:
            sheet_name_calls = st.text_input("Nome da aba do relatório (Chamadas)", value="Relatório de chamadas ")
        with c2:
            header_linha_calls = st.number_input("Linha do cabeçalho (1-based)", min_value=1, max_value=60, value=6, step=1)

        if st.button("Carregar Chamadas", type="primary"):
            if not hist_xlsx:
                st.error("Envie o XLSX do histórico de chamadas.")
            else:
                xbytes = hist_xlsx.getvalue()
                df_calls, meta = carregar_historico_chamadas_vivo(
                    xbytes,
                    sheet_name=sheet_name_calls,
                    header_linha_1based=int(header_linha_calls),
                )
                st.session_state["df_calls"] = df_calls
                st.session_state["meta_calls"] = meta
                st.success(f"Chamadas carregadas. Registros: {len(df_calls)}")
                st.info(
                    f"Linha alvo (Chamadas): {meta.get('linha_alvo') or '(não detectada)'} | "
                    f"Período: {meta.get('periodo_inicio')} a {meta.get('periodo_fim')}"
                )

        if st.session_state["df_calls"] is not None:
            st.subheader("Pré-visualização (Chamadas)")
            dfv = st.session_state["df_calls"].copy()
            cols_pref = [
                "data_hora_inicio", "direcao",
                "Chamador", "Chamado", "outro_numero",
                "duracao_s", "Status", "Tra",
                "Local Origem", "Local Destino",
                "Tec Origem", "Tec Destino",
            ]
            cols_show = [c for c in cols_pref if c in dfv.columns]
            st.write(f"Linhas: {len(dfv)}")
            st.dataframe(dfv[cols_show] if cols_show else dfv, width="stretch", height=420)


    st.subheader("2) Histórico WhatsApp (XLSX - sem alvo)")

    with st.expander("Pré-processamento obrigatório (WhatsApp) — gerar XLSX para o SICOT", expanded=True):
        st.warning(
            "Antes de anexar aqui, processe o arquivo de interceptação do WhatsApp no site abaixo. "
            "Após o processamento, baixe o **XLSX gerado** e só então faça o upload no SICOT.",
            icon="⚠️"
        )
        st.link_button(
            "Abrir processador de interceptação WhatsApp",
            "https://italofds.github.io/whatsapp-extract-processor/"
        )
        st.caption("Ferramenta de pré-processamento desenvolvida por Ítalo Santos — italofds.dev@gmail.com — https://github.com/italofds")

        wa_xlsx = st.file_uploader("Anexar XLSX do WhatsApp", type=["xlsx"], key="wa_xlsx")
        wa_sheet_name = st.text_input("Nome da aba (WhatsApp)", value="Mensagens")

        if st.button("Carregar WhatsApp"):
            if not wa_xlsx:
                st.error("Envie o XLSX do WhatsApp.")
            else:
                df_msgs = carregar_historico_whatsapp(wa_xlsx.getvalue(), sheet_name=wa_sheet_name)
                st.session_state["df_msgs"] = df_msgs
                st.success(f"WhatsApp carregado. Registros (após explode destinatários): {len(df_msgs)}")

        if st.session_state["df_msgs"] is not None:
            st.subheader("Pré-visualização (WhatsApp)")
            dfm = st.session_state["df_msgs"].copy()
            cols_pref = [
                "data_hora",
                "Data", "Hora",
                "Remetente", "Destinatário(s)",
                "remetente_norm", "destinatario_norm",
                "ID do Grupo",
                "Endereço IP (Remetente)", "Porta Lógica (Remetente)",
                "Cidade", "Dispositivo", "Tipo",
            ]
            cols_show = [c for c in cols_pref if c in dfm.columns]
            st.write(f"Linhas: {len(dfm)}")
            st.dataframe(dfm[cols_show] if cols_show else dfm, width="stretch", height=420)

        st.subheader("3) Dados cadastrais (uploads)")
        vivo_txts = st.file_uploader("VIVO (TXT) - múltiplos", type=["txt"], accept_multiple_files=True, key="vivo_txts")
        claro_pdfs = st.file_uploader("CLARO (PDF) - múltiplos", type=["pdf"], accept_multiple_files=True, key="claro_pdfs")
        tim_pdfs = st.file_uploader("TIM (PDF) - múltiplos", type=["pdf"], accept_multiple_files=True, key="tim_pdfs")

        debug = st.checkbox("Mostrar debug de extração", value=True)

        if st.button("Processar Cadastros e Cruzar (Chamadas + WhatsApp)"):
            base_vivo = build_base_vivo_from_uploads(vivo_txts)
            base_claro = build_base_claro_from_uploads(claro_pdfs)
            base_tim = build_base_tim_from_uploads(tim_pdfs)

            if debug:
                st.write("DEBUG extração (Cadastros):")
                st.write(f"VIVO: arquivos={len(vivo_txts) if vivo_txts else 0} | registros={len(base_vivo)}")
                st.write(f"CLARO: arquivos={len(claro_pdfs) if claro_pdfs else 0} | registros={len(base_claro)}")
                st.write(f"TIM: arquivos={len(tim_pdfs) if tim_pdfs else 0} | registros={len(base_tim)}")

                for f in (claro_pdfs or []):
                    t = pdf_bytes_to_text(f.getvalue())
                    if len(t.strip()) < 50:
                        st.warning(f"CLARO {f.name}: PDF com pouco/nenhum texto (possível escaneado).")
                for f in (tim_pdfs or []):
                    t = pdf_bytes_to_text(f.getvalue())
                    if len(t.strip()) < 50:
                        st.warning(f"TIM {f.name}: PDF com pouco/nenhum texto (possível escaneado).")

            base_all: Dict[str, Tuple[float, Dict[str, Optional[str]]]] = {}
            for b in (base_vivo, base_claro, base_tim):
                for tel, (ts, rec) in b.items():
                    old = base_all.get(tel)
                    if old is None or ts > old[0]:
                        base_all[tel] = (ts, rec)

            df_cad = dict_base_to_df(base_all)
            df_cad["Telefone"] = df_cad["Telefone"].apply(normalize_phone_br)
            df_cad = df_cad[df_cad["Telefone"] != ""].drop_duplicates("Telefone").copy()
            st.session_state["df_cad"] = df_cad

            if df_cad.empty:
                st.warning("Nenhum cadastro extraído (Vivo/Claro/Tim).")
                st.stop()

            # Cruzar Chamadas -> Identificados (apenas números presentes no histórico)
            if st.session_state["df_calls"] is not None:
                df_calls = st.session_state["df_calls"].copy()
                if "outro_numero" in df_calls.columns:
                    nums_hist = df_calls["outro_numero"].fillna("").astype(str).apply(normalize_phone_br)
                    nums_hist = sorted({n for n in nums_hist.tolist() if n})

                    df_hist = pd.DataFrame({"Telefone": nums_hist}).drop_duplicates()
                    df_ident_calls = df_hist.merge(
                        df_cad[["Telefone", "Nome", "CPF", "CIDADE", "OPERADORA", "Arquivo"]],
                        on="Telefone",
                        how="left",
                    )
                    df_ident_calls = df_ident_calls[
                        df_ident_calls["Nome"].notna()
                        | df_ident_calls["CPF"].notna()
                        | df_ident_calls["CIDADE"].notna()
                    ].copy()
                    st.session_state["df_ident_calls"] = df_ident_calls
                else:
                    st.session_state["df_ident_calls"] = pd.DataFrame(columns=["Telefone", "Nome", "CPF", "CIDADE", "OPERADORA", "Arquivo"])
            else:
                st.session_state["df_ident_calls"] = None

            # Cruzar WhatsApp (sem alvo)
            if st.session_state["df_msgs"] is not None:
                st.session_state["df_ident_wa"] = cruzar_whatsapp_sem_alvo(st.session_state["df_msgs"], df_cad)
            else:
                st.session_state["df_ident_wa"] = None

            # TXT não identificados
            txt_bytes, unknown_count = gerar_txt_nao_identificados(
                df_cad=df_cad,
                df_calls=st.session_state["df_calls"],
                df_msgs=st.session_state["df_msgs"],
            )
            st.session_state["unknown_txt_bytes"] = txt_bytes
            st.session_state["unknown_count"] = unknown_count

            n_calls = len(st.session_state["df_calls"]) if st.session_state["df_calls"] is not None else 0
            n_wa = len(st.session_state["df_msgs"]) if st.session_state["df_msgs"] is not None else 0
            n_cad = len(df_cad)
            n_id_calls = len(st.session_state["df_ident_calls"]) if st.session_state["df_ident_calls"] is not None else 0
            n_id_wa = len(st.session_state["df_ident_wa"]) if st.session_state["df_ident_wa"] is not None else 0

            st.info(
                f"Cadastros: {n_cad} | Chamadas: {n_calls} (Identificados: {n_id_calls}) | "
                f"WhatsApp: {n_wa} (Identificados: {n_id_wa}) | Não identificados (TXT): {unknown_count}"
            )

        # Resultados
        if st.session_state["df_cad"] is not None:
            st.subheader("4) Cadastros extraídos (base completa)")
            st.write(f"Linhas: {len(st.session_state['df_cad'])}")
            st.dataframe(st.session_state["df_cad"], width="stretch", height=420)

        if st.session_state["df_ident_calls"] is not None:
            st.subheader("5) Identificados - Chamadas")
            df_ic = st.session_state["df_ident_calls"].copy()
            st.write(f"Linhas: {len(df_ic)}")
            st.dataframe(df_ic, width="stretch", height=420)

        if st.session_state["df_ident_wa"] is not None:
            st.subheader("6) Identificados - WhatsApp (Remetente + Destinatário(s))")
            df_iw = st.session_state["df_ident_wa"].copy()
            st.write(f"Linhas: {len(df_iw)}")
            st.dataframe(df_iw, width="stretch", height=420)

        # Export TXT (não identificados)
        st.subheader("7) Exportar NÃO identificados (TXT)")
        if st.session_state.get("unknown_txt_bytes"):
            st.download_button(
                f"Baixar NÃO identificados (TXT) - {st.session_state.get('unknown_count', 0)} números",
                data=st.session_state["unknown_txt_bytes"],
                file_name="nao_identificados.txt",
                mime="text/plain",
            )
        else:
            st.caption("Após processar cadastros, o TXT de não identificados ficará disponível aqui.")

        # Export Excel
        st.subheader("8) Exportar Excel")
        if st.button("Gerar arquivo Excel"):
            if st.session_state["df_cad"] is None:
                st.error("Não há cadastros processados ainda.")
            else:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    if st.session_state["df_calls"] is not None:
                        st.session_state["df_calls"].to_excel(writer, index=False, sheet_name="Chamadas")
                    if st.session_state["df_ident_calls"] is not None:
                        st.session_state["df_ident_calls"].to_excel(writer, index=False, sheet_name="Identificados_Chamadas")

                    if st.session_state["df_msgs"] is not None:
                        st.session_state["df_msgs"].to_excel(writer, index=False, sheet_name="Mensagens")
                    if st.session_state["df_ident_wa"] is not None:
                        st.session_state["df_ident_wa"].to_excel(writer, index=False, sheet_name="Identificados_Mensagens")

                    st.session_state["df_cad"].to_excel(writer, index=False, sheet_name="Cadastros")

                output.seek(0)
                st.session_state["xlsx_bytes"] = output.getvalue()
                st.success("Excel gerado. Use o botão abaixo para baixar.")

        if st.session_state.get("xlsx_bytes"):
            st.download_button(
                "Baixar Excel (Chamadas + WhatsApp + Identificados + Cadastros)",
                data=st.session_state["xlsx_bytes"],
                file_name="resultado_i2.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


    # =========================================================
    # TAB 2 - ESTATÍSTICAS
    # =========================================================
    with tabs[1]:
        st.subheader("Resumo estatístico (Top números mais frequentes)")

        top_n = st.slider("Top N", min_value=5, max_value=100, value=20, step=5)

        df_calls = st.session_state.get("df_calls")
        df_msgs = st.session_state.get("df_msgs")
        df_cad = st.session_state.get("df_cad")

        if df_calls is None and df_msgs is None:
            st.warning("Carregue pelo menos um histórico (Chamadas ou WhatsApp) na aba Operação.")
        else:
            top_calls, top_wa = montar_top_frequencias(df_calls, df_msgs, df_cad, top_n=top_n)

            c1, c2 = st.columns(2)

            with c1:
                st.markdown("### Chamadas — Top números (com Nome se identificado)")
                if top_calls.empty:
                    st.caption("Sem dados de chamadas (ou coluna outro_numero não encontrada).")
                else:
                    st.dataframe(top_calls, width="stretch", height=520)
                    st.caption(
                        f"Identificados no Top: {int(top_calls['Identificado'].sum())} | "
                        f"Não identificados no Top: {int((~top_calls['Identificado']).sum())}"
                    )

            with c2:
                st.markdown("### WhatsApp — Top números (com Nome se identificado)")
                if top_wa.empty:
                    st.caption("Sem dados de WhatsApp.")
                else:
                    st.dataframe(top_wa, width="stretch", height=520)
                    st.caption(
                        f"Identificados no Top: {int(top_wa['Identificado'].sum())} | "
                        f"Não identificados no Top: {int((~top_wa['Identificado']).sum())}"
                    )

            st.markdown("### Visual (Pizza) — Identificados vs Não identificados")

            known = build_known_set(df_cad)

            # Chamadas: únicos em outro_numero
            calls_set = set()
            if df_calls is not None and "outro_numero" in df_calls.columns:
                s = df_calls["outro_numero"].fillna("").astype(str).apply(normalize_phone_br)
                calls_set = {n for n in s.tolist() if n}

            calls_ident = len([n for n in calls_set if n in known]) if known else 0
            calls_nao = len(calls_set) - calls_ident

            # WhatsApp: únicos em remetente_norm + destinatario_norm
            wa_set = set()
            if df_msgs is not None:
                if "remetente_norm" in df_msgs.columns:
                    s1 = df_msgs["remetente_norm"].fillna("").astype(str).apply(normalize_phone_br)
                    wa_set |= {n for n in s1.tolist() if n}
                if "destinatario_norm" in df_msgs.columns:
                    s2 = df_msgs["destinatario_norm"].fillna("").astype(str).apply(normalize_phone_br)
                    wa_set |= {n for n in s2.tolist() if n}

            wa_ident = len([n for n in wa_set if n in known]) if known else 0
            wa_nao = len(wa_set) - wa_ident

            g1, g2 = st.columns(2)

            with g1:
                st.markdown("#### Chamadas (únicos)")
                if len(calls_set) == 0:
                    st.caption("Sem dados suficientes para gráfico de chamadas.")
                else:
                    fig = plt.figure()
                    plt.pie(
                        [calls_ident, calls_nao],
                        labels=[f"Identificados ({calls_ident})", f"Não identificados ({calls_nao})"],
                        autopct="%1.1f%%",
                        startangle=90,
                    )
                    plt.title("Chamadas — Identificados vs Não")
                    st.pyplot(fig, clear_figure=True)

            with g2:
                st.markdown("#### WhatsApp (únicos)")
                if len(wa_set) == 0:
                    st.caption("Sem dados suficientes para gráfico de WhatsApp.")
                else:
                    fig = plt.figure()
                    plt.pie(
                        [wa_ident, wa_nao],
                        labels=[f"Identificados ({wa_ident})", f"Não identificados ({wa_nao})"],
                        autopct="%1.1f%%",
                        startangle=90,
                    )
                    plt.title("WhatsApp — Identificados vs Não")
                    st.pyplot(fig, clear_figure=True)

            st.caption("Gráfico considera NÚMEROS ÚNICOS (não volume de eventos).")


if __name__ == '__main__':
    main()
