import io
import re
import unicodedata
from typing import List, Dict, Optional, Tuple

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Quantificação BK - Excel v2", page_icon="📊", layout="wide")

st.title("📊 Quantificação de Materiais (AutoCAD → copiar/colar no Excel)")
st.caption("v2: captura automaticamente o **nome do desenho** (ex: 'ESCADA PARA DISJUNTOR') a partir das células próximas da tabela.")
st.markdown("---")

CANON_HEADERS = ["ITEM", "CÓD. BK", "CÓD. CLIENTE", "DESCRIÇÃO", "QUANT.", "UN."]

def strip_accents(s: str) -> str:
    s = unicodedata.normalize("NFKD", str(s))
    return "".join(ch for ch in s if not unicodedata.combining(ch))

def norm_cell(x) -> str:
    if x is None:
        return ""
    s = str(x).replace("\u00A0", " ").strip()
    s = re.sub(r"\s+", " ", s)
    return s

def norm_key(s: str) -> str:
    s = strip_accents(norm_cell(s)).upper()
    s = s.replace(".", "")
    s = re.sub(r"[^A-Z0-9]+", "", s)
    return s

HDR_KEYS = {
    "ITEM": {"ITEM"},
    "CÓD. BK": {"CODBK", "CODIGOBK", "CDBK"},
    "CÓD. CLIENTE": {"CODCLIENTE", "CODIGOCLIENTE", "CDCLIENTE"},
    "DESCRIÇÃO": {"DESCRICAO", "DESCR", "DESCRI", "MATERIAL", "NOME"},
    "QUANT.": {"QUANT", "QUANTIDADE", "QTD", "QTDE"},
    "UN.": {"UN", "UND", "UNID", "UNIDADE"},
}

def looks_header_row(row: List[str]) -> Optional[Dict[str,int]]:
    keys = [norm_key(c) for c in row]
    found: Dict[str,int] = {}
    for i,k in enumerate(keys):
        if not k:
            continue
        for h, variants in HDR_KEYS.items():
            if h in found:
                continue
            if k in variants:
                found[h] = i
            else:
                # contains fallback
                if h == "DESCRIÇÃO" and "DESCR" in k:
                    found[h] = i
                if h == "QUANT." and ("QUANT" in k or "QTD" in k):
                    found[h] = i
                if h == "CÓD. BK" and ("BK" in k and "COD" in k):
                    found[h] = i
                if h == "CÓD. CLIENTE" and ("CLIENTE" in k and "COD" in k):
                    found[h] = i
    if all(h in found for h in CANON_HEADERS):
        return found
    return None

def to_float_br(x) -> float:
    s = norm_cell(x)
    if not s:
        return 0.0
    s = s.replace(" ", "")
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0

def is_blank_row(row: List[str]) -> bool:
    return all(not norm_cell(c) for c in row)

def is_candidate_title(text: str) -> bool:
    t = norm_cell(text)
    if len(t) < 4:
        return False
    if not re.search(r"[A-Za-zÀ-ÿ]", t):
        return False
    u = strip_accents(t).upper()
    bad = ["ITEM", "COD", "CÓD", "CLIENTE", "DESCR", "DESCRI", "QUANT", "UN"]
    if any(b in u for b in bad):
        return False
    if re.fullmatch(r"[0-9\W]+", t):
        return False
    return True

def find_title_near(matrix: List[List[str]], header_row_idx: int) -> str:
    r0 = max(0, header_row_idx - 3)
    r1 = min(len(matrix) - 1, header_row_idx + 3)

    candidates: List[Tuple[int, str]] = []
    for r in range(r0, r1 + 1):
        for val in matrix[r]:
            txt = norm_cell(val)
            if is_candidate_title(txt):
                candidates.append((len(txt), txt))

    if not candidates:
        return ""
    candidates.sort(reverse=True, key=lambda x: x[0])
    return candidates[0][1]

def parse_blocks(df_raw: pd.DataFrame, fonte_nome: str, sheet_name: str, fallback_title: str = "") -> pd.DataFrame:
    matrix = [[norm_cell(x) for x in row] for row in df_raw.values.tolist()
]
    max_cols = df_raw.shape[1]

    rows_out = []
    i = 0
    bloco_n = 0

    while i < len(matrix):
        mapping = looks_header_row(matrix[i])
        if mapping:
            bloco_n += 1
            title = find_title_near(matrix, i) or fallback_title

            i += 1
            blank_streak = 0

            while i < len(matrix):
                r = matrix[i]
                if looks_header_row(r):
                    break

                if is_blank_row(r):
                    blank_streak += 1
                    if blank_streak >= 2:
                        break
                    i += 1
                    continue
                blank_streak = 0

                desc = r[mapping["DESCRIÇÃO"]] if mapping["DESCRIÇÃO"] < max_cols else ""
                if norm_cell(desc):
                    rows_out.append({
                        "ARQUIVO_EXCEL": fonte_nome,
                        "ABA": sheet_name,
                        "DESENHO": title,
                        "BLOCO_TABELA": bloco_n,
                        "ITEM": r[mapping["ITEM"]] if mapping["ITEM"] < max_cols else "",
                        "CÓD. BK": r[mapping["CÓD. BK"]] if mapping["CÓD. BK"] < max_cols else "",
                        "CÓD. CLIENTE": r[mapping["CÓD. CLIENTE"]] if mapping["CÓD. CLIENTE"] < max_cols else "",
                        "DESCRIÇÃO": desc,
                        "QUANT.": r[mapping["QUANT."]] if mapping["QUANT."] < max_cols else "",
                        "UN.": r[mapping["UN."]] if mapping["UN."] < max_cols else "",
                    })
                i += 1
            continue

        i += 1

    return pd.DataFrame(rows_out)

def consolidate(df_detail: pd.DataFrame) -> pd.DataFrame:
    if df_detail.empty:
        return df_detail

    agg = {}
    order = []

    for _, row in df_detail.iterrows():
        desc = str(row["DESCRIÇÃO"]).strip()
        if not desc:
            continue

        if desc not in agg:
            order.append(desc)
            agg[desc] = {
                "CÓD. BK": row.get("CÓD. BK", ""),
                "CÓD. CLIENTE": row.get("CÓD. CLIENTE", ""),
                "UN.": row.get("UN.", ""),
                "QSUM": 0.0,
                "FONTES": set(),
                "DESENHOS": set(),
            }

        agg[desc]["QSUM"] += to_float_br(row.get("QUANT.", "0"))
        fonte = f'{row.get("ARQUIVO_EXCEL","")} | {row.get("ABA","")} | T{row.get("BLOCO_TABELA","")}'
        agg[desc]["FONTES"].add(fonte)
        if str(row.get("DESENHO","")).strip():
            agg[desc]["DESENHOS"].add(str(row.get("DESENHO","")).strip())

    out = []
    for desc in order:
        d = agg[desc]
        out.append({
            "CÓD. BK": d["CÓD. BK"],
            "CÓD. CLIENTE": d["CÓD. CLIENTE"],
            "DESCRIÇÃO": desc,
            "QUANT.": d["QSUM"],
            "UN.": d["UN."],
            "DESENHOS": "; ".join(sorted(d["DESENHOS"])),
            "ARQUIVOS_ORIGEM": "; ".join(sorted(d["FONTES"])),
        })
    return pd.DataFrame(out)

# UI
st.sidebar.header("⚙️ Configurações")
max_files = st.sidebar.number_input("Máx. arquivos por lote", min_value=1, max_value=40, value=40, step=1)
modo_debug = st.sidebar.checkbox("🔍 Depuração (mostrar amostras)", value=False)
fallback_title = st.sidebar.text_input("Título padrão (se não achar perto da tabela)", value="")

st.sidebar.info("""
**Como preparar o Excel:**
- Copie a tabela do AutoCAD e cole no Excel
- Cole várias tabelas **uma abaixo da outra**
- Pode repetir o cabeçalho
- Se o nome do desenho estiver ao lado (ex: na coluna K), o app captura automaticamente
""")

files = st.file_uploader("📁 Envie os arquivos Excel (.xlsx) (até 40)", type=["xlsx"], accept_multiple_files=True)

if files:
    files = files[: int(max_files)]
    st.success(f"✅ {len(files)} arquivo(s) selecionado(s)")

    if st.button("🚀 Processar e Gerar Excel", type="primary"):
        all_detail = []

        for f in files:
            try:
                xls = pd.ExcelFile(f)
                for sheet in xls.sheet_names:
                    df_raw = pd.read_excel(xls, sheet_name=sheet, header=None, dtype=str)
                    df_detail = parse_blocks(df_raw, f.name, sheet, fallback_title=fallback_title.strip())
                    if modo_debug:
                        st.write(f"📌 {f.name} / {sheet}: linhas extraídas = {len(df_detail)}")
                        if not df_detail.empty:
                            st.dataframe(df_detail.head(12), use_container_width=True)
                            st.write("Títulos encontrados (amostra):", sorted(set(df_detail["DESENHO"].astype(str).head(50))))
                    if not df_detail.empty:
                        all_detail.append(df_detail)
            except Exception as e:
                st.error(f"❌ {f.name}: {e}")

        if not all_detail:
            st.error("❌ Não encontrei nenhuma tabela (cabeçalho) nos arquivos enviados.")
            st.info("Dica: confira se o cabeçalho tem: ITEM / CÓD. BK / CÓD. CLIENTE / DESCRIÇÃO / QUANT. / UN.")
        else:
            df_detail = pd.concat(all_detail, ignore_index=True).fillna("")
            st.markdown("### 🔎 Detalhado (linhas extraídas)")
            st.dataframe(df_detail, use_container_width=True, height=320)

            st.info("🔄 Consolidando por **DESCRIÇÃO** (nome idêntico) e somando **QUANT.**")
            df_cons = consolidate(df_detail)

            st.markdown("### 📊 Consolidado")
            st.dataframe(df_cons, use_container_width=True, height=320)

            out = io.BytesIO()
            with pd.ExcelWriter(out, engine="openpyxl") as writer:
                df_detail.to_excel(writer, index=False, sheet_name="Detalhado")
                df_cons.to_excel(writer, index=False, sheet_name="Consolidado")
            out.seek(0)

            st.download_button(
                "⬇️ Baixar Excel Final",
                data=out,
                file_name="quantificacao_materiais_BK.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
else:
    st.info("👆 Envie o(s) Excel(s) com as tabelas coladas para começar.")

st.markdown("---")
st.markdown("<div style='text-align:center;color:#666;'><small>BK Engenharia | Streamlit + Excel (v2)</small></div>", unsafe_allow_html=True)
