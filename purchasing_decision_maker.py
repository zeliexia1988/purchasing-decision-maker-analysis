import streamlit as st
import pandas as pd
import math
import urllib.parse
import json
import io
import re
import time
from datetime import datetime

# 报价分析模块依赖
import google.generativeai as genai
from pdf2image import convert_from_bytes

# ===============================
# 1. 基础配置与数据加载
# ===============================
st.set_page_config(page_title="SADE 采购决策支持系统", layout="wide")

@st.cache_data
def load_all_data():
    try:
        # 1. 加载合同表 (Excel)
        df_contracts = pd.read_excel("contracts_b.xlsx")
        for col in ["DE", "PN", "Price", "MOQ 12ml"]:
            df_contracts[col] = pd.to_numeric(df_contracts[col], errors='coerce')

        # 2. 加载运费表
        df_transport = pd.read_excel("Transport PE.xlsx")

        # 3. 数据清洗
        df_transport.columns = df_transport.columns.str.strip()
        df_transport['Dpt'] = df_transport['Dpt'].astype(str).str.strip()
        df_transport['DEPARTEMENTS'] = df_transport['DEPARTEMENTS'].astype(str).str.strip()
        df_transport['Supplier'] = df_transport['Supplier'].astype(str).str.strip()

        # 4. 生成省份列表
        dept_info = df_transport[['Dpt', 'DEPARTEMENTS']].drop_duplicates().sort_values('Dpt')
        dept_list = [f"{row['Dpt']} - {row['DEPARTEMENTS']}" for _, row in dept_info.iterrows()]

        return df_contracts, df_transport, dept_list
    except Exception as e:
        st.error(f"加载文件失败，请检查文件是否存在且格式正确: {e}")
        return None, None, []

contracts, transport_db, dept_options_list = load_all_data()

# ===============================
# 2. 核心业务规则 (Rules)
# ===============================
def rule_distributor_purchase(quantity, package, DE):
    return (package == "couronne" or DE < 125 or (DE < 200 and quantity < 960) or (225 <= DE <= 355 and quantity < 360))

def rule_contract_purchase(quantity, package, DE):
    return ((package == "barre" and 125 <= DE <= 200 and 960 <= quantity < 2000)
            or (package == "barre" and 225 <= DE <= 355 and 360 <= quantity < 1000))

def rule_factory_purchase(quantity, package, DE):
    return ((package == "barre" and 225 <= DE <= 355 and 1000 <= quantity)
            or (package == "barre" and 125 <= DE <= 200 and 2000 <= quantity)
            or package.lower() == "touret" or (package == "barre" and 355 < DE))

def rule_distributor_purchase_dipipe(quantity, DE):
    return (DE < 80)

def rule_contract_purchase_dipipe(quantity, DE):
    conditions = [
        (DE >= 300 and quantity <= 264), (DE >= 250 and quantity <= 396),
        (DE >= 200 and quantity <= 440), (DE >= 150 and quantity <= 594),
        (DE >= 125 and quantity <= 770), (DE >= 100 and quantity <= 891),
        (DE >= 80 and quantity <= 968)
    ]
    return any(conditions)

def rule_factory_purchase_dipipe(quantity, DE):
    return not rule_contract_purchase_dipipe(quantity, DE) and DE >= 80

def generate_email_template(material, quantity, de, pn, package, dept):
    subject = f"Demande de prix - {material} - DE{de} PN{pn}"
    body = (f"Bonjour,\n\nDans le cadre d'un nouveau projet, nous souhaiterions obtenir "
            f"votre meilleure offre pour :\n"
            f"- Produit : {material}\n"
            f"- DE : {de} / PN : {pn}\n"
            f"- Quantité : {quantity} ml\n"
            f"- Conditionnement : {package}\n"
            f"- Département de livraison : {dept}\n\nCordialement,")
    return subject, body

# ===============================
# 3. 价格计算逻辑 (MOQ + Transport)
# ===============================
def calculate_all_totals(material, de, pn, quantity, package, dept_code, today):
    pkg_str = str(package).lower() if package else ""
    mask = (
        (contracts["Material"] == material) &
        (contracts["Valid_Until"] >= today) &
        (contracts["DE"] == float(de)) &
        (contracts["PN"] == float(pn)) &
        (contracts["Package"].astype(str).str.lower() == pkg_str)
    )
    valid_matches = contracts[mask].copy()

    valid_matches = valid_matches[valid_matches["MOQ 12ml"].notna() & (valid_matches["MOQ 12ml"] > 0)]
    if valid_matches.empty:
        return None

    valid_matches["Nb_Camions"] = valid_matches["MOQ 12ml"].apply(lambda x: math.ceil(quantity / x))

    def get_fee(supplier):
        fee_m = (transport_db["Supplier"].str.contains(supplier, case=False, na=False)) & (transport_db["Dpt"] == str(dept_code))
        res = transport_db[fee_m]["Transport"]
        return res.iloc[0] if not res.empty else 0

    valid_matches["Transport_Unit"] = valid_matches["Supplier"].apply(get_fee)

    valid_matches["Material_Total"] = valid_matches["Price"] * quantity
    valid_matches["Total_Transport"] = valid_matches["Nb_Camions"] * valid_matches["Transport_Unit"]
    valid_matches["Grand_Total"] = valid_matches["Material_Total"] + valid_matches["Total_Transport"]

    display_df = valid_matches[["Supplier", "Price", "Nb_Camions", "Transport_Unit", "Total_Transport", "Grand_Total"]].copy()
    display_df.columns = ["Fournisseur", "Unit (€/ml)", "Camions", "Frais/Cam", "Total Trans", "TOTAL HT"]

    for col in ["Unit (€/ml)", "Frais/Cam", "Total Trans", "TOTAL HT"]:
        display_df[col] = display_df[col].map("{:,.2f} €".format)
    return display_df.sort_values("TOTAL HT")

# ===============================
# 4. 报价提取核心函数
# ===============================
QUOTE_FIELDS = ["Fournisseur", "Material", "DE", "PN", "Package",
                "Quantite_ml", "Prix_unitaire", "Devise", "Delai", "Date_validite"]


def _get_model():
    """初始化 Gemini，未配置 key 时返回 None"""
    api_key = st.secrets.get("GEMINI_API_KEY")
    if not api_key:
        st.error("⚠️ 未配置 GEMINI_API_KEY，请在 Streamlit Secrets 中添加。")
        return None
    genai.configure(api_key=api_key)
    return genai.GenerativeModel("gemini-2.5-flash")


def _call_gemini(model, content, label="", max_retries=3):
    """带限流重试的 Gemini 调用，返回清理后的文本或 None"""
    for attempt in range(max_retries):
        try:
            response = model.generate_content(content)
            return response.text.strip().replace("```json", "").replace("```", "").strip()
        except Exception as e:
            msg = str(e)
            is_quota = ("ResourceExhausted" in str(type(e)) or "429" in msg
                        or "quota" in msg.lower() or "exhausted" in msg.lower())
            if is_quota and attempt < max_retries - 1:
                wait = 30 * (attempt + 1)
                st.warning(f"⏳ 触发速率限制，{wait}秒后重试 ({attempt + 1}/{max_retries})… {label}")
                time.sleep(wait)
            elif is_quota:
                st.error(f"⚠️ 多次重试后仍被限流。免费层限额较低，请减少同时上传的文件数，"
                         f"或稍后再试。{label}")
                return None
            else:
                st.error(f"调用模型失败 {label}: {e}")
                return None
    return None


def _parse_json_rows(raw, label=""):
    """把模型返回文本解析成 DataFrame"""
    if raw is None:
        return None
    try:
        data = json.loads(raw)
        if isinstance(data, dict):
            data = [data]
        return pd.DataFrame(data)
    except json.JSONDecodeError:
        st.error(f"模型返回的不是有效 JSON {label}，原始输出如下:")
        st.code(raw[:3000])
        return None


PROMPT_RULES = f"""Renvoie UNIQUEMENT un tableau JSON (pas de texte, pas de markdown),
où chaque objet a ces clés exactes:
{json.dumps(QUOTE_FIELDS, ensure_ascii=False)}

Règles:
- Fournisseur: nom du fournisseur émetteur du devis
- DE: diamètre extérieur ou DN (nombre uniquement)
- PN: pression nominale (nombre uniquement)
- Package: conditionnement (barre, couronne, touret...)
- Quantite_ml: quantité (nombre)
- Prix_unitaire: prix unitaire (nombre uniquement, sans symbole)
- Devise: ex "EUR", "USD"
- Delai: délai de livraison si mentionné, sinon ""
- Date_validite: date de validité de l'offre si mentionnée, sinon ""
- Si une valeur est absente, mets "" ou null.
- Une ligne par produit/référence."""


def extract_quote_from_pdf(pdf_bytes, label=""):
    """PDF -> 图片 -> Gemini 多模态 -> 结构化 JSON"""
    model = _get_model()
    if model is None:
        return None

    images = convert_from_bytes(pdf_bytes, dpi=150)
    prompt = ("Tu es un assistant achats. Analyse ce devis fournisseur "
              "(canalisation/tuyauterie) et extrais TOUTES les lignes de produits.\n\n"
              + PROMPT_RULES)
    raw = _call_gemini(model, [prompt] + images, label=label)
    return _parse_json_rows(raw, label=label)


# ---- Excel：pandas 切表 + Gemini 仅映射列名 + pandas 批量转换 ----
HEADER_KEYWORDS = ["dn", "diam", "precio", "price", "prix", "preço", "cant", "qty",
                   "quant", "codigo", "código", "code", "descrip", "denomin",
                   "unit", "p.u", "peso", "pn", "référence", "designation", "désignation"]


def _detect_header_row(raw_df):
    """找出最像表头的行号（匹配关键词最多的行）"""
    def score(row):
        cells = [str(x).strip().lower() for x in row if str(x).strip() and str(x) != "nan"]
        return sum(any(k in c for c in cells) for k in HEADER_KEYWORDS)
    scores = raw_df.apply(lambda r: score(r.tolist()), axis=1)
    if scores.max() < 2:          # 关键词太少，判定为非规整表
        return None
    return int(scores.idxmax())


def _is_number(x):
    try:
        float(str(x).replace(",", ".").strip())
        return True
    except (ValueError, TypeError):
        return False


def _extract_pn(txt):
    m = re.search(r'PN\s*?(\d{1,3})', str(txt), re.I)
    return m.group(1) if m else ""


def extract_quote_from_excel(file_bytes, filename):
    """
    规整表：pandas 定位表头+数据区，只把表头发给 Gemini 做列名映射，
            再用 pandas 批量转换全部行（零逐行 LLM、不限流、不丢行）。
    乱表：  退回整表文本发 Gemini。
    """
    try:
        raw_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, header=None, dtype=str)
    except Exception as e:
        st.error(f"读取 Excel 失败 ({filename}): {e}")
        return None

    header_idx = _detect_header_row(raw_df)

    # ---------- 情况 A：识别到规整表头 ----------
    if header_idx is not None:
        header = [str(x).strip() for x in raw_df.iloc[header_idx].fillna("").tolist()]
        data = raw_df.iloc[header_idx + 1:].copy()
        data.columns = range(len(header))

        # 只保留首列为数字编号的产品行（去掉单位说明、合计、条款）
        data = data[data[0].apply(_is_number)].reset_index(drop=True)
        if data.empty:
            # 首列不是编号？放宽：保留至少一半单元格非空的行
            data = raw_df.iloc[header_idx + 1:].copy()
            data.columns = range(len(header))
            keep = data.apply(lambda r: r.notna().sum() >= len(header) / 2, axis=1)
            data = data[keep].reset_index(drop=True)

        # 让 Gemini 只映射列名 -> 列索引（小请求，省额度不丢行）
        model = _get_model()
        if model is None:
            return None
        header_desc = "\n".join(f"  [{i}] {h}" for i, h in enumerate(header) if h)
        map_prompt = f"""Tu mappes les colonnes d'un devis vers des champs standard.
Colonnes disponibles (index entre crochets):
{header_desc}

Renvoie UNIQUEMENT un objet JSON associant chaque champ standard à l'INDEX de colonne
correspondant (entier), ou null si absent. Champs:
- "DE": diamètre extérieur / DN
- "Material": désignation / description du produit
- "Package": conditionnement (barre/couronne/touret) si présent
- "Quantite_ml": quantité
- "Prix_unitaire": prix unitaire

Exemple de sortie: {{"DE": 3, "Material": 2, "Package": null, "Quantite_ml": 6, "Prix_unitaire": 8}}"""

        raw_map = _call_gemini(model, map_prompt, label=f"[{filename} en-têtes]")
        if raw_map is None:
            return None
        try:
            col_map = json.loads(raw_map)
        except json.JSONDecodeError:
            st.warning(f"列名映射解析失败，退回整表模式 ({filename})。")
            col_map = None

        if col_map:
            out = pd.DataFrame()
            for field in ["DE", "Material", "Package", "Quantite_ml", "Prix_unitaire"]:
                idx = col_map.get(field)
                out[field] = data[idx] if (idx is not None and idx in data.columns) else ""

            # PN：优先描述里正则提取
            out["PN"] = out["Material"].apply(_extract_pn)
            # 供应商/货币：整表第一次出现的线索由 Gemini 之外简单补；留空交由用户校对
            out["Fournisseur"] = ""
            out["Devise"] = ""
            out["Delai"] = ""
            out["Date_validite"] = ""

            # 清理：DE/数量/单价转为去千分位的字符串
            for c in ["DE", "Quantite_ml", "Prix_unitaire"]:
                out[c] = out[c].apply(lambda v: str(v).strip() if pd.notna(v) else "")
            # 去掉 DE 非数字的残留行（单位说明等）
            out = out[out["DE"].apply(lambda x: _is_number(x) or x == "")].reset_index(drop=True)
            out = out[QUOTE_FIELDS]
            return out

    # ---------- 情况 B：乱表，整表文本发 Gemini ----------
    model = _get_model()
    if model is None:
        return None
    excel_text = raw_df.to_csv(index=False, header=False)
    prompt = ("Tu es un assistant achats. Voici un devis fournisseur "
              "(canalisation/tuyauterie) exporté depuis Excel. Les colonnes ont des noms "
              "variables.\n\nExtrais TOUTES les lignes de produits.\n\n"
              + PROMPT_RULES + f"\n\nContenu:\n{excel_text}")
    raw = _call_gemini(model, prompt, label=f"[{filename}]")
    return _parse_json_rows(raw, label=f"[{filename}]")

# ===============================
# 5. Google Sheets 存档 (未配置时静默跳过)
# ===============================
def archive_to_sheets(df):
    try:
        import gspread
        from google.oauth2.service_account import Credentials

        if "gcp_service_account" not in st.secrets:
            return False, "Google Sheets 未配置(密钥未填)，已跳过云端存档。"

        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]), scopes=scopes
        )
        gc = gspread.authorize(creds)
        sh = gc.open_by_key(st.secrets["SHEET_KEY"])
        ws = sh.sheet1

        df_archive = df.copy()
        df_archive.insert(0, "Date_import", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

        if not ws.get_all_values():
            ws.append_row(df_archive.columns.tolist())
        ws.append_rows(df_archive.values.tolist())
        return True, "✅ 已存档到 Google Sheets。"
    except Exception as e:
        return False, f"云端存档失败: {e}"

# ===============================
# 5b. 根据报价参数匹配合同价格（两段式：优先本供应商，退回市场最低）
# ===============================
def lookup_contract_price(row, today=None):
    """
    返回 (合同单价, 来源说明)
    1. 先按 Material/DE/PN/Package 锁定规格
    2. 若报价供应商在该规格下有合同 -> 用其合同价（来源=Contrat 供应商）
    3. 否则 -> 取该规格市场最低合同价（来源=Marché min）
    """
    if contracts is None:
        return None, ""
    if today is None:
        today = datetime.today()

    try:
        de_val = float(row.get("DE")) if pd.notna(row.get("DE")) and str(row.get("DE")).strip() != "" else None
        pn_val = float(row.get("PN")) if pd.notna(row.get("PN")) and str(row.get("PN")).strip() != "" else None
    except (ValueError, TypeError):
        return None, ""

    material = str(row.get("Material", "")).strip().lower()
    package = str(row.get("Package", "")).strip().lower()
    supplier = str(row.get("Fournisseur", "")).strip()

    # 第一步：按规格筛选（不含 supplier）
    mask = (contracts["Valid_Until"] >= today)
    if material:
        mask &= (contracts["Material"].astype(str).str.strip().str.lower() == material)
    if de_val is not None:
        mask &= (contracts["DE"] == de_val)
    if pn_val is not None:
        mask &= (contracts["PN"] == pn_val)
    if package:
        mask &= (contracts["Package"].astype(str).str.strip().str.lower() == package)

    spec_matches = contracts[mask]
    if spec_matches.empty:
        return None, ""

    # 第二步：优先本供应商的合同价
    if supplier:
        own = spec_matches[
            spec_matches["Supplier"].astype(str).str.contains(supplier, case=False, na=False)
        ]
        if not own.empty:
            return own["Price"].min(), f"Contrat {supplier}"

    # 第三步：退回市场最低合同价
    best = spec_matches.loc[spec_matches["Price"].idxmin()]
    return best["Price"], f"Marché min ({best['Supplier']})"


def add_contract_prices(df):
    """给报价 DataFrame 增加合同价列、来源列与对比列"""
    df = df.copy()

    results = df.apply(lambda r: lookup_contract_price(r), axis=1)
    df["Prix_contrat"] = [r[0] for r in results]
    df["Source_contrat"] = [r[1] for r in results]

    def calc_ecart(r):
        try:
            pu = float(r.get("Prix_unitaire"))
            pc = r.get("Prix_contrat")
            if pc is None or pd.isna(pc):
                return None
            return round(pu - pc, 2)
        except (ValueError, TypeError):
            return None

    df["Ecart_vs_contrat"] = df.apply(calc_ecart, axis=1)
    return df


# ===============================
# 6. Excel 导出
# ===============================
def to_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Devis")
    return output.getvalue()

# ===============================
# 7. Streamlit UI (两个标签页)
# ===============================
st.title("🛡️ SADE Purchasing Decision Support")

tab1, tab2 = st.tabs(["🛡️ Décision d'achat", "📄 Analyse de devis (PDF)"])

# ---------- 标签页 1：决策系统 ----------
with tab1:
    if contracts is not None:
        with st.form("purchase_form"):
            col1, col2 = st.columns(2)
            with col1:
                material_choice = st.selectbox("Matériau:", options=[""] + sorted(contracts["Material"].dropna().unique().tolist()))
                package_choice = st.selectbox("Conditionnement:", options=["", "barre", "couronne", "touret"])
                qty_input = st.number_input("Quantité (ml):", min_value=0, step=1)
            with col2:
                de_choice = st.selectbox("DE (Diamètre):", options=[""] + sorted([int(x) for x in contracts["DE"].dropna().unique()]))
                pn_choice = st.selectbox("PN (Pression):", options=[""] + sorted([float(x) for x in contracts["PN"].dropna().unique()]))
                dept_full = st.selectbox("Département de livraison:", options=[""] + dept_options_list)

            submit_btn = st.form_submit_button("Run Decision", type="primary")

        if submit_btn and material_choice and package_choice and de_choice and dept_full:
            dept_code = dept_full.split(" - ")[0]
            today = datetime.today()

            is_fonte = "fonte" in material_choice.lower()
            price_table = None
            decision_msg = ""

            if is_fonte:
                if rule_contract_purchase_dipipe(qty_input, de_choice):
                    price_table = calculate_all_totals(material_choice, de_choice, pn_choice, qty_input, package_choice, dept_code, today)
                    decision_msg = "✅ Decision: Application tarif contractuel Electrosteel"
                elif rule_factory_purchase_dipipe(qty_input, de_choice):
                    decision_msg = "✅ Decision: Consultation Electrosteel (Usine)"
                else:
                    decision_msg = "🛒 Decision: Consultation Négoce"
            else:
                if rule_contract_purchase(qty_input, package_choice, de_choice):
                    price_table = calculate_all_totals(material_choice, de_choice, pn_choice, qty_input, package_choice, dept_code, today)
                    decision_msg = "✅ Decision: Application tarif contractuelle"
                elif rule_factory_purchase(qty_input, package_choice, de_choice):
                    price_table = calculate_all_totals(material_choice, de_choice, pn_choice, qty_input, package_choice, dept_code, today)
                    decision_msg = "✅ Decision: Consultation Fabricant (Elydan, Centraltubi) pour avoir meilleur prix que les conditions contractuelles"
                else:
                    decision_msg = "🛒 Decision: Consultation Négoce"

            st.divider()
            st.subheader(decision_msg)

            if price_table is not None:
                st.write("### 💰 Comparatif des prix (Transport inclus)")
                st.table(price_table)
            else:
                if "Application" in decision_msg:
                    st.warning("⚠️ Contrat trouvé mais MOQ 12ml non renseignée dans le fichier Excel.")

                if "Consultation" in decision_msg:
                    st.info("📧 **Brouillon d'Email de consultation**")
                    subject, body = generate_email_template(material_choice, qty_input, de_choice, pn_choice, package_choice, dept_full)

                    st.text_area("Copier :", value=body, height=160)

                    mailto_link = (
                        f"mailto:?subject={urllib.parse.quote(subject)}"
                        f"&body={urllib.parse.quote(body)}"
                    )
                    st.markdown(
                        f"""
                        <a href="{mailto_link}" target="_blank">
                            <button style="
                                background-color: #0072C6;
                                color: white;
                                padding: 8px 16px;
                                border: none;
                                border-radius: 4px;
                                cursor: pointer;
                                font-size: 14px;
                            ">
                            📨 Ouvrir dans Outlook
                            </button>
                        </a>
                        """,
                        unsafe_allow_html=True,
                    )

# ---------- 标签页 2：报价分析 ----------
with tab2:
    st.header("📄 Analyse automatique des devis")
    st.caption("Uploadez un ou plusieurs devis (PDF ou Excel). Les données seront extraites et consolidées.")
    st.info("💡 Conseil : déposez **10 fichiers maximum** à la fois (limite du quota gratuit Gemini). "
            "Pour traiter davantage de devis, faites plusieurs lots espacés d'une à deux minutes.")

    uploaded_files = st.file_uploader(
        "Déposez vos devis (PDF ou Excel)",
        type=["pdf", "xlsx", "xls"],
        accept_multiple_files=True
    )

    if uploaded_files and len(uploaded_files) > 10:
        st.warning(f"⚠️ Vous avez sélectionné {len(uploaded_files)} fichiers. "
                   f"Au-delà de 10, le traitement peut être lent ou atteindre la limite du quota gratuit. "
                   f"Pensez à traiter par lots plus petits.")

    if uploaded_files and st.button("🔍 Extraire les données", type="primary"):
        all_dfs = []
        progress = st.progress(0, text="Extraction en cours...")
        for i, f in enumerate(uploaded_files):
            with st.spinner(f"Analyse de {f.name}..."):
                file_bytes = f.read()
                if f.name.lower().endswith(".pdf"):
                    df = extract_quote_from_pdf(file_bytes, label=f"[{f.name}]")
                else:  # .xlsx / .xls
                    df = extract_quote_from_excel(file_bytes, f.name)

                if df is not None and not df.empty:
                    df.insert(0, "Fichier_source", f.name)
                    all_dfs.append(df)

            # 免费层限流保护：文件之间留间隔（最后一个不等）
            if i < len(uploaded_files) - 1:
                time.sleep(7)
            progress.progress((i + 1) / len(uploaded_files))

        if all_dfs:
            consolidated = pd.concat(all_dfs, ignore_index=True)
            consolidated = add_contract_prices(consolidated)   # 新增合同价/来源/价差列
            st.session_state["quote_df"] = consolidated
            progress.empty()
            st.success(f"✅ {len(consolidated)} ligne(s) extraite(s) depuis {len(uploaded_files)} fichier(s).")

    if "quote_df" in st.session_state:
        st.write("### 📝 Vérifiez et corrigez si nécessaire")
        edited = st.data_editor(
            st.session_state["quote_df"],
            use_container_width=True,
            num_rows="dynamic",
        )

        col_a, col_b = st.columns(2)
        with col_a:
            st.download_button(
                "⬇️ Télécharger Excel",
                data=to_excel_bytes(edited),
                file_name=f"devis_consolide_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        with col_b:
            if st.button("☁️ Archiver dans Google Sheets"):
                ok, msg = archive_to_sheets(edited)
                if ok:
                    st.success(msg)
                else:
                    st.warning(msg)

