"""
Espelho DFC - VENA
Gera novo arquivo xlsx com 4 abas baseado nas bases do mês anterior:
  1. Projecao de Entradas DFC  → Parcelas de Cartões (mês anterior)
  2. Projecao de Compra        → Rateio de Títulos (mês anterior)
  3. Parametros_ProjecaoCompra → Informativo de Estoque (mês anterior)
  4. Condicoes de Compra       → FC Consolidado (linha 33) + CMV (canal)
"""

import re
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Constantes ────────────────────────────────────────────────────────────────
MESES_DISPLAY = [
    "Janeiro","Fevereiro","Março","Abril","Maio","Junho",
    "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"
]
MES_3 = {
    "Janeiro":"Jan","Fevereiro":"Fev","Março":"Mar","Abril":"Abr",
    "Maio":"Mai","Junho":"Jun","Julho":"Jul","Agosto":"Ago",
    "Setembro":"Set","Outubro":"Out","Novembro":"Nov","Dezembro":"Dez"
}
MES_NUM = {m: i+1 for i, m in enumerate(MESES_DISPLAY)}
MES_ABR = {
    "Janeiro":"JAN","Fevereiro":"FEV","Março":"MAR","Abril":"ABR",
    "Maio":"MAI","Junho":"JUN","Julho":"JUL","Agosto":"AGO",
    "Setembro":"SET","Outubro":"OUT","Novembro":"NOV","Dezembro":"DEZ"
}

# ── Estilo helper ─────────────────────────────────────────────────────────────
DARK_BG  = "1F3864"
MID_BG   = "2E75B6"
LIGHT_BG = "D6E4F0"
WHITE    = "FFFFFF"

def _style(cell, bold=False, bg=None, fg="000000", center=False, num_fmt=None):
    cell.font = Font(bold=bold, color=fg, name="Calibri", size=10)
    if bg:
        cell.fill = PatternFill("solid", start_color=bg)
    if center:
        cell.alignment = Alignment(horizontal="center", vertical="center")
    thin = Side(style="thin", color="BFBFBF")
    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
    if num_fmt:
        cell.number_format = num_fmt

def _header(ws, title, n_cols, row=1):
    ws.merge_cells(f"A{row}:{get_column_letter(n_cols)}{row}")
    cell = ws[f"A{row}"]
    cell.value = title
    _style(cell, bold=True, bg=DARK_BG, fg=WHITE, center=True)
    ws.row_dimensions[row].height = 22

def mes_label(period):
    """Period → 'Mar'26'"""
    return f"{MES_3[MESES_DISPLAY[period.month-1]]}'{str(period.year)[2:]}"

def mes_anterior(mes_display):
    idx = MESES_DISPLAY.index(mes_display)
    return MESES_DISPLAY[idx-1] if idx > 0 else MESES_DISPLAY[11]

def ano_anterior(mes_display, ano):
    return ano-1 if mes_display == "Janeiro" else ano

# ── Helpers de leitura ────────────────────────────────────────────────────────
def _parse_br_csv(path):
    for enc in ["latin1", "utf-8", "cp1252"]:
        try: return pd.read_csv(path, encoding=enc, sep=None, engine="python")
        except: pass
    raise ValueError(f"Não foi possível ler {path}")

def _store_code(s):
    m = re.match(r"^(\d+)\s*[-–]", str(s).strip())
    return int(m.group(1)) if m else None


# ══════════════════════════════════════════════════════════════════════════════
# ABA 1 — Projeção de Entradas DFC
# ══════════════════════════════════════════════════════════════════════════════
def _aba_projecao_entradas(wb, cartoes_paths, mes_display):
    """
    Linhas: Saldo Anterior Crédito (Boleto Garantido) e Saldo Anterior Boleto (demais)
    Colunas: meses de vencimento encontrados no arquivo
    Excluir: onde Observações está preenchida
    """
    frames = [pd.read_excel(p, header=0) for p in cartoes_paths]
    df = pd.concat(frames, ignore_index=True)

    # Filtro Observações: excluir onde está preenchida (NaN real ou string vazia = válido)
    if "Observações" in df.columns:
        obs_raw = df["Observações"]
        # Aceitar: NaN real, string "NaN", string "nan", string vazia
        df = df[
            obs_raw.isna() |
            (obs_raw.astype(str).str.strip().str.lower().isin(["", "nan"]))
        ]

    df["_venc"] = pd.to_datetime(df["Dt. Vencimento"] if "Dt. Vencimento" in df.columns else df.iloc[:,10], errors="coerce")
    df["_mod"]  = (df["Modalidade"] if "Modalidade" in df.columns else df.iloc[:,7]).astype(str).str.strip()
    df["_vb"]   = pd.to_numeric(df["Valor Bruto"] if "Valor Bruto" in df.columns else df.iloc[:,11], errors="coerce").fillna(0)
    df["_mes"]  = df["_venc"].dt.to_period("M")

    boleto  = df[df["_mod"] == "Boleto Garantido"].groupby("_mes")["_vb"].sum()
    credito = df[df["_mod"] != "Boleto Garantido"].groupby("_mes")["_vb"].sum()
    months  = sorted(set(list(boleto.index) + list(credito.index)))

    ws = wb.create_sheet("Projecao de Entradas DFC")
    _header(ws, "PROJEÇÕES | ENTRADAS", 1 + len(months))
    ws.column_dimensions["A"].width = 30

    # Row 3: month headers
    ws.cell(row=3, column=1, value="VOLUMES DE ENTRADA")
    _style(ws.cell(row=3, column=1), bold=True, bg=DARK_BG, fg=WHITE, center=True)
    for i, m in enumerate(months):
        c = ws.cell(row=3, column=i+2, value=mes_label(m))
        _style(c, bold=True, bg=DARK_BG, fg=WHITE, center=True)
        ws.column_dimensions[get_column_letter(i+2)].width = 16

    rows_data = [
        ("Saldo Anterior Crédito", credito),   # não-Boleto Garantido
        ("Saldo Anterior Boleto",  boleto),    # Boleto Garantido
    ]
    for r_i, (label, series) in enumerate(rows_data):
        row = r_i + 4
        ws.cell(row=row, column=1, value=label)
        _style(ws.cell(row=row, column=1), bold=True, bg=MID_BG, fg=WHITE)
        for i, m in enumerate(months):
            c = ws.cell(row=row, column=i+2, value=round(float(series.get(m, 0)), 2))
            _style(c, bg=LIGHT_BG, center=True, num_fmt="#,##0.00")
        ws.row_dimensions[row].height = 18

    return {"linhas": len(months), "meses": [str(m) for m in months]}


# ══════════════════════════════════════════════════════════════════════════════
# ABA 2 — Projeção de Compra
# ══════════════════════════════════════════════════════════════════════════════
def _aba_projecao_compra(wb, rateio_path, fc_path):
    """
    Linhas: Compras Mercadorias, Taxa De Remuneração De Franquia, Esforço De Mkt
    Colunas: meses de vencimento encontrados no Rateio
    """
    # Auto-detect header row (different CPs may have different structures)
    df = None
    for _h in range(15):
        try:
            _df = pd.read_excel(rateio_path, header=_h, nrows=2)
            if "Plano de Contas" in _df.columns and "Valor Bruto" in _df.columns:
                df = pd.read_excel(rateio_path, header=_h)
                break
        except: pass
    if df is None:
        raise ValueError("Não foi possível detectar o cabeçalho do arquivo Contas a Pagar")
    if "Tipo de Documento" in df.columns:
        df = df[df["Tipo de Documento"].astype(str).str.strip() != "Previsão"]
    if "Status" in df.columns:
        df = df[df["Status"].astype(str).str.strip() != "Baixado"]
    df["Valor Bruto"] = pd.to_numeric(df["Valor Bruto"], errors="coerce").fillna(0)
    df["_venc"] = pd.to_datetime(df["Vencimento"], errors="coerce")
    df["_mes"]  = df["_venc"].dt.to_period("M")
    df["_plano"] = df["Plano de Contas"].astype(str).str.strip()

    # Calamo categories from Plano de Contas F360
    try:
        df_pc = pd.read_excel(fc_path, sheet_name="Plano de Contas F360", header=0)
        df_pc["_grupo"] = df_pc["Fornecedor/Grupo"].astype(str).str.strip()
        df_pc["_cat"]   = df_pc["Categoria"].astype(str).str.strip()
        calamo_cats = set(df_pc[df_pc["_grupo"].str.lower().str.contains("calamo", na=False)]["_cat"])
    except Exception:
        calamo_cats = set()

    merc = df[df["_plano"].isin(calamo_cats)].groupby("_mes")["Valor Bruto"].sum()
    # TRF: match qualquer variação do nome (com ou sem acento, com ou sem "(FC)")
    trf  = df[df["_plano"].str.lower().str.contains("taxa de remunera", na=False)].groupby("_mes")["Valor Bruto"].sum()
    mkt  = df[df["_plano"].str.lower().str.contains("marketing institucional", na=False)].groupby("_mes")["Valor Bruto"].sum()

    months = sorted(set(list(merc.index) + list(trf.index) + list(mkt.index)))

    ws = wb.create_sheet("Projecao de Compra")
    _header(ws, "PROJEÇÕES | COMPRAS | PROJEÇÃO", 1 + len(months))
    ws.column_dimensions["A"].width = 36

    # Row 3: headers
    ws.cell(row=3, column=1, value="Saldo à pagar Industria")
    _style(ws.cell(row=3, column=1), bold=True, bg=DARK_BG, fg=WHITE, center=True)
    for i, m in enumerate(months):
        c = ws.cell(row=3, column=i+2, value=mes_label(m))
        _style(c, bold=True, bg=DARK_BG, fg=WHITE, center=True)
        ws.column_dimensions[get_column_letter(i+2)].width = 16

    rows_data = [
        ("Compras Mercadorias",              merc),
        ("Taxa De Remuneração De Franquia",  trf),
        ("Esforço De Mkt",                   mkt),
    ]
    for r_i, (label, series) in enumerate(rows_data):
        row = r_i + 4
        ws.cell(row=row, column=1, value=label)
        _style(ws.cell(row=row, column=1), bold=True, bg=MID_BG, fg=WHITE)
        for i, m in enumerate(months):
            c = ws.cell(row=row, column=i+2, value=round(float(series.get(m, 0)), 2))
            _style(c, bg=LIGHT_BG, center=True, num_fmt="#,##0.00")
        ws.row_dimensions[row].height = 18

    return {"linhas": 3, "meses": [str(m) for m in months]}


# ══════════════════════════════════════════════════════════════════════════════
# ABA 3 — Parametros_ProjecaoCompra
# ══════════════════════════════════════════════════════════════════════════════
def _aba_parametros(wb, estoque_path, mes_display, ano=2026):
    """
    Lojas e Saldo Final do Informativo de Estoque do mês anterior
    """
    mes_ant = mes_anterior(mes_display)
    aba_est = MES_ABR[mes_ant]
    mes_label_ant = f"{MES_3[mes_ant].lower()}-{str(ano if mes_display != 'Janeiro' else ano-1)[2:]}"

    xl_est = pd.ExcelFile(estoque_path)
    aba_real = aba_est if aba_est in xl_est.sheet_names else next(
        (s for s in xl_est.sheet_names if s.upper() == aba_est), None)
    if aba_real is None:
        return {"erro": f"Aba '{aba_est}' não encontrada no estoque"}

    df_est = pd.read_excel(xl_est, sheet_name=aba_real, header=None)
    lojas = []
    for i, row in df_est.iterrows():
        val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        est = row.iloc[1] if pd.notna(row.iloc[1]) else None
        if re.match(r"^\d+\s*-\s*", val) and est is not None:
            lojas.append({"loja": val, "estoque": float(est)})

    ws = wb.create_sheet("Parametros_ProjecaoCompra")
    headers = ["LOJA", "MÊS", "SALDO INICIAL R$", "COMPRAS R$", "SALDO FINAL R$"]
    widths  = [45, 12, 18, 18, 18]

    for j, (h, w) in enumerate(zip(headers, widths)):
        c = ws.cell(row=1, column=j+1, value=h)
        _style(c, bold=True, bg=DARK_BG, fg=WHITE, center=True)
        ws.column_dimensions[get_column_letter(j+1)].width = w
    ws.row_dimensions[1].height = 20

    for i, loja in enumerate(lojas):
        row = i + 2
        c = ws.cell(row=row, column=1, value=loja["loja"])
        _style(c, bold=True, bg=MID_BG, fg=WHITE)
        c = ws.cell(row=row, column=2, value=mes_label_ant)
        _style(c, bg=LIGHT_BG, center=True)
        c = ws.cell(row=row, column=3, value=None)
        _style(c, bg=LIGHT_BG, center=True)
        c = ws.cell(row=row, column=4, value=None)
        _style(c, bg=LIGHT_BG, center=True)
        c = ws.cell(row=row, column=5, value=round(loja["estoque"], 2))
        _style(c, bg=LIGHT_BG, center=True, num_fmt="#,##0.00")
        ws.row_dimensions[row].height = 18

    return {"lojas": [l["loja"] for l in lojas]}


# ══════════════════════════════════════════════════════════════════════════════
# ABA 4 — Condições de Compra
# ══════════════════════════════════════════════════════════════════════════════
def _aba_condicoes(wb, fc_path, cmv_path, estoque_path, mes_display, ano=2026):
    """
    Lojas + Canal (do CMV) + Cobert. Estoque (linha 33 do FC, mês atual)
    """
    # Canal por código de loja (do CMV)
    VD_NOMES    = ["VD","VENDA DIRETA","Venda Direta","VENDAS DIRETAS"]
    LOJAS_NOMES = ["LOJAS","Lojas","LOJA"]
    xl_cmv = pd.ExcelFile(cmv_path)
    canal_por_codigo = {}
    for sheet_names, canal in [(VD_NOMES, "VD"), (LOJAS_NOMES, "Lojas")]:
        sheet = next((n for n in sheet_names if n in xl_cmv.sheet_names), None)
        if not sheet: continue
        df = pd.read_excel(xl_cmv, sheet_name=sheet, header=None, nrows=60)
        for i, row in df.iterrows():
            for v in row:
                if pd.notna(v):
                    m = re.match(r"^(\d{4,6})\s*[-–]", str(v).strip())
                    if m:
                        canal_por_codigo[int(m.group(1))] = canal

    # Prazo Médio de Estoque do FC (linha 33, mês atual)
    try:
        _xl_aba = pd.ExcelFile(fc_path)
        _aba_r  = next((s for s in _xl_aba.sheet_names if "real" in s.lower()), "Real 2022")
        df_fc  = pd.read_excel(fc_path, sheet_name=_aba_r, header=None)
        row7   = df_fc.iloc[6,:]
        mes_h  = f"{MES_3[mes_display]}'{str(ano)[2:]}"
        fc_col = next((i for i,v in enumerate(row7) if pd.notna(v) and str(v).strip().upper()==mes_h.upper()), None)
        if fc_col is None:
            # Fallback: use last available month with data
            valid_cols = [i for i,v in enumerate(row7) if pd.notna(v) and "'" in str(v)]
            fc_col = max(valid_cols) if valid_cols else None
        prazo_medio = 0
        if fc_col is not None:
            # Buscar linha "Prazo Médio de Estoque" nas colunas C e D (índices 2 e 3)
            _row_prazo = None
            for _ci in [2, 3]:
                _col = df_fc.iloc[:, _ci].apply(lambda x: str(x).strip().lower() if pd.notna(x) else "")
                _row_prazo = next((i for i,v in enumerate(_col) if "prazo" in v and "estoque" in v), None)
                if _row_prazo is not None:
                    break
            if _row_prazo is None:
                _row_prazo = 32  # fallback
            val = df_fc.iloc[_row_prazo, fc_col]
            prazo_medio = round(float(val)) if pd.notna(val) else 0
    except Exception as _e:
        prazo_medio = 0

    # Lojas do estoque mês anterior
    mes_ant = mes_anterior(mes_display)
    aba_est = MES_ABR[mes_ant]
    xl_est  = pd.ExcelFile(estoque_path)
    aba_real = aba_est if aba_est in xl_est.sheet_names else next(
        (s for s in xl_est.sheet_names if s.upper() == aba_est), None)

    lojas = []
    if aba_real:
        df_est = pd.read_excel(xl_est, sheet_name=aba_real, header=None)
        for i, row in df_est.iterrows():
            val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            if re.match(r"^\d+\s*-\s*", val):
                m = re.match(r"^(\d+)", val)
                code  = int(m.group(1)) if m else None
                canal = canal_por_codigo.get(code, "Lojas")
                lojas.append({"loja": val, "canal": canal})

    ws = wb.create_sheet("Condicoes de Compra")
    headers = ["Lojas", "Canal", "Cobert. Estoque"]
    widths  = [45, 12, 18]

    for j, (h, w) in enumerate(zip(headers, widths)):
        c = ws.cell(row=1, column=j+1, value=h)
        _style(c, bold=True, bg=DARK_BG, fg=WHITE, center=True)
        ws.column_dimensions[get_column_letter(j+1)].width = w
    ws.row_dimensions[1].height = 20

    for i, loja in enumerate(lojas):
        row = i + 2
        c = ws.cell(row=row, column=1, value=loja["loja"])
        _style(c, bold=True, bg=MID_BG, fg=WHITE)
        c = ws.cell(row=row, column=2, value=loja["canal"])
        _style(c, bg=LIGHT_BG, center=True)
        c = ws.cell(row=row, column=3, value=prazo_medio)
        _style(c, bg=LIGHT_BG, center=True, num_fmt="0")
        ws.row_dimensions[row].height = 18

    return {"lojas": [l["loja"] for l in lojas], "prazo_medio": prazo_medio}


# ══════════════════════════════════════════════════════════════════════════════
# FUNÇÃO PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════
def gerar_espelho_dfc(
    cartoes_paths,    # list[str] — Parcelas de Cartões do mês anterior
    rateio_path,      # str       — Rateio de Títulos do mês anterior
    fc_path,          # str       — FC Consolidado (xlsx/xlsm)
    cmv_path,         # str       — Planilha CMV mestre
    estoque_path,     # str       — Informativo de Estoque
    output_path,      # str       — Arquivo de saída
    mes_display,      # str       — Mês que está sendo fechado (ex: "Março")
    ano=2026,
):
    report = {}
    wb = Workbook()
    wb.remove(wb.active)  # remove default sheet

    report["aba1"] = _aba_projecao_entradas(wb, cartoes_paths, mes_display)
    report["aba2"] = _aba_projecao_compra(wb, rateio_path, fc_path)
    report["aba3"] = _aba_parametros(wb, estoque_path, mes_display, ano)
    report["aba4"] = _aba_condicoes(wb, fc_path, cmv_path, estoque_path, mes_display, ano)

    wb.save(output_path)
    return report
