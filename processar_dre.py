"""
processar_dre.py
Módulo responsável por atualizar os arquivos ValoresDaDRE com lançamentos manuais.
Cada função recebe o caminho do arquivo ValoresDaDRE e os dados necessários.
"""
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import re
import io


# ── Helpers ──────────────────────────────────────────────────────────────────

def _br_to_float(val):
    """Converte valor brasileiro '1.234,56' para float."""
    if pd.isna(val):
        return 0.0
    s = str(val).strip().replace("R$", "").replace(" ", "")
    if not s or s == "-":
        return 0.0
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _extrair_codigo(texto):
    """Extrai o código numérico do início de uma string como '5546 - SMI - Loja'."""
    if pd.isna(texto):
        return None
    m = re.match(r"^\s*(\d+)", str(texto))
    return m.group(1).strip() if m else None


def _mes_para_data(mes_display, ano=2026):
    """
    Converte 'Março' → datetime(2026, 3, 1)
    Retorna objeto datetime.date com o primeiro dia do mês.
    """
    MESES = {
        "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4,
        "Maio": 5, "Junho": 6, "Julho": 7, "Agosto": 8,
        "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
    }
    from datetime import date
    num = MESES.get(mes_display, 1)
    return date(ano, num, 1)


def _carregar_fat_lojas(fat_path):
    """
    Carrega o GerencialVendas CSV e retorna dict {cod_loja: valor_receita_liquida_trocas}.
    Coluna F (índice 5) = Receita Líquida (-) Trocas.
    """
    df = pd.read_csv(fat_path, sep=None, engine="python", encoding="latin1")
    col_loja = df.columns[0]
    df["_cod"] = df[col_loja].astype(str).str.extract(r"^(\d+)")[0].str.strip()
    df["_val"] = df.iloc[:, 5].apply(_br_to_float)
    return dict(zip(df["_cod"], df["_val"]))


def _carregar_fat_vd(vd_path):
    """
    Carrega o ConsultaPedidos XLSX e retorna dict {cod_loja: soma_ValorPedido}.
    Código extraído de EstruturaPai.
    """
    df = pd.read_excel(vd_path)
    df["_cod"] = df["EstruturaPai"].astype(str).str.extract(r"^(\d+)")[0].str.strip()
    soma = df.groupby("_cod")["ValorPedido"].sum()
    return soma.to_dict()


def _identificar_canal_cmv(cmv_path):
    """
    Lê o CMV mestre e retorna dois sets: {codigos_vd}, {codigos_loja}.
    Busca abas com nome contendo 'VD' ou 'VENDA' e 'LOJA'.
    """
    # Tentar abrir o CMV (pode ser xlsm)
    try:
        wb = load_workbook(cmv_path, read_only=True, data_only=True, keep_vba=False)
    except Exception:
        try:
            wb = load_workbook(cmv_path, read_only=True, data_only=True, keep_vba=True)
        except Exception as e2:
            raise ValueError(f"Não foi possível abrir o CMV mestre: {e2}")
    nomes = wb.sheetnames

    # Encontrar aba VD
    aba_vd = None
    for n in nomes:
        if re.search(r"venda\s*direta|^vd$", n, re.IGNORECASE):
            aba_vd = n
            break

    # Encontrar aba LOJAS
    aba_lojas = None
    for n in nomes:
        if re.search(r"loja", n, re.IGNORECASE) and n != aba_vd:
            aba_lojas = n
            break

    codigos_vd    = set()
    codigos_loja  = set()

    def _extrair_codigos_aba(ws):
        cods = set()
        for row in ws.iter_rows(min_row=1, values_only=True):
            for cell in row:
                m = re.match(r"^\s*(\d{4,6})\b", str(cell or ""))
                if m:
                    cods.add(m.group(1))
            break  # só primeira linha com dados relevantes? Na verdade vamos varrer tudo
        # Varrer coluna A
        for row in ws.iter_rows(min_col=1, max_col=2, values_only=True):
            for cell in row:
                m = re.match(r"^\s*(\d{4,6})\b", str(cell or ""))
                if m:
                    cods.add(m.group(1))
        return cods

    if aba_vd:
        ws = wb[aba_vd]
        codigos_vd = _extrair_codigos_aba(ws)
    if aba_lojas:
        ws = wb[aba_lojas]
        codigos_loja = _extrair_codigos_aba(ws)

    wb.close()
    return codigos_vd, codigos_loja


def _abrir_dre(dre_path):
    """
    Carrega o ValoresDaDRE independente do formato (.xlsx, .xlsm, .xls, .xlsb).
    Se não conseguir abrir direto, converte via pandas para um xlsx temporário.
    Retorna (wb openpyxl, ws).
    """
    import os, tempfile

    ext = os.path.splitext(str(dre_path))[1].lower()

    # 1. Tentar openpyxl direto (xlsx / xlsm)
    for kv in (False, True):
        try:
            wb = load_workbook(dre_path, keep_vba=kv, data_only=True)
            return wb, wb.active
        except Exception:
            continue

    # 2. Fallback: ler com pandas e recriar workbook em memória
    engine = None
    if ext == ".xlsb":
        engine = "pyxlsb"
    elif ext == ".xls":
        engine = "xlrd"

    df_raw = pd.read_excel(dre_path, sheet_name=0, header=None, engine=engine)

    # Recriar como openpyxl Workbook
    from openpyxl import Workbook as _WB
    wb2 = _WB()
    ws2 = wb2.active
    ws2.title = "ValoresDaDRE"
    for r_idx, row in enumerate(df_raw.itertuples(index=False), start=1):
        for c_idx, val in enumerate(row, start=1):
            # Converter NaN para None
            cell_val = None if (isinstance(val, float) and pd.isna(val)) else val
            ws2.cell(row=r_idx, column=c_idx, value=cell_val)

    # Salvar em temp e recarregar para ter comportamento consistente
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp.close()
    wb2.save(tmp.name)
    wb3 = load_workbook(tmp.name, data_only=True)
    os.unlink(tmp.name)
    return wb3, wb3.active


def _encontrar_header_row(ws):
    """Encontra a linha de cabeçalho no ValoresDaDRE (contém 'Competência' e 'Valor')."""
    for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
        vals = [str(c or "").strip().lower() for c in row]
        if "competência" in vals or "competencia" in vals:
            return i
    return 5  # fallback linha 5


def _get_col_indices(ws, header_row):
    """
    Retorna dict com índices das colunas importantes.
    Colunas: A=Descrição, B=Empresa, C=Centro de Custo, D=Competência, E=Valor
    """
    row_vals = list(ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True))[0]
    cols = {}
    for i, v in enumerate(row_vals, start=1):
        n = str(v or "").strip().lower()
        if "descrição" in n or "descricao" in n:
            cols["descricao"] = i
        elif "empresa" in n:
            cols["empresa"] = i
        elif "centro" in n:
            cols["centro"] = i
        elif "competência" in n or "competencia" in n:
            cols["competencia"] = i
        elif "valor" in n:
            cols["valor"] = i
    # fallback padrão
    cols.setdefault("descricao", 1)
    cols.setdefault("empresa",   2)
    cols.setdefault("centro",    3)
    cols.setdefault("competencia", 4)
    cols.setdefault("valor",     5)
    return cols


# ── Função principal: Vendas de Mercadorias ───────────────────────────────────

def atualizar_vendas_mercadorias(
    dre_path: str,
    fat_lojas_path: str,
    fat_vd_path: str,
    cmv_path: str,
    mes_display: str,
    ano: int = 2026,
    output_path: str = None
) -> dict:
    """
    Atualiza o ValoresDaDRE de Vendas de Mercadorias:
      - Col D (Competência) = primeiro dia do mês selecionado
      - Col E (Valor)       = Receita Líquida (Loja) ou soma ValorPedido (VD)
                              identificado pelo código numérico da Col C

    Retorna dict com relatório: lojas atualizadas, não encontradas, erros.
    """
    data_competencia = _mes_para_data(mes_display, ano)

    # Carregar valores das bases
    map_lojas = _carregar_fat_lojas(fat_lojas_path)   # {cod: valor}
    map_vd    = _carregar_fat_vd(fat_vd_path)         # {cod: valor}

    # Identificar canal de cada loja via CMV
    codigos_vd, codigos_loja = _identificar_canal_cmv(cmv_path)

    # Combinar: VD tem prioridade se o código aparecer lá
    # map_vd já tem os valores corretos para VD
    # map_lojas tem os valores para LOJA

    # Abrir DRE
    wb, ws = _abrir_dre(dre_path)
    header_row = _encontrar_header_row(ws)
    cols = _get_col_indices(ws, header_row)

    relatorio = {
        "atualizadas": [],
        "nao_encontradas": [],  # código no DRE mas sem valor nas bases
        "sem_codigo": [],       # linha no DRE sem código identificável
    }

    # Iterar linhas de dados (após cabeçalho)
    for row_idx in range(header_row + 1, ws.max_row + 1):
        centro_val = ws.cell(row=row_idx, column=cols["centro"]).value
        if not centro_val:
            continue

        cod = _extrair_codigo(str(centro_val))
        if not cod:
            relatorio["sem_codigo"].append(str(centro_val))
            continue

        # Determinar canal e valor
        valor = None
        canal = None

        if cod in codigos_vd:
            canal = "VD"
            valor = map_vd.get(cod)
        elif cod in codigos_loja:
            canal = "Loja"
            valor = map_lojas.get(cod)
        else:
            # Tentar nas duas bases
            if cod in map_vd:
                canal = "VD"
                valor = map_vd[cod]
            elif cod in map_lojas:
                canal = "Loja"
                valor = map_lojas[cod]

        if valor is None:
            relatorio["nao_encontradas"].append(f"{cod} (centro: {centro_val})")
            continue

        # Gravar Competência (col D) e Valor (col E)
        ws.cell(row=row_idx, column=cols["competencia"]).value = data_competencia
        ws.cell(row=row_idx, column=cols["valor"]).value       = round(valor, 2)

        relatorio["atualizadas"].append(
            f"{cod} ({canal}) → R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        )

    # Salvar
    out = output_path or dre_path
    wb.save(out)
    wb.close()

    return relatorio
