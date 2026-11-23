"""Escrita da planilha (xlwings). Todos os cálculos vêm do Python; fórmulas Excel só nas
calculadoras (consulta na grade + fórmulas fechadas) e em rótulos do input do Book."""
from __future__ import annotations

import datetime as dt
import re

import numpy as np
import pandas as pd
import xlwings as xw

from mesa import book as B
from mesa import calendario as cal
from mesa import saida, validacao
from mesa.curvas.montagem import Curvas
from mesa.dados.leitura import DadosD0
from mesa.excel import estilo as E

ABAS = ["Painel", "PRE", "IPCA", "USD", "EUR", "Book", "Feriados", "Config", "Listas"]
LINHA_GRADE = 8                        # primeira linha de dados das grades (cabeçalho na 7)
COL_GRADE = 24                         # coluna X
COL_VT = 16                            # tabela de vértices na coluna P
VT_VENC, VT_TAXA = COL_VT + 1, COL_VT + 4
INPUT_BOOK = (8, 37)                   # linhas do input do Book
COLS_INPUT = ["Família", "Ticker", "Lado", "Quantidade", "Operado hoje?", "Taxa/Preço de entrada", "Unidade", "Status"]

# ------------------------------------------------------------------ utilidades


def col(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _limpo(df: pd.DataFrame) -> list[list]:
    out = df.astype(object).where(pd.notna(df), None).values.tolist()
    return [[v.item() if isinstance(v, np.generic) else v for v in row] for row in out]


def _aba(wb: xw.Book, nome: str) -> xw.Sheet:
    nomes = [s.name for s in wb.sheets]
    if nome in nomes:
        sh = wb.sheets[nome]
    else:
        anterior = [n for n in ABAS[: ABAS.index(nome)] if n in nomes]
        sh = wb.sheets.add(nome, after=wb.sheets[anterior[-1]] if anterior else None)
    for ch in list(sh.api.ChartObjects()):
        ch.Delete()
    for shp in list(sh.api.Shapes):
        shp.Delete()
    sh.cells.api.UnMerge()
    sh.cells.clear()
    sh.cells.color = "#" + E.FUNDO
    f = sh.cells.api.Font
    f.Name, f.Size, f.Color = E.FONTE, 10, E.rgb(E.CINZA_CLARO)
    sh.range("A:A").column_width = 2
    try:
        sh.activate()
        wb.app.api.ActiveWindow.DisplayGridlines = False
    except Exception:
        pass
    return sh


def _titulo(sh, texto, sub):
    E.pinta(sh.range("B1"), cor=E.BRANCO, negrito=True, tamanho=16)
    sh.range("B1").value = texto
    E.pinta(sh.range("B2"), cor=E.CINZA, tamanho=9)
    sh.range("B2").value = sub


def _kpis(sh, linha, col_ini, itens: list[tuple[str, float, str]], largura=2):
    for k, (rot, val, fmt) in enumerate(itens):
        c = col_ini + k * largura
        r1 = sh.range((linha, c), (linha, c + largura - 1))
        r2 = sh.range((linha + 1, c), (linha + 1, c + largura - 1))
        for r in (r1, r2):
            r.api.Merge()
        E.pinta(r1, E.CARTAO, E.CINZA, tamanho=8, alinhar="centro")
        E.pinta(r2, E.CARTAO, E.BRANCO, True, 13, "centro", fmt)
        r1.value, r2.value = rot, val


def _tabela(sh, linha, col_ini, cab: list[str], df: pd.DataFrame | list[list], formatos: list[str | None] | None = None):
    dados = _limpo(df) if isinstance(df, pd.DataFrame) else df
    n = len(dados)
    h = sh.range((linha, col_ini), (linha, col_ini + len(cab) - 1))
    h.value = cab
    E.pinta(h, E.CABECALHO, E.BRANCO, True, 9, "centro")
    h.api.WrapText = True
    if n:
        corpo = sh.range((linha + 1, col_ini), (linha + n, col_ini + len(cab) - 1))
        corpo.value = dados
        E.pinta(corpo, E.LINHA, E.BRANCO, tamanho=9, alinhar="centro")
        for j, fmt in enumerate(formatos or []):
            if fmt:
                E.aplica_formato(sh.range((linha + 1, col_ini + j), (linha + n, col_ini + j)), fmt)
    return n


def _secao(sh, celula, texto):
    E.pinta(sh.range(celula), cor=E.AZUL_CLARO, negrito=True, tamanho=11)
    sh.range(celula).value = texto


def _botao(sh, esq, topo, larg, alt, texto, macro, cor=E.LARANJA):
    shp = sh.api.Shapes.AddShape(5, esq, topo, larg, alt)   # retângulo arredondado
    shp.Fill.ForeColor.RGB = E.rgb(cor)
    shp.Line.Visible = 0
    tr = shp.TextFrame2.TextRange
    tr.Text = texto
    tr.Font.Bold, tr.Font.Size, tr.Font.Name = True, 11, E.FONTE
    tr.Font.Fill.ForeColor.RGB = E.rgb(E.FUNDO)
    shp.TextFrame2.VerticalAnchor = 3
    tr.ParagraphFormat.Alignment = 2
    shp.OnAction = macro
    return shp


def _grafico(sh, ancora: str, larg, alt, titulo, series, fmt_x="mmm/yy", fmt_y="0.00", tipo_x="data"):
    """series: (nome, rng_x, rng_y, cor, estilo) — estilo 'linha', 'tracejado' ou 'pontos'."""
    if larg is None:
        larg = sh.range("F:N").width
    a = sh.range(ancora)
    ch = sh.api.Shapes.AddChart2(-1, 75, a.left, a.top, larg, alt).Chart   # 75 = dispersão com linhas
    while ch.SeriesCollection().Count:
        ch.SeriesCollection(1).Delete()
    for nome, rx, ry, cor, estilo in series:
        s = ch.SeriesCollection().NewSeries()
        s.Name, s.XValues, s.Values = nome, rx.api, ry.api
        if estilo == "pontos":
            s.ChartType = -4169                      # dispersão só marcadores
            s.MarkerStyle, s.MarkerSize = 8, 6
            s.MarkerBackgroundColor = s.MarkerForegroundColor = E.rgb(cor)
        else:
            s.ChartType = 75
            s.MarkerStyle = -4142
            ln = s.Format.Line
            ln.ForeColor.RGB, ln.Weight = E.rgb(cor), 2.25 if estilo == "linha" else 1.75
            if estilo == "tracejado":
                ln.DashStyle = 4
    ch.HasTitle = True
    ch.ChartTitle.Text = titulo
    ch.ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = E.rgb(E.BRANCO)
    ch.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 11
    ch.ChartArea.Format.Fill.ForeColor.RGB = E.rgb(E.FUNDO)
    ch.ChartArea.Format.Line.ForeColor.RGB = E.rgb(E.CABECALHO)
    ch.PlotArea.Format.Fill.ForeColor.RGB = E.rgb(E.FUNDO)
    ch.HasLegend = True
    ch.Legend.Position = -4160
    ch.Legend.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = E.rgb(E.CINZA_CLARO)
    for eixo, fmt in ((1, fmt_x), (2, fmt_y)):
        ax = ch.Axes(eixo)
        ax.TickLabels.NumberFormatLocal = E.formato_local(fmt, *E.convencoes(sh.book.app))
        ax.TickLabels.Font.Color = E.rgb(E.CINZA)
        ax.Format.Line.ForeColor.RGB = E.rgb(E.CABECALHO)
        ax.HasMajorGridlines = eixo == 2
        if eixo == 2:
            ax.MajorGridlines.Format.Line.ForeColor.RGB = E.rgb(E.CABECALHO)
    if tipo_x == "data":
        xs = [v for _, rx, *_ in series for v in (rx.value if isinstance(rx.value, list) else [rx.value]) if v]
        if xs:
            ch.Axes(1).MinimumScale = pd.Timestamp(min(xs)).toordinal() - 693594
            ch.Axes(1).MaximumScale = pd.Timestamp(max(xs)).toordinal() - 693594
    return ch


def _grade(sh, df: pd.DataFrame, cab: list[str], formatos: list[str]):
    _tabela(sh, LINHA_GRADE - 1, COL_GRADE, cab, df, formatos)
    E.pinta(sh.range((LINHA_GRADE - 2, COL_GRADE)), cor=E.AZUL_CLARO, negrito=True, tamanho=11)
    sh.range((LINHA_GRADE - 2, COL_GRADE)).value = f"GRADE DIÁRIA (todo dia útil até {saida.MAX_DC} DC)"
    for j in range(len(cab)):
        sh.range((1, COL_GRADE + j)).column_width = 12
    return LINHA_GRADE, LINHA_GRADE + len(df) - 1


def _rg(sh, j: int, ini: int, fim: int) -> xw.Range:
    """Coluna j (0 = Data) da grade, linhas ini..fim."""
    return sh.range((ini, COL_GRADE + j), (fim, COL_GRADE + j))


def _ref(j, ini, fim) -> str:
    """Endereço absoluto da coluna j (0 = Data) da grade, linhas ini..fim (para fórmulas)."""
    c = col(COL_GRADE + j)
    return f"${c}${ini}:${c}${fim}"


# ------------------------------------------------------------------ calculadora


def _calculadora(sh, linha, itens: list[tuple], titulo="CALCULADORA (data livre)"):
    """itens: (chave, rótulo, fórmula ou None, formato, valor inicial).

    Fórmula None = célula de entrada (editável). Nas fórmulas, {chave} é trocado pelo
    endereço da célula daquele item. Colunas: B (rótulo) e D (valor).
    """
    _secao(sh, f"B{linha}", titulo)
    end = {chave: f"$D${linha + 1 + k}" for k, (chave, *_) in enumerate(itens)}
    for k, (chave, rot, form, fmt, inicial) in enumerate(itens):
        r = linha + 1 + k
        sh.range(f"B{r}").value = rot
        c = sh.range(f"D{r}")
        if form is None:
            E.pinta(sh.range(f"B{r}:C{r}"), E.CARTAO, E.CINZA, tamanho=9)
            c.value = inicial
            E.pinta(c, E.INPUT, E.LARANJA, True, 10, "centro", fmt)
        else:
            E.pinta(sh.range(f"B{r}:C{r}"), E.LINHA, E.CINZA_CLARO, tamanho=9)
            c.formula = form.format(**end)
            E.pinta(c, E.LINHA, E.BRANCO, False, 10, "centro", fmt)
    return linha + 1 + len(itens)


FORA_DA_GRADE = f"fora da grade (D+1 a {saida.MAX_DC} DC)"


def _linha_da_data(cel_data: str, ini: int, fim: int) -> str:
    """Posição na grade (1 = D+1) do 1º dia útil ≥ a data informada; #N/D se a data não estiver em
    (D0, último dia da grade]. D0 = 1ª data da grade − seu DC. O IFERROR cobre datas antes da 1ª
    linha (ex.: o próprio D+1, ou um sábado logo após D0), em que o MATCH(data−1) não acha nada."""
    d, c = col(COL_GRADE), col(COL_GRADE + 2)
    d0, ultimo, rng = f"(${d}${ini}-${c}${ini})", f"${d}${fim}", f"${d}${ini}:${d}${fim}"
    return f"IF(OR({cel_data}<={d0},{cel_data}>{ultimo}),NA(),IFERROR(MATCH({cel_data}-1,{rng},1),0)+1)"


# ------------------------------------------------------------------ abas


def aba_curva(wb, nome, titulo, sub, kpis, grade, cab, formatos, vertices: pd.DataFrame | None,
              fmt_taxa, calc, graficos, notas=()):
    sh = _aba(wb, nome)
    _titulo(sh, titulo, sub)
    _kpis(sh, 4, 2, kpis)
    ini, fim = _grade(sh, grade, cab, formatos)
    for c, w in zip("BCDEFGHIJKLMNOPQRSTUVW", (22, 6, 14, 2) + (11,) * 9 + (2, 11, 11, 6, 6, 9, 9, 10, 2)):
        sh.range(f"{c}:{c}").column_width = w
    linha_calc = calc(sh, ini, fim)
    if vertices is not None:
        _secao(sh, f"{col(COL_VT)}8", "VÉRTICES (ajustes B3) · forward entre vencimentos")
        _tabela(sh, 9, COL_VT, ["Vértice", "Vencimento", "DU", "DC", "Taxa", "Forward", "Tipo"],
                vertices.assign(tipo="Original"), [None, "dd/mm/yyyy", "0", "0", fmt_taxa, fmt_taxa, None])
    for k, g in enumerate(graficos):
        g(sh, f"F{8 + k * 23}", ini, fim)
    for k, n in enumerate(notas):
        r = linha_calc + 2 + 3 * k
        c = sh.range(f"B{r}:D{r + 2}")
        c.api.Merge()
        E.pinta(c, cor=E.CINZA, tamanho=8)
        c.api.WrapText = True
        c.api.VerticalAlignment = -4160
        sh.range(f"B{r}").value = n
    return sh


def _primeira_linha(grade: pd.DataFrame, du_min: int, ini: int) -> int:
    return ini + int(np.argmax(grade.du.values >= du_min))


def escrever(wb: xw.Book, dados: DadosD0, curvas: Curvas, executado_em: dt.datetime, python_exe: str, script: str):
    g = saida.grades(dados, curvas)
    vt = saida.vertices(dados, curvas)
    k = saida.kpis(dados, g)
    ref = dados.data_ref
    padrao = cal.ajusta_dia_util(ref + dt.timedelta(days=365))
    fonte = "Fonte: ajustes e curvas B3 · indicadores B3/BCB (via IdeaFin) · calendário ANBIMA"

    # ---------------- PRE
    def calc_pre(sh, ini, fim):
        D, T, F = _ref(0, ini, fim), _ref(3, ini, fim), _ref(6, ini, fim)
        U = _ref(1, ini, fim)
        L, L2 = _linha_da_data("{data}", ini, fim), _linha_da_data("{d2}", ini, fim)
        return _calculadora(sh, 8, [
            ("data", "Data", None, "dd/mm/yyyy", padrao),
            ("noc", "Nocional no vencimento (R$)", None, "#,##0", 1_000_000),
            ("dia", "Dia útil considerado", f'=IFERROR(INDEX({D},{L}),"{FORA_DA_GRADE}")', "dd/mm/yyyy", None),
            ("du", "DU", f"=INDEX({U},{L})", "0", None),
            ("dc", "DC", f"=INDEX({_ref(2, ini, fim)},{L})", "0", None),
            ("tx", "Taxa spot PRE (% a.a.)", f"=INDEX({T},{L})", "0.000", None),
            ("fw", "Forward do trecho (% a.a.)", f"=INDEX({F},{L})", "0.000", None),
            ("pu", "PU (100.000 no venc.)", "=100000/(1+{tx}/100)^({du}/252)", "#,##0.00", None),
            ("vp", "Valor presente (R$)", "={noc}*{pu}/100000", "#,##0.00", None),
            ("dm", "Duration modificada (anos)", "=({du}/252)/(1+{tx}/100)", "0.0000", None),
            ("cv", "Convexidade", "=({du}/252)*({du}/252+1)/(1+{tx}/100)^2", "0.0000", None),
            ("dv", "DV01 (R$ por +1 bp)", "=-{vp}*{dm}*0.0001", "#,##0.00", None),
            ("d2", "Data 2 (forward entre datas)", None, "dd/mm/yyyy", cal.ajusta_dia_util(padrao + dt.timedelta(days=365))),
            ("f2", "Forward Data → Data 2 (% a.a.)",
             f'=IF({{d2}}="","",((1+INDEX({T},{L2})/100)^(INDEX({U},{L2})/252)/(1+{{tx}}/100)^({{du}}/252))'
             f'^(252/(INDEX({U},{L2})-{{du}}))-1)*100', "0.000", None),
        ])

    def graf_pre(sh, anc, ini, fim):
        n = int((vt["PRE"].dc <= saida.MAX_DC).sum())
        _grafico(sh, anc, None, 330, "Curva PRE (DI × Pré) — spot e forward", [
            ("Spot", _rg(sh, 0, ini, fim), _rg(sh, 3, ini, fim), E.AZUL, "linha"),
            ("Forward (flat forward)", _rg(sh, 0, ini, fim), _rg(sh, 6, ini, fim), E.LARANJA, "tracejado"),
            ("Vértices DI1", sh.range((10, VT_VENC), (9 + n, VT_VENC)), sh.range((10, VT_TAXA), (9 + n, VT_TAXA)), E.BRANCO, "pontos")])

    sh = aba_curva(wb, "PRE", "CURVA PRE — DI × PRÉ",
                   f"Pregão {ref:%d/%m/%Y} · Manual de Curvas B3 2.1 · Flat Forward 252 (1.4.2) · {fonte}",
                   [("CDI", k["CDI"], "0.00"), ("PRE 1 ano", k["PRE 1 ano"], "0.000"), ("PRE 5 anos", k["PRE 5 anos"], "0.000"),
                    ("Inclinação 1a–5a (bps)", k["Inclinação 1a–5a (bps)"], "+0;-0")],
                   g["PRE"], ["Data", "DU", "DC", "Spot (% a.a.)", "Fator", "PU", "Forward (% a.a.)", "Tipo"],
                   ["dd/mm/yyyy", "0", "0", "0.000", "0.000000", "#,##0.00", "0.000", None], vt["PRE"], "0.000",
                   calc_pre, [graf_pre])

    # ---------------- IPCA
    def calc_ipca(sh, ini, fim):
        D = _ref(0, ini, fim)
        L = _linha_da_data("{data}", ini, fim)
        return _calculadora(sh, 8, [
            ("data", "Data", None, "dd/mm/yyyy", padrao),
            ("noc", "Nocional no vencimento (R$ reais)", None, "#,##0", 1_000_000),
            ("dia", "Dia útil considerado", f'=IFERROR(INDEX({D},{L}),"{FORA_DA_GRADE}")', "dd/mm/yyyy", None),
            ("du", "DU", f"=INDEX({_ref(1, ini, fim)},{L})", "0", None),
            ("cp", "Cupom IPCA (% a.a.)", f"=INDEX({_ref(3, ini, fim)},{L})", "0.00", None),
            ("pre", "PRE (% a.a.)", f"=INDEX({_ref(4, ini, fim)},{L})", "0.000", None),
            ("inf", "Inflação implícita (% a.a.)", f"=INDEX({_ref(5, ini, fim)},{L})", "0.000", None),
            ("fw", "Forward cupom do trecho (% a.a.)", f"=INDEX({_ref(7, ini, fim)},{L})", "0.00", None),
            ("pu", "PU DAP (100.000 no venc.)", "=100000/(1+{cp}/100)^({du}/252)", "#,##0.00", None),
            ("dm", "Duration modificada (anos)", "=({du}/252)/(1+{cp}/100)", "0.0000", None),
            ("cv", "Convexidade", "=({du}/252)*({du}/252+1)/(1+{cp}/100)^2", "0.0000", None),
            ("dv", "DV01 cupom (R$ por +1 bp)", "=-{noc}*{pu}/100000*{dm}*0.0001", "#,##0.00", None),
        ])

    p_ipca = _primeira_linha(g["IPCA"], int(curvas.inflacao.du[1]), LINHA_GRADE)

    def graf_ipca(sh, anc, ini, fim):
        n = int((vt["IPCA"].dc <= saida.MAX_DC).sum())
        _grafico(sh, anc, None, 330, "Cupom IPCA (DPL) — spot, forward e inflação implícita", [
            ("Cupom IPCA spot", _rg(sh, 0, p_ipca, fim), _rg(sh, 3, p_ipca, fim), E.AZUL, "linha"),
            ("Forward cupom", _rg(sh, 0, p_ipca + 1, fim), _rg(sh, 7, p_ipca + 1, fim), E.LARANJA, "tracejado"),
            ("Inflação implícita", _rg(sh, 0, p_ipca, fim), _rg(sh, 5, p_ipca, fim), E.VERDE, "linha"),
            ("Vértices DAP", sh.range((11, VT_VENC), (9 + n, VT_VENC)), sh.range((11, VT_TAXA), (9 + n, VT_TAXA)), E.BRANCO, "pontos")])

    aba_curva(wb, "IPCA", "CUPOM IPCA — DI × IPCA (DPL)",
              f"Pregão {ref:%d/%m/%Y} · Manual de Curvas B3 3.2 · FF252 sobre a inflação implícita · vértice 1 publicado pela B3",
              [("Cupom IPCA 2 anos", k["Cupom IPCA 2 anos"], "0.00"), ("Inflação implícita 2 anos", k["Inflação implícita 2 anos"], "0.00"),
               ("PRE 1 ano", k["PRE 1 ano"], "0.000"), ("CDI", k["CDI"], "0.00")],
              g["IPCA"], ["Data", "DU", "DC", "Cupom IPCA (% a.a.)", "PRE (% a.a.)", "Inflação implícita (% a.a.)", "PU DAP",
                          "Forward cupom (% a.a.)", "Tipo"],
              ["dd/mm/yyyy", "0", "0", "0.00", "0.000", "0.000", "#,##0.00", "0.00", None], vt["IPCA"], "0.00",
              calc_ipca, [graf_ipca],
              ["Gráfico a partir do 1º vencimento DAP; antes dele a B3 usa a projeção mensal de IPCA (dado não disponível)."])

    # ---------------- USD
    dol_aj = saida.ajustes_moeda(dados, "DOL")

    def calc_usd(sh, ini, fim):
        D = _ref(0, ini, fim)
        L = _linha_da_data("{data}", ini, fim)
        pre = f"INDEX(PRE!{_ref(3, ini, fim)},{L})"
        return _calculadora(sh, 8, [
            ("data", "Data", None, "dd/mm/yyyy", padrao),
            ("noc", "Nocional (US$)", None, "#,##0", 1_000_000),
            ("dia", "Dia útil considerado", f'=IFERROR(INDEX({D},{L}),"{FORA_DA_GRADE}")', "dd/mm/yyyy", None),
            ("du", "DU", f"=INDEX({_ref(1, ini, fim)},{L})", "0", None),
            ("dc", "DC", f"=INDEX({_ref(2, ini, fim)},{L})", "0", None),
            ("cs", "Cupom sujo (% a.a., lin 360)", f"=INDEX({_ref(3, ini, fim)},{L})", "0.000", None),
            ("cl", "Cupom limpo (% a.a., lin 360)", f"=INDEX({_ref(4, ini, fim)},{L})", "0.00", None),
            ("fw", "Forward cupom limpo do trecho", f"=INDEX({_ref(5, ini, fim)},{L})", "0.00", None),
            ("fx", "Dólar forward (R$/US$)", f"=INDEX({_ref(6, ini, fim)},{L})", "0.0000", None),
            ("val", "Valor em R$ no vencimento", "={noc}*{fx}", "#,##0.00", None),
            ("dvp", "DV01 PRE (R$ por +1 bp)", f"={{val}}*({{du}}/252)/(1+{pre}/100)*0.0001", "#,##0.00", None),
            ("dvc", "DV01 cupom (R$ por +1 bp)", "=-{val}*({dc}/36000)/(1+{cs}*{dc}/36000)*0.01", "#,##0.00", None),
            ("dfx", "Delta spot (R$ por +1%)", "={val}*0.01", "#,##0.00", None),
        ])

    p_usd = _primeira_linha(g["USD"], int(curvas.doc.du[1]), LINHA_GRADE)

    def graf_cupom_usd(sh, anc, ini, fim):
        n = int((vt["USD"].dc <= saida.MAX_DC).sum())
        _grafico(sh, anc, None, 330, "Cupom cambial limpo USD — spot e forward (FRC)", [
            ("Cupom limpo", _rg(sh, 0, p_usd, fim), _rg(sh, 4, p_usd, fim), E.AZUL, "linha"),
            ("Forward cupom limpo", _rg(sh, 0, p_usd + 1, fim), _rg(sh, 5, p_usd + 1, fim), E.LARANJA, "tracejado"),
            ("Vértices DDI (limpo)", sh.range((11, VT_VENC), (9 + n, VT_VENC)), sh.range((11, VT_TAXA), (9 + n, VT_TAXA)), E.BRANCO, "pontos")])

    def graf_dolar(sh, anc, ini, fim):
        base = 9 + len(vt["USD"]) + 3
        nd = int((dol_aj.dc <= saida.MAX_DC).sum())
        _grafico(sh, anc, None, 330, "Dólar forward (R$/US$) × ajustes DOL", [
            ("Dólar forward (paridade)", _rg(sh, 0, ini, fim), _rg(sh, 6, ini, fim), E.AZUL, "linha"),
            ("Ajustes DOL", sh.range((base + 1, VT_VENC), (base + nd, VT_VENC)),
             sh.range((base + 1, VT_TAXA), (base + nd, VT_TAXA)), E.LARANJA, "pontos")], fmt_y="0.00")

    sh = aba_curva(wb, "USD", "CUPOM CAMBIAL E DÓLAR FORWARD",
                   f"Pregão {ref:%d/%m/%Y} · Manual de Curvas B3 4.4/4.5/4.1 · FF252 com convenção linear (1.4.3) · Ptax D-1 {k['Ptax D-1']:.4f}",
                   [("Cupom limpo 1 ano", k["Cupom limpo USD 1 ano"], "0.00"), ("Dólar forward 1 ano", k["Dólar forward 1 ano"], "0.0000"),
                    ("Ptax D-1", k["Ptax D-1"], "0.0000"), ("Dólar cupom limpo", k["Dólar cupom limpo (spot)"], "0.0000")],
                   g["USD"], ["Data", "DU", "DC", "Cupom sujo (% a.a.)", "Cupom limpo (% a.a.)", "Forward cupom limpo",
                              "Dólar forward (R$/US$)", "Tipo"],
                   ["dd/mm/yyyy", "0", "0", "0.000", "0.00", "0.00", "0.0000000", None], vt["USD"], "0.00",
                   calc_usd, [graf_cupom_usd, graf_dolar],
                   notas=["Gráfico de cupom a partir do 1º DDI: no vértice 1 o cupom anualiza a variação da Ptax em 1 dia (pode passar de 100% a.a.).",
                          "Forward de cupom = FRC: calculado dos DDIs (3 casas), sem o arredondamento de 2 casas da curva DOC."])
    base = 9 + len(vt["USD"]) + 3
    _secao(sh, f"{col(COL_VT)}{base - 1}", "AJUSTES DOL (R$/US$ 1.000)")
    _tabela(sh, base, COL_VT, ["Ticker", "Vencimento", "DU", "DC", "Ajuste", "Contr. abertos"],
            dol_aj.assign(pa=dol_aj.pa / 1000), [None, "dd/mm/yyyy", "0", "0", "0.0000", "#,##0"])

    # ---------------- EUR
    eur_aj = saida.ajustes_moeda(dados, "EUR")

    def calc_eur(sh, ini, fim):
        D = _ref(0, ini, fim)
        L = _linha_da_data("{data}", ini, fim)
        pre = f"INDEX(PRE!{_ref(3, ini, fim)},{L})"
        ecc = f"INDEX({_ref(4, ini, fim)},{L})"
        return _calculadora(sh, 8, [
            ("data", "Data", None, "dd/mm/yyyy", padrao),
            ("noc", "Nocional (€)", None, "#,##0", 1_000_000),
            ("dia", "Dia útil considerado", f'=IFERROR(INDEX({D},{L}),"{FORA_DA_GRADE}")', "dd/mm/yyyy", None),
            ("du", "DU", f"=INDEX({_ref(1, ini, fim)},{L})", "0", None),
            ("dc", "DC", f"=INDEX({_ref(2, ini, fim)},{L})", "0", None),
            ("cl", "Cupom limpo EUR (% a.a.)", f"=INDEX({_ref(3, ini, fim)},{L})", "0.0000", None),
            ("cs", "Cupom sujo EUR (≥ 15 DC)", f'=IF({ecc}="","n/d",{ecc})', "0.0000", None),
            ("fx", "Euro forward (R$/€)", f"=INDEX({_ref(6, ini, fim)},{L})", "0.0000", None),
            ("val", "Valor em R$ no vencimento", "={noc}*{fx}", "#,##0.00", None),
            ("dvp", "DV01 PRE (R$ por +1 bp)", f"={{val}}*({{du}}/252)/(1+{pre}/100)*0.0001", "#,##0.00", None),
            ("dvc", "DV01 cupom EUR (R$ por +1 bp)", "=-{val}*({dc}/36000)/(1+{cl}*{dc}/36000)*0.01", "#,##0.00", None),
            ("dfx", "Delta spot (R$ por +1%)", "={val}*0.01", "#,##0.00", None),
        ])

    p_eur = _primeira_linha(g["EUR"], int(curvas.doc.du[1]), LINHA_GRADE)

    def graf_cupom_eur(sh, anc, ini, fim):
        _grafico(sh, anc, None, 330, "Cupom limpo EUR (EUC) — spot e forward", [
            ("Cupom limpo EUR", _rg(sh, 0, p_eur, fim), _rg(sh, 3, p_eur, fim), E.AZUL, "linha"),
            ("Forward cupom limpo", _rg(sh, 0, p_eur + 1, fim), _rg(sh, 5, p_eur + 1, fim), E.LARANJA, "tracejado")])

    def graf_euro(sh, anc, ini, fim):
        base = 9 + len(vt["EUR"]) + 3
        _grafico(sh, anc, None, 330, "Euro forward (R$/€) × ajustes EUR", [
            ("Euro forward (paridade)", _rg(sh, 0, ini, fim), _rg(sh, 6, ini, fim), E.AZUL, "linha"),
            ("Ajustes EUR", sh.range((base + 1, VT_VENC), (base + len(eur_aj), VT_VENC)),
             sh.range((base + 1, VT_TAXA), (base + len(eur_aj), VT_TAXA)), E.LARANJA, "pontos")], fmt_y="0.00")

    sh = aba_curva(wb, "EUR", "CUPOM DE EURO E EURO FORWARD",
                   f"Pregão {ref:%d/%m/%Y} · Manual de Curvas B3 4.7/4.2 · EUC = DOC ⊗ SDE · spot TXEURO {dados.indicadores['euro_cupom_limpo']:.7f}",
                   [("Cupom EUR 1 ano", k["Cupom EUR 1 ano"], "0.0000"), ("Euro forward 1 ano", k["Euro forward 1 ano"], "0.0000"),
                    ("EUR/USD", dados.indicadores["eurusd"], "0.0000"), ("Euro cupom limpo", dados.indicadores["euro_cupom_limpo"], "0.0000")],
                   g["EUR"], ["Data", "DU", "DC", "Cupom limpo EUR (% a.a.)", "Cupom sujo EUR (% a.a.)", "Forward cupom limpo",
                              "Euro forward (R$/€)", "SDE", "Tipo"],
                   ["dd/mm/yyyy", "0", "0", "0.0000", "0.0000", "0.0000", "0.0000000", "0.00", None], vt["EUR"], "0.0000",
                   calc_eur, [graf_cupom_eur, graf_euro],
                   notas=["Cupom sujo EUR: curva não documentada no Manual B3; exibida a partir de 15 DC (antes, a B3 usa um spot de euro não disponível).",
                          "Futuros EUR: ilíquidos; ajustes exibidos só como referência. A curva EUR da B3 sai de DDI, SDE e "
                          "TXEURO (não dos futuros) e costuma ficar acima dos ajustes EUR (até ~0,07%)."])
    base = 9 + len(vt["EUR"]) + 3
    _secao(sh, f"{col(COL_VT)}{base - 1}", "AJUSTES EUR (R$/€ 1.000) — referência")
    _tabela(sh, base, COL_VT, ["Ticker", "Vencimento", "DU", "DC", "Ajuste", "Contr. abertos"],
            eur_aj.assign(pa=eur_aj.pa / 1000), [None, "dd/mm/yyyy", "0", "0", "0.0000", "#,##0"])

    # ---------------- Feriados / Config / Listas
    shf = _aba(wb, "Feriados")
    _titulo(shf, "FERIADOS NACIONAIS (ANBIMA)", "Base dos dias úteis (DU/252). Fonte: anbima.com.br/feriados")
    fer = pd.read_csv(cal.ARQ_FERIADOS, parse_dates=["data"])
    fer = fer[fer.data.dt.year.between(ref.year, ref.year + 12)]
    _tabela(shf, 4, 2, ["Data", "Dia da semana", "Feriado"], fer.assign(data=fer.data.dt.date), ["dd/mm/yyyy", None, None])
    for c, w in zip("BCD", (12, 16, 40)):
        shf.range(f"{c}:{c}").column_width = w

    # Config é lida pela macro dos botões: atualiza no lugar (sem apagar a aba), para a macro nunca ler C4/C5 vazios
    shc = wb.sheets["Config"] if "Config" in [s.name for s in wb.sheets] else _aba(wb, "Config")
    _titulo(shc, "CONFIGURAÇÃO", "Preenchida pelo main.py — usada pelos botões. Não editar.")
    shc.range("B4:B6").value = [["Python"], ["main.py"], ["Pregão carregado"]]
    shc.range("C6").api.NumberFormat = "@"          # texto: impede o Excel de converter em data
    shc.range("C4:C6").value = [[python_exe], [script], [ref.isoformat()]]
    E.pinta(shc.range("B4:B6"), E.CARTAO, E.CINZA, tamanho=9)
    E.pinta(shc.range("C4:C6"), E.LINHA, E.BRANCO, tamanho=9)
    shc.range("B:B").column_width, shc.range("C:C").column_width = 18, 90
    E.pinta(shc.range("B9"), cor=E.LARANJA, tamanho=9)
    shc.range("B9").value = ("Privacidade: C4 e C5 guardam caminhos desta máquina (incluem o seu usuário do Windows). "
                             "Antes de enviar a planilha a alguém, apague C4:C5 — o main.py regrava na próxima execução.")

    _listas(wb, dados)
    _painel(wb, dados, curvas, g, k, executado_em)


def _listas(wb, dados: DadosD0):
    sh = _aba(wb, "Listas")
    for n in list(wb.names):
        if n.name.startswith(("Lista_", "Lado_", "Familias", "SimNao")):
            n.delete()
    colunas = {"DI1": None, "DAP": None, "DDI": None, "FRC": B.tickers_frc(dados), "DOL": None, "EUR": None}
    for j, (fam, tks) in enumerate(colunas.items(), start=1):
        if tks is None:
            x = dados.ajustes[(dados.ajustes.familia == fam) & (dados.ajustes.du > 0)].sort_values("du")
            tks = x.ticker.tolist()
        sh.range((1, j)).value = fam
        sh.range((2, j)).options(transpose=True).value = tks or ["—"]
        wb.names.add(f"Lista_{fam}", f"=Listas!${col(j)}$2:${col(j)}${1 + max(len(tks), 1)}")
        c = 10 + j
        sh.range((1, c)).value = fam
        sh.range((2, c)).options(transpose=True).value = list(B.LADOS[fam])
        wb.names.add(f"Lado_{fam}", f"=Listas!${col(c)}$2:${col(c)}$3")
    sh.range("H1").value = "Familias"
    sh.range("H2").options(transpose=True).value = list(B.LADOS)
    wb.names.add("Familias", f"=Listas!$H$2:$H${1 + len(B.LADOS)}")
    sh.range("I1").value = "SimNao"
    sh.range("I2").options(transpose=True).value = ["Sim", "Não"]
    wb.names.add("SimNao", "=Listas!$I$2:$I$3")
    sh.api.Visible = 0


def _painel(wb, dados, curvas, g, k, executado_em):
    sh = _aba(wb, "Painel")
    ref = dados.data_ref
    _titulo(sh, "MESA DE CURVAS B3", "Curvas montadas a partir dos ajustes da B3 · metodologia do Manual de Curvas B3 · sempre o último pregão (D0)")
    _botao(sh, sh.range("J1").left, 4, 250, 30, "▶  RODAR PREGÃO DO DIA", "RodarPregaoDoDia")
    E.pinta(sh.range("J3:M3"), cor=E.CINZA, tamanho=8)
    sh.range("J3").value = "Baixa os ajustes do último pregão e recalcula curvas, book e risco."
    E.pinta(sh.range("B4:H4"), E.CARTAO, E.VERDE, True, 10)
    sh.range("B4").value = "Pregão de referência"
    status = f"{ref:%d/%m/%Y}  ·  rodado em {executado_em:%d/%m/%Y %H:%M}  ·  ✔ 3 arquivos do mesmo pregão"
    sh.range("C4").value = status
    wb.sheets["Config"].range("B7:C7").value = ["Status do pregão", status]
    E.pinta(sh.range("B4"), E.CARTAO, E.CINZA, tamanho=9)
    _kpis(sh, 6, 2, [("CDI", k["CDI"], "0.00"), ("PRE 1 ano", k["PRE 1 ano"], "0.000"), ("PRE 5 anos", k["PRE 5 anos"], "0.000"),
                     ("Inclinação 1a–5a", k["Inclinação 1a–5a (bps)"], '+0" bps";-0" bps"'),
                     ("Inflação implícita 2a", k["Inflação implícita 2 anos"], "0.00")])
    _kpis(sh, 9, 2, [("Cupom limpo USD 1a", k["Cupom limpo USD 1 ano"], "0.00"), ("Dólar forward 1a", k["Dólar forward 1 ano"], "0.0000"),
                     ("Ptax D-1", k["Ptax D-1"], "0.0000"), ("Cupom EUR 1a", k["Cupom EUR 1 ano"], "0.0000"),
                     ("Euro forward 1a", k["Euro forward 1 ano"], "0.0000")])
    ini, fim = LINHA_GRADE, LINHA_GRADE + len(g["PRE"]) - 1
    pre, ipca = wb.sheets["PRE"], wb.sheets["IPCA"]
    p_ipca = _primeira_linha(g["IPCA"], int(curvas.inflacao.du[1]), LINHA_GRADE)
    _grafico(sh, "B12", sh.range("B:K").width - 6, 320, "PRE × cupom IPCA × inflação implícita", [
        ("PRE", _rg(pre, 0, ini, fim), _rg(pre, 3, ini, fim), E.AZUL, "linha"),
        ("Cupom IPCA", _rg(ipca, 0, p_ipca, fim), _rg(ipca, 3, p_ipca, fim), E.LARANJA, "linha"),
        ("Inflação implícita", _rg(ipca, 0, p_ipca, fim), _rg(ipca, 5, p_ipca, fim), E.VERDE, "tracejado")])
    v = validacao.contra_b3(dados, curvas)
    _secao(sh, "M12", "VALIDAÇÃO × CURVAS OFICIAIS B3 (vértice a vértice)")
    passou = ((v.dentro == v.vertices) & (v.vertices > 0)).values     # 0 vértices comparáveis = ✘, não ✔
    v["ok"] = np.where(passou, "✔", "✘")
    _tabela(sh, 13, 13, ["Curva", "Vértices", "Dentro da tolerância", "Maior diferença", "Tolerância", ""],
            v[["curva", "vertices", "dentro", "max_dif", "tolerancia", "ok"]], [None, "0", "0", "0.0000000", "0.0000000", None])
    for k, ok in enumerate(passou):                  # ✔ verde, ✘ vermelho (linha a linha)
        sh.range((14 + k, 18)).api.Font.Color = E.rgb(E.VERDE if ok else E.VERMELHO)
    p = validacao.paridade_dol(dados, curvas)
    lin = 13 + len(v) + 3
    _secao(sh, f"M{lin - 1}", "PARIDADE: DÓLAR TEÓRICO (curvas) × AJUSTE DOL (pontos)")
    _tabela(sh, lin, 13, ["Ticker", "Vencimento", "DU", "Ajuste B3", "Teórico", "Diferença"],
            p, [None, "dd/mm/yyyy", "0", "#,##0.000", "#,##0.000", "+0.000;-0.000"])
    for c, w in zip("BCDEFGHIJKLMNOPQR", (13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 2, 26, 11, 13, 13, 12, 10)):
        sh.range(f"{c}:{c}").column_width = w
    E.pinta(sh.range("B36"), cor=E.CINZA, tamanho=8)
    sh.range("B36").value = ("Dados: ajustes e curvas referenciais B3, indicadores B3/BCB (arquivos diários públicos via IdeaFin), "
                             "feriados ANBIMA. Cálculos em Python; metodologia e itens do manual citados em cada aba.")
    sh.activate()


# ------------------------------------------------------------------ Book


def ler_input_book(wb) -> list[list]:
    if "Book" not in [s.name for s in wb.sheets]:
        return []
    a, b = INPUT_BOOK
    return wb.sheets["Book"].range(f"B{a}:G{b}").value


_NUMERO = re.compile(r"[+-]?\d+(?:,\d+)?")      # texto: inteiro ou vírgula decimal ("1.000" seria ambíguo)


def _numero(v) -> float | None:
    """Número digitado no input: None se vazio; NaN se for texto que não é número (o Book acusa na linha).
    Aceita número guardado como texto ("10", "13,5")."""
    if v in (None, ""):
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    s = str(v).strip()
    return float(s.replace(",", ".")) if _NUMERO.fullmatch(s) else float("nan")


def operacoes(valores: list[list]) -> list[tuple[int, B.Operacao]]:
    ops = []
    for n, (fam, tk, lado, qtd, hoje, entrada) in enumerate(valores):
        if not any(v not in (None, "") for v in (fam, tk, lado, qtd)):
            continue
        q = _numero(qtd)
        ops.append((n, B.Operacao(str(fam or "").strip().upper(), str(tk or "").strip().upper(),
                                  str(lado or "").strip().capitalize(), 0.0 if q is None else q,
                                  str(hoje).strip().lower() == "sim", _numero(entrada))))
    return ops


def exemplo_book(dados: DadosD0) -> list[list]:
    def primeiro(fam, du_min):
        x = dados.ajustes[(dados.ajustes.familia == fam) & (dados.ajustes.du >= du_min)].sort_values("du")
        return x.ticker.iloc[0]
    return [["DI1", primeiro("DI1", 252), "Tomado", 100, "Não", None],
            ["DI1", primeiro("DI1", 1000), "Dado", 50, "Não", None],
            ["DAP", primeiro("DAP", 1500), "Tomado", 30, "Não", None],
            ["DDI", primeiro("DDI", 252), "Dado", 20, "Não", None],
            ["DOL", primeiro("DOL", 2), "Compra", 10, "Não", None],
            ["EUR", primeiro("EUR", 2), "Venda", 5, "Não", None]]


def escrever_book(wb, dados: DadosD0, curvas: Curvas, valores: list[list] | None, executado_em: dt.datetime):
    a, b = INPUT_BOOK
    valores = valores if valores and any(any(v not in (None, "") for v in r) for r in valores) else exemplo_book(dados)
    valores = (valores + [[None] * 6] * (b - a + 1))[: b - a + 1]
    sh = _aba(wb, "Book")
    _titulo(sh, "BOOK DE FUTUROS B3", f"Pregão {dados.data_ref:%d/%m/%Y} · P&L do dia pelo ajuste B3 · risco em R$ (DV01 = +1 bp)")
    _botao(sh, sh.range("K1").left, 4, 260, 30, "↻  RECALCULAR BOOK (sem baixar dados)", "RecalcularBook", E.AZUL)
    E.pinta(sh.range("B4:I4"), cor=E.CINZA, tamanho=8)
    sh.range("B4").value = ("Juros/cupom: lado em TAXA (Tomado = comprado em taxa = vendido em PU). Moedas: Compra/Venda do preço. "
                            "FRC: só operação do dia, a partir do vencimento seguinte ao DDI base.")
    _secao(sh, "B6", "INPUT DE POSIÇÕES")
    h = sh.range(f"B{a - 1}:I{a - 1}")
    h.value = COLS_INPUT
    E.pinta(h, E.CABECALHO, E.BRANCO, True, 9, "centro")
    inp = sh.range(f"B{a}:G{b}")
    inp.value = valores
    E.pinta(inp, E.INPUT, E.BRANCO, tamanho=9, alinhar="centro")
    E.pinta(sh.range(f"H{a}:I{b}"), E.LINHA, E.CINZA_CLARO, tamanho=9, alinhar="centro")
    sh.range(f"H{a}:H{b}").formula = f'=IF(B{a}="","",IF(OR(B{a}="DOL",B{a}="EUR"),"R$/1.000","% a.a."))'
    _validacoes(sh, a, b)

    ops = operacoes(valores)
    r = B.calcular([o for _, o in ops], dados, curvas)
    status = [[None] for _ in range(b - a + 1)]
    for (n, _), s in zip(ops, r.posicoes.status):
        status[n] = [s]
    sh.range(f"I{a}:I{b}").value = status
    for k in range(b - a + 1):
        if status[k][0] and status[k][0] != "OK":
            sh.range(f"I{a + k}").api.Font.Color = E.rgb(E.VERMELHO)
        elif status[k][0] == "OK":
            sh.range(f"I{a + k}").api.Font.Color = E.rgb(E.VERDE)

    pos = r.posicoes[r.posicoes.status == "OK"].copy() if len(r.posicoes) else r.posicoes
    lin = b + 3
    _secao(sh, f"B{lin}", "POSIÇÕES · MtM e P&L do dia")
    cols = ["familia", "ticker", "lado", "quantidade", "vencimento", "du", "preco_ajuste", "taxa_ajuste", "pnl_dia", "exposicao",
            "dv01", "dv01_pre", "dv01_cupom", "delta_fx_1pct", "duration_mod", "convexidade"]
    cab = ["Família", "Ticker", "Lado", "Qtd", "Vencimento", "DU", "Preço ajuste", "Taxa ajuste", "P&L do dia (R$)", "Exposição (R$)",
           "DV01 (R$)", "DV01 PRE (R$)", "DV01 cupom (R$)", "Delta FX +1% (R$)", "Duration mod.", "Convexidade"]
    fmts = [None, None, None, "0", "dd/mm/yyyy", "0", "#,##0.000", "0.000", "#,##0.00;-#,##0.00", "#,##0", "#,##0.00;-#,##0.00",
            "#,##0.00;-#,##0.00", "#,##0.00;-#,##0.00", "#,##0.00;-#,##0.00", "0.0000", "0.0000"]
    corpo = pos[cols] if len(pos) else pd.DataFrame(columns=cols)
    n = _tabela(sh, lin + 1, 2, cab, corpo, fmts)
    tot = lin + 2 + n
    sh.range(f"B{tot}").value = "TOTAL"
    for j in (8, 9, 10, 11, 12, 13):
        c = col(2 + j)
        sh.range(f"{c}{tot}").formula = f"=SUM({c}{lin + 2}:{c}{lin + 1 + max(n, 1)})"
        E.aplica_formato(sh.range(f"{c}{tot}"), fmts[j])
    E.pinta(sh.range(f"B{tot}:Q{tot}"), E.CARTAO, E.BRANCO, True, 9)
    E.pinta(sh.range(f"B{tot + 1}"), cor=E.CINZA, tamanho=8)
    sh.range(f"B{tot + 1}").value = f"DAP: valor do ponto = R$ 0,00025 × IPCA pro rata do dia (PRT) — {r.prt_origem}"

    lin = tot + 3
    _secao(sh, f"B{lin}", "RISCO POR FATOR (R$)")
    rf = r.risco_fator.assign(medida=np.where(r.risco_fator.fator.str.startswith("Spot"), "Delta +1%", "DV01 +1 bp"))
    _tabela(sh, lin + 1, 2, ["Fator", "Risco (R$)", "Medida"], rf[["fator", "risco", "medida"]], [None, "#,##0.00;-#,##0.00", None])
    _secao(sh, f"F{lin}", "DV01 POR VÉRTICE (R$)")
    rv = r.risco_vertice.assign(rotulo=r.risco_vertice.fator + " · " + r.risco_vertice.vertice)
    nv = _tabela(sh, lin + 1, 6, ["Fator", "Vértice", "DV01 (R$)"], rv[["fator", "vertice", "dv01"]], [None, None, "#,##0.00;-#,##0.00"])
    _secao(sh, f"J{lin}", "CENÁRIOS DE CHOQUE POR FATOR (R$)")
    nc = _tabela(sh, lin + 1, 10, ["Cenário", "Reprecificação cheia", "Aprox. duration + convexidade", "Erro da aproximação"],
                 r.cenarios, [None] + ["#,##0.00;-#,##0.00"] * 3)
    nota = sh.range(f"J{lin + 3 + nc}:M{lin + 5 + nc}")      # bloco com quebra de linha (não vaza para a direita)
    nota.api.Merge()
    E.pinta(nota, cor=E.CINZA, tamanho=8)
    nota.api.WrapText = True
    nota.api.VerticalAlignment = -4160
    sh.range(f"J{lin + 3 + nc}").value = ("Cada cenário move um fator e mantém os demais. Steepening: −50 bps no curto até +50 bps "
                                          "em 10 anos (2.520 DU); Flattening: o inverso. Spot: aproximação = delta (exata: o futuro "
                                          "de moeda é linear no spot).")
    if nv:
        rot = sh.range((lin + 2, 19), (lin + 1 + nv, 19))
        rot.value = [[x] for x in rv.rotulo]
        rot.api.Font.Color = E.rgb(E.FUNDO)
        a_ = sh.range("K6")
        shp = sh.api.Shapes.AddChart2(-1, 57, a_.left, a_.top, sh.range("K:Q").width, 330)   # 57 = barras horizontais
        ch = shp.Chart
        while ch.SeriesCollection().Count:
            ch.SeriesCollection(1).Delete()
        s = ch.SeriesCollection().NewSeries()
        s.Name, s.XValues, s.Values = "DV01 (R$)", rot.api, sh.range((lin + 2, 8), (lin + 1 + nv, 8)).api
        s.Format.Fill.ForeColor.RGB = E.rgb(E.AZUL)
        s.InvertIfNegative = False
        ch.HasTitle, ch.HasLegend = True, False
        ch.ChartTitle.Text = "DV01 por vértice (R$ por +1 bp)"
        ch.ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = E.rgb(E.BRANCO)
        ch.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 11
        ch.ChartArea.Format.Fill.ForeColor.RGB = E.rgb(E.FUNDO)
        ch.PlotArea.Format.Fill.ForeColor.RGB = E.rgb(E.FUNDO)
        for eixo in (1, 2):
            ch.Axes(eixo).TickLabels.Font.Color = E.rgb(E.CINZA_CLARO)
            ch.Axes(eixo).TickLabels.Font.Size = 8
        ch.Axes(1).TickLabelSpacing = 1
        ch.Axes(1).TickLabelPosition = -4134          # rótulos à esquerda, fora das barras
        ch.Axes(2).MajorGridlines.Format.Line.ForeColor.RGB = E.rgb(E.CABECALHO)
        ch.Axes(2).TickLabels.NumberFormatLocal = E.formato_local("#,##0", *E.convencoes(sh.book.app))
    for c, w in zip("BCDEFGHIJKLMNOPQ", (11, 10, 9, 9, 12, 9, 13, 11, 18, 15, 13, 13, 13, 14, 11, 11)):
        sh.range(f"{c}:{c}").column_width = w
    E.pinta(sh.range("K4"), cor=E.CINZA, tamanho=8)
    sh.range("K4").value = f"Book recalculado em {executado_em:%d/%m/%Y %H:%M}"
    return r


def _nome_local(sh, funcao: str) -> str:
    """Nome da função no idioma do Excel do usuário (ex.: INDIRECT → INDIRETO em Excel pt-BR)."""
    t = sh.range("XFD1")
    t.formula = f'={funcao}("A1")'
    nome = str(t.api.FormulaLocal)[1:].split("(")[0]
    t.clear()
    return nome or funcao


def _validacoes(sh, a, b):
    """Listas suspensas em cascata: Família → Ticker/Lado (as validações do Excel usam o idioma local)."""
    def lista(rng, formula):
        v = rng.api.Validation
        v.Delete()
        v.Add(3, 1, 1, formula)
        v.IgnoreBlank = True
    indireto = _nome_local(sh, "INDIRECT")
    lista(sh.range(f"B{a}:B{b}"), "=Familias")
    lista(sh.range(f"C{a}:C{b}"), f'={indireto}("Lista_"&$B{a})')
    lista(sh.range(f"D{a}:D{b}"), f'={indireto}("Lado_"&$B{a})')
    lista(sh.range(f"F{a}:F{b}"), "=SimNao")
