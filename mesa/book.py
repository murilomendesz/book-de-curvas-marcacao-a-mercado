"""Book de futuros B3: validação do input, MtM/P&L do dia, risco e cenários.

Convenções
- Contratos de taxa (DI1, DAP, DDI, FRC): lado em taxa — "Tomado" (comprado em
  taxa = vendido em PU) ou "Dado" (vendido em taxa = comprado em PU).
  Moedas (DOL, EUR): "Compra" / "Venda" do preço.
- Sinais em R$ do ponto de vista da posição: DV01 = ΔMtM para +1 bp; tomado em
  DI1 → DV01 positivo.
- P&L do dia (posição de D-1): (PA hoje − PA de abertura) × valor do ponto × N,
  com o PA de abertura do arquivo B3 = PA de D-1 × FC (fator de correção da B3:
  CDI no DI1; CDI ÷ variação da Ptax no DDI; CDI ÷ IPCA pro rata no DAP; nenhum
  em DOL/EUR) — confirmado com os pregões consecutivos de 16/11 e 17/11/2026
  (teste `test_pa_de_abertura_e_o_pa_de_d1_corrigido`). Operação de hoje:
  (PA − preço da operação).
- DAP: valor do ponto = R$ 0,00025 × PRT_t (IPCA pro rata do dia, ver `ipca_pro_rata`).
- FRC (especificação B3): não gera posição própria; vira ponta longa no DDI do
  vencimento (mesma natureza, qtd do FRC) e ponta curta no DDI base (natureza
  inversa, qtd q1 = q / (1 + FRC·(n2−n1)/36000), registrada pelo preço de ajuste do
  dia). O Book exibe o FRC como uma linha, calculada pela soma das pontas. FRC só é
  aceito como **operação do dia** (o Book não reconstrói FRC de pregões anteriores,
  cujas pontas dependem da taxa e dos prazos daquele dia) e a partir do
  **vencimento seguinte ao DDI base** (contrato + 1).
- Cenários: cada um move UM fator de risco (PRE, cupom IPCA, cupom USD, cupom EUR,
  dólar, euro) e mantém os demais. Aproximação = duration + convexidade nas taxas e
  delta no spot (o futuro de moeda é linear no spot).
"""
from __future__ import annotations

import datetime as dt
from dataclasses import dataclass

import numpy as np
import pandas as pd

from mesa import calendario as cal
from mesa import instrumentos as ins
from mesa.curvas.montagem import Curvas
from mesa.dados.leitura import DadosD0

FAMILIAS_MOEDA = ("DOL", "EUR")
LADOS = {"DI1": ("Tomado", "Dado"), "DAP": ("Tomado", "Dado"), "DDI": ("Tomado", "Dado"),
         "FRC": ("Tomado", "Dado"), "DOL": ("Compra", "Venda"), "EUR": ("Compra", "Venda")}
FATOR = {"DI1": "PRE", "DAP": "Cupom IPCA", "DDI": "Cupom USD", "FRC": "Cupom USD"}


def _paralelo(bps: float):
    return lambda du: bps


# (nome, fator, choque(du)): bps na taxa do fator (função do DU do contrato) ou % no spot.
# Steepening: −50 bps no curto subindo linearmente até +50 bps em 2.520 DU (10 anos); Flattening: o inverso.
CENARIOS = [
    ("PRE paralelo +100", "PRE", _paralelo(100.0)),
    ("PRE paralelo −100", "PRE", _paralelo(-100.0)),
    ("PRE steepening −50/+50", "PRE", lambda du: -50 + 100 * min(du, 2520) / 2520),
    ("PRE flattening +50/−50", "PRE", lambda du: 50 - 100 * min(du, 2520) / 2520),
    ("Cupom IPCA +100", "Cupom IPCA", _paralelo(100.0)),
    ("Cupom IPCA −100", "Cupom IPCA", _paralelo(-100.0)),
    ("Cupom USD +100", "Cupom USD", _paralelo(100.0)),
    ("Cupom USD −100", "Cupom USD", _paralelo(-100.0)),
    ("Cupom EUR +100", "Cupom EUR", _paralelo(100.0)),
    ("Cupom EUR −100", "Cupom EUR", _paralelo(-100.0)),
    ("Dólar +10%", "Spot USD", _paralelo(10.0)),
    ("Dólar −10%", "Spot USD", _paralelo(-10.0)),
    ("Euro +10%", "Spot EUR", _paralelo(10.0)),
    ("Euro −10%", "Spot EUR", _paralelo(-10.0)),
]


@dataclass
class Operacao:
    familia: str
    ticker: str
    lado: str
    quantidade: float
    operado_hoje: bool = False
    entrada: float | None = None     # taxa (% a.a.) para contratos de taxa; preço (pontos) para moedas


@dataclass
class Perna:
    """Posição elementar em um contrato com preço de ajuste (FRC é decomposto em pernas DDI)."""
    familia: str
    ticker: str
    sinal_pu: int        # +1 comprado em PU/preço, −1 vendido
    qtd: float
    du: int
    dc: int
    pa: float
    pa_abertura: float
    taxa: float
    valor_ponto: float
    po: float | None     # preço da operação (se operada hoje)


@dataclass
class ResultadoBook:
    posicoes: pd.DataFrame
    risco_fator: pd.DataFrame
    risco_vertice: pd.DataFrame
    cenarios: pd.DataFrame
    prt: float           # IPCA pro rata usado no valor do ponto do DAP
    prt_origem: str      # como o PRT foi obtido (exibido no Book)


# ---------------------------------------------------------------- IPCA pro rata (DAP)
def _br(x: float, casas: int) -> str:
    """Número no formato brasileiro (7.706,59) para textos exibidos na planilha."""
    return f"{x:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _dia15_ref(t: dt.date) -> dt.date:
    """Dia 15 que inicia o período do IPCA pro rata de `t` (o do mês, se t ≥ 15; senão o do mês anterior)."""
    if t.day >= 15:
        return dt.date(t.year, t.month, 15)
    anterior = t.replace(day=1) - dt.timedelta(days=1)
    return dt.date(anterior.year, anterior.month, 15)


def _du_apos(a: dt.date, b: dt.date) -> int:
    """Dias úteis em (a, b] — a pode não ser dia útil (o dia 15 cai em fim de semana)."""
    return cal.du(a + dt.timedelta(days=1), b + dt.timedelta(days=1))


def _fator_abertura(dados: DadosD0, familia: str) -> float:
    """Mediana de PA_abertura(t) / PA(t−1) da família: o fator de correção FC que a B3 aplicou."""
    a = dados.ajustes_d1.set_index("ticker")
    b = dados.ajustes.set_index("ticker")
    tk = [t for t in a.index.intersection(b.index) if t.startswith(familia)]
    r = (b.loc[tk, "pa_abertura"] / a.loc[tk, "pa"]).dropna()
    return float(np.median(r)) if len(r) else np.nan


def ipca_pro_rata(dados: DadosD0) -> tuple[float, str]:
    """PRT_t, o IPCA pro rata do dia (OC 034/2025-VPC, anexo 2, cláusula 2):

        PRT_t = IPCA_{t-1} × (1 + Proj)^(dud_t / du_m)

    - IPCA_{t-1}: número-índice do mês = vértice 1 da curva IAS da B3 (2 casas, constante até o dia 15).
    - g = (1 + Proj)^(1/du_m), o crescimento do PRT por dia útil, está embutido no PA de abertura:
      FC_DI1 = fator DI e FC_DAP = fator DI ÷ (PRT_t/PRT_{t−k}) ⇒ g^k = FC_DI1 / FC_DAP
      (exige os ajustes do pregão anterior — o arquivo "D1" do site).
    - dud_t = dias úteis do dia 15 de referência (exclusive) até t (inclusive).
    Sem o pregão anterior (ou se ele for de outro período mensal), usa IPCA_{t-1} — aproximação que
    erra ~0 no dia 15 e até ~0,4% (uma projeção mensal) antes do dia 15 seguinte.
    """
    t = dados.data_ref
    indice = float(dados.curvas_b3.IAS.iloc[0])
    ref15 = _dia15_ref(t)
    dud = _du_apos(ref15, t)
    if dud == 0:
        return indice, f"PRT = IPCA do mês = {_br(indice, 2)} (hoje é o dia 15, aniversário do índice)"
    if dados.ajustes_d1 is not None and dados.data_d1 is not None and _dia15_ref(dados.data_d1) == ref15:
        k = cal.du(dados.data_d1, t)
        g = (_fator_abertura(dados, "DI1") / _fator_abertura(dados, "DAP")) ** (1 / k) if 1 <= k <= 3 else np.nan
        if np.isfinite(g) and 0.999 < g < 1.002:
            prt = indice * g ** dud
            du_m = _du_apos(ref15, dt.date(ref15.year + ref15.month // 12, ref15.month % 12 + 1, 15))
            proj = (g ** du_m - 1) * 100
            return prt, (f"PRT = {_br(prt, 2)} = IPCA do mês {_br(indice, 2)} × {_br(g, 7)}^{dud} DU "
                         f"(projeção do IPCA implícita no PA de abertura do DAP: {_br(proj, 2)}% no mês)")
    return indice, f"PRT ≈ IPCA do mês = {_br(indice, 2)} (aproximação: sem os ajustes do pregão anterior)"


# ---------------------------------------------------------------- validação
def validar(op: Operacao, dados: DadosD0) -> str | None:
    if op.familia not in LADOS:
        return f"Família inválida: {op.familia!r}."
    if op.lado not in LADOS[op.familia]:
        return f"Lado inválido para {op.familia}: use {' ou '.join(LADOS[op.familia])}."
    if not op.ticker or not str(op.ticker).startswith(op.familia):
        return f"Ticker {op.ticker!r} não é da família {op.familia}."
    try:
        q = float(op.quantidade)
    except (TypeError, ValueError):
        return "Quantidade inválida."
    if not np.isfinite(q) or q <= 0 or q != int(q):
        return "Quantidade deve ser inteira e maior que zero."
    linha = dados.ajustes[dados.ajustes.ticker == op.ticker]
    if linha.empty:
        return f"{op.ticker} não está no arquivo de ajustes de {dados.data_ref:%d/%m/%Y} (vencido ou inexistente)."
    if int(linha.du.iloc[0]) <= 0:
        return f"{op.ticker} vence hoje."
    if op.familia == "FRC":
        if not op.operado_hoje:
            return "FRC: apenas operações do dia (marque 'Operado hoje' e informe a taxa)."
        base = _ddi_base(dados)
        if linha.vencimento.iloc[0] <= base.vencimento:
            return f"FRC disponível a partir do vencimento seguinte ao DDI base ({base.ticker})."
    if op.operado_hoje:
        if op.entrada is None:
            return "Operado hoje: informe a taxa/preço de entrada."
        if not np.isfinite(op.entrada):
            return "Taxa/preço de entrada inválido: digite só o número (ex.: 13,5)."
    return None


def tickers_frc(dados: DadosD0) -> list[str]:
    """FRC negociáveis: vencimentos após o DDI base (contrato + 1)."""
    base = _ddi_base(dados)
    frc = dados.ajustes[(dados.ajustes.familia == "FRC") & (dados.ajustes.vencimento > base.vencimento)]
    return frc.sort_values("du").ticker.tolist()


# ---------------------------------------------------------------- pernas
def _valor_ponto(familia: str, dados: DadosD0, prt: float) -> float:
    if familia == "DI1":
        return 1.0
    if familia == "DAP":
        return 0.00025 * prt
    if familia == "DDI":
        return 0.5 * dados.indicadores["ptax_d1"]
    return 50.0


def _linha(dados: DadosD0, ticker: str) -> pd.Series:
    return dados.ajustes.set_index("ticker").loc[ticker]


def _po(familia: str, entrada: float, du: int, dc: int) -> float:
    if familia in ("DI1", "DAP"):
        return ins.pu_exp(entrada, du)
    if familia == "DDI":
        return ins.pu_lin(entrada, dc)
    return entrada


def _perna(dados, ticker, sinal_pu, qtd, prt, po=None) -> Perna:
    x = _linha(dados, ticker)
    fam = ticker[:3]
    return Perna(fam, ticker, sinal_pu, qtd, int(x.du), int(x.dc), float(x.pa), float(x.pa_abertura),
                 float(x.taxa) if pd.notna(x.taxa) else np.nan, _valor_ponto(fam, dados, prt), po)


def _ddi_base(dados: DadosD0) -> pd.Series:
    ddi = dados.ajustes[(dados.ajustes.familia == "DDI") & (dados.ajustes.du > 2)]  # rola no penúltimo dia
    return ddi.sort_values("du").iloc[0]


def pernas(op: Operacao, dados: DadosD0, prt: float | None = None) -> list[Perna]:
    prt = ipca_pro_rata(dados)[0] if prt is None else prt
    q = float(op.quantidade)
    sinal_taxa = 1 if op.lado in ("Tomado", "Compra") else -1
    if op.familia in FAMILIAS_MOEDA:
        po = op.entrada if op.operado_hoje else None
        return [_perna(dados, op.ticker, sinal_taxa, q, prt, po)]
    if op.familia != "FRC":
        x = _linha(dados, op.ticker)
        po = _po(op.familia, op.entrada, int(x.du), int(x.dc)) if op.operado_hoje else None
        return [_perna(dados, op.ticker, -sinal_taxa, q, prt, po)]
    # FRC → pontas DDI
    frc = _linha(dados, op.ticker)
    longa_tk = "DDI" + op.ticker[3:]
    base = _ddi_base(dados)
    cfrc = op.entrada                                                  # FRC: sempre operação do dia
    n1, n2 = int(base.dc), int(frc.dc)
    q1 = round(q / (1 + cfrc * (n2 - n1) / 36000))
    c1 = float(base.taxa)                                              # taxa da ponta curta (= ajuste do dia)
    c2 = ((1 + c1 * n1 / 36000) * (1 + cfrc * (n2 - n1) / 36000) - 1) * 36000 / n2
    po_curta, po_longa = float(base.pa), ins.pu_lin(c2, n2)            # spec FRC (b): ponta curta pelo preço de ajuste
    return [_perna(dados, longa_tk, -sinal_taxa, q, prt, po_longa),
            _perna(dados, base.ticker, +sinal_taxa, q1, prt, po_curta)]


# ---------------------------------------------------------------- cálculo
def _pnl_dia(p: Perna) -> float:
    ref = p.po if p.po is not None else p.pa_abertura
    return (p.pa - ref) * p.valor_ponto * p.qtd * p.sinal_pu


def calcular(ops: list[Operacao], dados: DadosD0, curvas: Curvas) -> ResultadoBook:
    i = dados.indicadores
    prt, prt_origem = ipca_pro_rata(dados)
    linhas, fatores, vertices, cen = [], [], [], []
    for n, op in enumerate(ops, start=1):
        erro = validar(op, dados)
        base = {"linha": n, "familia": op.familia, "ticker": op.ticker, "lado": op.lado,
                "quantidade": op.quantidade}
        if erro:
            linhas.append({**base, "status": erro})
            continue
        base["quantidade"] = float(op.quantidade)
        ps = pernas(op, dados, prt)
        x = _linha(dados, op.ticker)
        reg = {**base, "status": "OK", "vencimento": x.vencimento, "du": int(x.du), "dc": int(x.dc),
               "preco_ajuste": float(x.pa) if pd.notna(x.pa) else np.nan,
               "taxa_ajuste": float(x.taxa) if pd.notna(x.taxa) else np.nan,
               "pnl_dia": sum(_pnl_dia(p) for p in ps),
               "exposicao": sum(p.pa * p.valor_ponto * p.qtd * p.sinal_pu for p in ps),
               "dv01": 0.0, "dv01_pre": 0.0, "dv01_cupom": 0.0, "delta_fx_1pct": 0.0,
               "duration_mod": np.nan, "convexidade": np.nan}
        for p in ps:
            mult = p.valor_ponto * p.qtd * p.sinal_pu
            if p.familia in ("DI1", "DAP", "DDI"):
                linear = p.familia == "DDI"
                s = ins.sens_lin(p.taxa, p.dc) if linear else ins.sens_exp(p.taxa, p.du)
                dv = s.dv01 * mult
                reg["dv01"] += dv
                fatores.append((FATOR[p.familia], dv))
                vertices.append((FATOR[p.familia], p.ticker, dv))
                if op.familia != "FRC":
                    reg["duration_mod"], reg["convexidade"] = s.duration_mod, s.convexidade
                for nome, fator, choque_de in CENARIOS:
                    if fator != FATOR[p.familia]:
                        continue
                    choque = float(choque_de(p.du))
                    taxa_choque = p.taxa + choque * ins.BP
                    novo = ins.pu_lin(taxa_choque, p.dc) if linear else ins.pu_exp(taxa_choque, p.du)
                    cen.append((nome, (novo - s.preco) * mult, s.variacao(choque) * mult))
            else:  # DOL / EUR
                v = curvas.avaliar([p.du], [p.dc]).iloc[0]
                if p.familia == "DOL":
                    spot, cupom, cup_curva, f_spot, f_cupom = i["ptax_d1"], v.DOL, curvas.dol, "Spot USD", "Cupom USD"
                else:
                    spot, cupom, cup_curva, f_spot, f_cupom = (i["euro_cupom_limpo"], v.EUC, curvas.doc, "Spot EUR",
                                                               "Cupom EUR")
                s = ins.sens_forward(spot, v.PRE, p.du, cupom, p.dc)
                d_pre, d_cup, d_fx = s.dv01_pre * mult, s.dv01_cupom * mult, s.delta_1pct * mult
                reg["dv01_pre"] += d_pre
                reg["dv01_cupom"] += d_cup
                reg["delta_fx_1pct"] += d_fx
                fatores += [("PRE", d_pre), (f_cupom, d_cup), (f_spot, d_fx)]
                for curva, fator, dv in ((curvas.pre, "PRE", d_pre), (cup_curva, f_cupom, d_cup)):
                    pesos = curva.pesos_vertices([p.du], [p.dc])[:, 0]
                    for j in np.nonzero(np.abs(pesos) > 1e-12)[0]:
                        vertices.append((fator, curva.origem[j], dv * pesos[j]))
                for nome, fator, choque_de in CENARIOS:
                    choque = float(choque_de(p.du))
                    if fator == "PRE":
                        novo = ins.forward(spot, v.PRE + choque * ins.BP, p.du, cupom, p.dc)
                        aprox = s.variacao(choque_pre_bp=choque)
                    elif fator == f_cupom:
                        novo = ins.forward(spot, v.PRE, p.du, cupom + choque * ins.BP, p.dc)
                        aprox = s.variacao(choque_cupom_bp=choque)
                    elif fator == f_spot:
                        novo = ins.forward(spot * (1 + choque / 100), v.PRE, p.du, cupom, p.dc)
                        aprox = s.variacao(choque_spot_pct=choque)
                    else:
                        continue
                    cen.append((nome, (novo - s.preco) * mult, aprox * mult))
        linhas.append(reg)

    posicoes = pd.DataFrame(linhas)
    # risco: DV01 (R$ para +1 bp) nos fatores de taxa; delta (R$ para +1%) nos fatores de spot
    risco_fator = (pd.DataFrame(fatores, columns=["fator", "risco"]).groupby("fator", sort=False).sum()
                   .reset_index())
    risco_vertice = (pd.DataFrame(vertices, columns=["fator", "vertice", "dv01"])
                     .groupby(["fator", "vertice"], sort=False).sum().reset_index())
    c = pd.DataFrame(cen, columns=["cenario", "reprecificacao", "aproximacao"])
    ordem = [nome for nome, *_ in CENARIOS if nome in set(c.cenario)]   # só fatores presentes no book
    cenarios = (c.groupby("cenario", sort=False)[["reprecificacao", "aproximacao"]].sum()
                .reindex(ordem).rename_axis("cenario").reset_index())
    cenarios["erro_aproximacao"] = cenarios.aproximacao - cenarios.reprecificacao
    return ResultadoBook(posicoes, risco_fator, risco_vertice, cenarios, prt, prt_origem)
