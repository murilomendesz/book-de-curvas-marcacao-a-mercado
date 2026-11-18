"""Tabelas de saída (pandas puro): grades diárias, vértices e KPIs de cada aba.

Forward: sob flat forward a taxa a termo é constante entre dois vértices. A coluna
de forward da grade traz, para cada dia, o forward do intervalo de vértices que o
contém (degrau). É calculado com as taxas de mercado dos próprios vértices, sem o
arredondamento B3 da curva spot (que geraria ruído artificial no forward):
- PRE: taxas DI1 (FRA de DI); IPCA: taxas DAP;
- cupom limpo USD/EUR: `doc_cheio` (limpo sem arredondar = forward dos DDIs, ou seja, a FRC).
"""
from __future__ import annotations

import numpy as np
import pandas as pd

from mesa import calendario as cal
from mesa.curvas.interpolacao import EXP252, LIN360, ln_fator
from mesa.curvas.montagem import Curvas
from mesa.dados.leitura import DadosD0

MAX_DC = 3000


def forward_degrau(p_du, p_dc, p_lnf, du, convencao: str) -> np.ndarray:
    """Forward do intervalo de vértices que contém cada prazo (vértice 0 = hoje, fator 1)."""
    p_du, p_dc, p_lnf = (np.r_[0.0, np.asarray(x, float)] for x in (p_du, p_dc, p_lnf))
    du = np.asarray(du, float)
    b = np.clip(np.searchsorted(p_du, du, side="left"), 1, len(p_du) - 1)
    a = b - 1
    dl = p_lnf[b] - p_lnf[a]
    if convencao == EXP252:
        return np.expm1(dl * 252 / (p_du[b] - p_du[a])) * 100
    return np.expm1(dl) * 36000 / (p_dc[b] - p_dc[a])


def _taxa_pivos(c: Curvas, chave: str, du, dc) -> np.ndarray:
    """Taxa de `chave` nos pivôs da sua curva, para o cálculo de forwards (ver docstring do módulo)."""
    du, dc = np.asarray(du, float), np.asarray(dc, float)
    if chave == "PRE":
        return c.pre.taxa_em(du, dc, False)
    if chave == "DPL":        # com a PRE arredondada, como na montagem: no vértice DAP dá a própria taxa do DAP
        pre, infl = c.pre.taxa_em(du, dc), c.inflacao.taxa_em(du, dc, False)
        return ((1 + pre / 100) / (1 + infl / 100) - 1) * 100
    doc = c.doc_cheio.taxa_em(du, dc, False)
    if chave == "DOC":
        return doc
    sde = c.sde.taxa_em(du, dc, False)
    return ((1 + doc * dc / 36000) * (1 + sde * dc / 36000) - 1) * 36000 / dc   # EUC = DOC ⊗ SDE


def _fwd(c: Curvas, chave: str, piv_du, piv_dc, du) -> np.ndarray:
    conv = EXP252 if chave in ("PRE", "DPL") else LIN360
    taxa = _taxa_pivos(c, chave, piv_du, piv_dc)
    return forward_degrau(piv_du, piv_dc, ln_fator(taxa, piv_du, piv_dc, conv), du, conv)


def grades(dados: DadosD0, c: Curvas, max_dc: int = MAX_DC) -> dict[str, pd.DataFrame]:
    g = cal.grade_du(dados.data_ref, max_dc)
    v = c.avaliar(g.du, g.dc)
    du = g.du.values
    base = g[["data", "du", "dc"]]
    f_pre = (1 + v.PRE / 100) ** (du / 252)
    f_dpl = (1 + v.DPL / 100) ** (du / 252)
    pre = base.assign(spot=v.PRE, fator=f_pre, pu=100000 / f_pre,
                      fwd=_fwd(c, "PRE", c.pre.du, c.pre.dc, du), tipo=v.tipo_PRE)
    ipca = base.assign(spot=v.DPL, pre=v.PRE, inflacao=((1 + v.PRE / 100) / (1 + v.DPL / 100) - 1) * 100,
                       pu=100000 / f_dpl, fwd=_fwd(c, "DPL", c.inflacao.du, c.inflacao.dc, du), tipo=v.tipo_IPCA)
    usd = base.assign(sujo=v.DOL, limpo=v.DOC, fwd=_fwd(c, "DOC", c.doc.du, c.doc.dc, du),
                      dolar_fwd=v.PTX, tipo=v.tipo_USD)
    eur = base.assign(limpo=v.EUC, sujo=v.ECC, fwd=_fwd(c, "EUC", c.doc.du, c.doc.dc, du),
                      euro_fwd=v.EUR, sde=v.SDE, tipo=v.tipo_USD)
    return {"PRE": pre, "IPCA": ipca, "USD": usd, "EUR": eur}


def vertices(dados: DadosD0, c: Curvas) -> dict[str, pd.DataFrame]:
    """Vértices de cada curva (ajustes B3) com o forward entre vencimentos consecutivos."""
    aj = dados.ajustes.set_index("ticker")
    venc1 = cal.soma_du(dados.data_ref, 1)
    out = {}
    for aba, curva, chave in (("PRE", c.pre, "PRE"), ("IPCA", c.inflacao, "DPL"), ("USD", c.doc, "DOC"),
                              ("EUR", c.doc, "EUC")):
        taxa = _taxa_pivos(c, chave, curva.du, curva.dc)
        conv = EXP252 if chave in ("PRE", "DPL") else LIN360
        fwd = forward_degrau(curva.du, curva.dc, ln_fator(taxa, curva.du, curva.dc, conv), curva.du, conv)
        nomes = curva.origem if aba != "EUR" else ["v1", *[o.replace("DDI", "DDI⊗SDE ") for o in curva.origem[1:]]]
        venc = [venc1 if k == 0 else aj.loc[o, "vencimento"] for k, o in enumerate(curva.origem)]
        out[aba] = pd.DataFrame({"vertice": nomes, "vencimento": venc, "du": curva.du.astype(int),
                                 "dc": curva.dc.astype(int), "taxa": taxa, "fwd": fwd})
    return out


def ajustes_moeda(dados: DadosD0, familia: str) -> pd.DataFrame:
    x = dados.ajustes[(dados.ajustes.familia == familia) & (dados.ajustes.du > 0)].sort_values("du")
    return x[["ticker", "vencimento", "du", "dc", "pa", "contratos_abertos"]].reset_index(drop=True)


def kpis(dados: DadosD0, grade: dict[str, pd.DataFrame]) -> dict[str, float]:
    def em(aba, col, du):
        g = grade[aba]
        return float(g.loc[(g.du - du).abs().idxmin(), col])
    pre1, pre5 = em("PRE", "spot", 252), em("PRE", "spot", 1260)
    return {
        "CDI": dados.indicadores["cdi"], "PRE 1 ano": pre1, "PRE 5 anos": pre5, "Inclinação 1a–5a (bps)": (pre5 - pre1) * 100,
        "Cupom IPCA 2 anos": em("IPCA", "spot", 504), "Inflação implícita 2 anos": em("IPCA", "inflacao", 504),
        "Cupom limpo USD 1 ano": em("USD", "limpo", 252), "Dólar forward 1 ano": em("USD", "dolar_fwd", 252),
        "Cupom EUR 1 ano": em("EUR", "limpo", 252), "Euro forward 1 ano": em("EUR", "euro_fwd", 252),
        "Ptax D-1": dados.indicadores["ptax_d1"], "Dólar cupom limpo (spot)": dados.indicadores["dolar_cupom_limpo"],
    }
