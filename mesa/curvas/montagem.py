"""Montagem das curvas a partir dos ajustes B3 de D0 (Manual de Curvas B3 V21).

Cada curva reproduz a metodologia do manual, inclusive os arredondamentos que a
B3 aplica em etapas intermediárias (validado vértice a vértice contra as curvas
oficiais — ver a seção "Validação" do README e a tabela do Painel).

| Curva | Item | Vértices                         | Interpolação   | Casas |
|-------|------|----------------------------------|----------------|-------|
| PRE   | 2.1  | CDI + DI1                        | FF252 (1.4.2)  | 3     |
| DOL   | 4.4  | v1 (CDI, Ptax t0/t-1) + DDI      | FF252 lin (1.4.3/1.4.10) | 3 |
| DOC   | 4.5  | v1 (CDI, Ptax t-1/DCL) + DDI limpo (2 casas) | 1.4.3 | 2 |
| DCL   | 4.6  | v1 (CDI, Ptax t0/DCL) + DDI limpo (2 casas)  | 1.4.3 | 2 |
| PTX   | 4.1  | PRE e DOL arredondadas, Ptax t-1 | paridade       | 7     |
| DPL   | 3.2  | v1 publicado pela B3 + DAP; FF252 sobre a inflação implícita | 1.4.2 | 2 |
| SDE   | 4.9  | vértices publicados pela B3      | 1.4.3 (convenção do projeto) | 2 |
| EUC   | 4.7  | DOC ⊗ SDE                        | —              | 4     |
| EUR   | 4.2  | PRE, EUC, TXEURO                 | paridade       | 7     |
| ECC   | —    | EUR fwd × spot Ptax t-1·EUR/USD (não documentada pela B3; só ≥ 15 DC) | — | 4 |

Forwards de cupom (FRC): calculados sobre `doc_cheio`, o cupom limpo SEM o arredondamento de 2
casas nos vértices DDI. Entre duas datas futuras o forward do limpo é igual ao do sujo (o fator
DolCL/Ptax t-1 se cancela), ou seja, sai dos DDIs com 3 casas — como a B3 faz o DDI a partir do FRC
(Apreçamento 1.4). Sobre o DOC arredondado o forward teria ruído de até 0,2 p.p.
"""
from __future__ import annotations

import datetime as dt
from dataclasses import dataclass

import numpy as np
import pandas as pd

from mesa import calendario as cal
from mesa.curvas.interpolacao import EXP252, LIN360, Curva
from mesa.dados.leitura import DadosD0

ECC_DC_MINIMO = 15


def _pivos(ajustes: pd.DataFrame, familia: str) -> pd.DataFrame:
    x = ajustes[(ajustes.familia == familia) & (ajustes.du > 1) & ajustes.taxa.notna()]
    return x.sort_values("du")


@dataclass
class Curvas:
    data_ref: dt.date
    indicadores: dict[str, float]
    pre: Curva
    dol: Curva
    doc: Curva
    dcl: Curva
    inflacao: Curva      # inflação implícita (PRE × DPL) nos vértices DAP — base da interpolação do DPL
    sde: Curva
    doc_cheio: Curva     # cupom limpo sem o arredondamento dos vértices: só para forwards (FRC)

    # ---------------- curvas derivadas ----------------
    def avaliar(self, du, dc) -> pd.DataFrame:
        """Todas as curvas nos prazos (du, dc), com os arredondamentos da B3."""
        du = np.atleast_1d(np.asarray(du, float))
        dc = np.atleast_1d(np.asarray(dc, float))
        i = self.indicadores
        pre = self.pre.taxa_em(du, dc)
        dol = self.dol.taxa_em(du, dc)
        doc = self.doc.taxa_em(du, dc)
        dcl = self.dcl.taxa_em(du, dc)
        sde = self.sde.taxa_em(du, dc)
        fator_pre = (1 + pre / 100) ** (du / 252)

        ptx = np.round(fator_pre / (1 + dol * dc / 36000) * i["ptax_d1"], 7)
        ptx[du == 1] = i["ptax_d0"]                                   # Manual 4.1: vértice 1 = Ptax t0

        infl = self.inflacao.taxa_em(du, dc, arredondar=False)
        dpl = np.round(((1 + pre / 100) / (1 + infl / 100) - 1) * 100, 2)

        euc = np.round(((1 + doc * dc / 36000) * (1 + sde * dc / 36000) - 1) * 36000 / dc, 4)
        eur = np.round(fator_pre / (1 + euc * dc / 36000) * i["euro_cupom_limpo"], 7)
        spot_eur = i["ptax_d1"] * i["eurusd"]
        ecc = np.round((fator_pre / (eur / spot_eur) - 1) * 36000 / dc, 4)
        ecc[dc < ECC_DC_MINIMO] = np.nan

        return pd.DataFrame({
            "du": du.astype(int), "dc": dc.astype(int),
            "PRE": pre, "DOL": dol, "DOC": doc, "DCL": dcl, "PTX": ptx,
            "DPL": dpl, "SDE": sde, "EUC": euc, "EUR": eur, "ECC": ecc,
            "tipo_PRE": self.pre.tipo_em(du),
            "tipo_USD": self.dol.tipo_em(du),
            "tipo_IPCA": self.inflacao.tipo_em(du),
        })


def montar(dados: DadosD0) -> Curvas:
    ref, i, aj, b3 = dados.data_ref, dados.indicadores, dados.ajustes, dados.curvas_b3
    cdi, pt0, pt1, dol_cl = i["cdi"], i["ptax_d0"], i["ptax_d1"], i["dolar_cupom_limpo"]
    du1, dc1 = 1, cal.dc(ref, cal.soma_du(ref, 1))
    f_cdi = (1 + cdi / 100) ** (du1 / 252)

    # PRE (2.1)
    di1 = _pivos(aj, "DI1")
    pre = Curva("PRE", EXP252, np.r_[du1, di1.du], np.r_[dc1, di1.dc], np.r_[cdi, di1.taxa],
                ["CDI", *di1.ticker], casas=3)

    # Cupom cambial (4.4, 4.5, 4.6)
    ddi = _pivos(aj, "DDI")
    du_c, dc_c = np.r_[du1, ddi.du], np.r_[dc1, ddi.dc]
    origem_c = ["CDI/Ptax", *ddi.ticker]
    v1_sujo = (f_cdi / (pt0 / pt1) - 1) * 36000 / dc1
    dol = Curva("DOL", LIN360, du_c, dc_c, np.r_[v1_sujo, ddi.taxa], origem_c, casas=3)
    limpo_cheio = ((1 + ddi.taxa.values * ddi.dc.values / 36000) * dol_cl / pt1 - 1) * 36000 / ddi.dc.values
    limpo = np.round(limpo_cheio, 2)                 # a B3 arredonda o limpo do vértice antes de interpolar (DOC/DCL)
    v1_doc = (f_cdi / (pt1 / dol_cl) - 1) * 36000 / dc1
    v1_dcl = (f_cdi / (pt0 / dol_cl) - 1) * 36000 / dc1
    doc = Curva("DOC", LIN360, du_c, dc_c, np.r_[v1_doc, limpo], origem_c, casas=2)
    dcl = Curva("DCL", LIN360, du_c, dc_c, np.r_[v1_dcl, limpo], origem_c, casas=2)
    doc_cheio = Curva("DOC (forwards)", LIN360, du_c, dc_c, np.r_[v1_doc, limpo_cheio], origem_c)

    # Cupom IPCA (3.2): FF252 sobre a inflação implícita PRE × DPL
    dap = _pivos(aj, "DAP")
    v1_dpl = float(b3.loc[b3.du == 1, "DPL"].iloc[0])           # vértice 1 publicado pela B3 (decisão do projeto)
    du_i, dc_i = np.r_[du1, dap.du], np.r_[dc1, dap.dc]
    cupom = np.r_[v1_dpl, dap.taxa]
    pre_piv = pre.taxa_em(du_i, dc_i)
    infl = ((1 + pre_piv / 100) / (1 + cupom / 100) - 1) * 100
    inflacao = Curva("Inflação implícita", EXP252, du_i, dc_i, infl, ["DPL B3 (v1)", *dap.ticker])

    # Spread EUR × USD (4.9): só vértices publicados
    sde = Curva("SDE", LIN360, b3.du, b3.dc, b3.SDE, casas=2)

    return Curvas(ref, i, pre, dol, doc, dcl, inflacao, sde, doc_cheio)
