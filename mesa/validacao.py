"""Validação diária das curvas montadas × curvas oficiais B3 e paridade DOL teórico × ajuste."""
from __future__ import annotations

import numpy as np
import pandas as pd

from mesa import instrumentos as ins
from mesa.curvas.montagem import Curvas
from mesa.dados.leitura import DadosD0

TOLERANCIA = {
    "PRE": 0.0005, "DOL": 0.0005, "DOC": 0.005, "DCL": 0.005, "PTX": 5e-8,
    "DPL": 0.01, "EUC": 5e-5, "EUR": 5e-8, "ECC": 0.02,
}
NOMES = {
    "PRE": "PRE (DI × Pré)", "DOL": "Cupom sujo USD", "DOC": "Cupom limpo USD (DOC)",
    "DCL": "Cupom limpo USD (DCL)", "PTX": "Dólar forward", "DPL": "Cupom IPCA",
    "EUC": "Cupom limpo EUR", "EUR": "Euro forward", "ECC": "Cupom sujo EUR (≥ 15 DC)",
}


def contra_b3(dados: DadosD0, curvas: Curvas) -> pd.DataFrame:
    b3 = dados.curvas_b3
    nosso = curvas.avaliar(b3.du, b3.dc)
    linhas = []
    for c, tol in TOLERANCIA.items():
        ok = nosso[c].notna().values & b3[c].notna().values
        dif = np.abs(nosso[c].values[ok] - b3[c].values[ok])
        linhas.append({"curva": NOMES[c], "vertices": int(ok.sum()),
                       "dentro": int((dif <= tol + 1e-12).sum()), "max_dif": float(dif.max()), "tolerancia": tol})
    return pd.DataFrame(linhas)


def paridade_dol(dados: DadosD0, curvas: Curvas) -> pd.DataFrame:
    dol = dados.ajustes[(dados.ajustes.familia == "DOL") & (dados.ajustes.du > 1)].sort_values("du")
    v = curvas.avaliar(dol.du, dol.dc)
    teorico = ins.forward(dados.indicadores["ptax_d1"], v.PRE.values, dol.du.values, v.DOL.values, dol.dc.values)
    return pd.DataFrame({"ticker": dol.ticker.values, "vencimento": dol.vencimento.values, "du": dol.du.values,
                         "ajuste": dol.pa.values, "teorico": teorico, "dif": teorico - dol.pa.values})
