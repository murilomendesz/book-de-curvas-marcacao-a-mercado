"""Leitura e normalização dos 3 arquivos diários (Ajustes, Indicadores, Taxas Referenciais).

Regras:
- A data de referência é lida de dentro dos arquivos; os três precisam concordar.
- O formato é validado: se faltar coluna/indicador/curva, para com mensagem clara.
- Aceita o formato antigo (jul/2026) e o atual (nov/2026) do arquivo de ajustes.
- Vencimento, DU e DC são sempre recalculados pelo calendário (o arquivo atual
  não os traz, e o antigo tinha datas de EUR desatualizadas).
- Indicadores são lidos pelo código (os nomes vêm com acentos corrompidos).
- Opcional: ajustes do pregão anterior (`ARQ_AJUSTES_D1`, o arquivo "D1" do site). Servem para
  medir o fator de correção que a B3 aplicou no PA de abertura (IPCA pro rata do DAP, ver book.py).
"""
from __future__ import annotations

import datetime as dt
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from mesa import calendario as cal

ARQ_AJUSTES = "Ajustes_do_Pregao_B3_D0.xlsx"
ARQ_INDICADORES = "Painel_Indicadores_D0.xlsx"
ARQ_TAXAS = "Taxas_Referencias_B3_D0.xlsx"
ARQ_AJUSTES_D1 = "Ajustes_do_Pregao_B3_D1.xlsx"     # pregão anterior (opcional)

FAMILIAS = ("DI1", "DDI", "FRC", "DAP", "DOL", "EUR")

# formato antigo -> atual
_RENOMEAR_AJUSTES = {
    "PU_Atual": "PA_Fechamento",
    "PU_Anterior": "PA_Abertura",
    "Taxa_Atual": "Taxa_Fechamento",
    "Taxa_Anterior": "Taxa_Abertura",
}
_COLUNAS_AJUSTES = {
    "TradeDate": "data_ref",
    "Ticker": "ticker",
    "Familia": "familia",
    "PA_Fechamento": "pa",
    "PA_Abertura": "pa_abertura",
    "Taxa_Fechamento": "taxa",
    "Taxa_Abertura": "taxa_abertura",
    "AjusteContrato": "ajuste_contrato",
    "TradQty": "negocios",
    "FinInstrmQty": "contratos",
    "OpenInterest": "contratos_abertos",
}
INDICADORES = {
    "RTDI1": "cdi",
    "RTDOL-CL": "dolar_cupom_limpo",
    "TXEURO": "euro_cupom_limpo",
    "RTREU-PF": "eurusd",
    "PTAX_USD_VENDA_D0": "ptax_d0",
    "PTAX_USD_VENDA_D-1": "ptax_d1",
}
CURVAS_B3 = ("PRE", "DOL", "DOC", "DCL", "PTX", "DPL", "SDE", "EUC", "ECC", "EUR", "IAS")  # IAS = IPCA pro rata (B3)


class FormatoInvalido(Exception):
    """O arquivo não tem a estrutura esperada (mudou de formato?)."""


class DatasDivergentes(Exception):
    """Os arquivos não são todos do mesmo pregão."""


@dataclass(frozen=True)
class DadosD0:
    data_ref: dt.date
    ajustes: pd.DataFrame        # ticker, familia, vencimento, du, dc, pa, taxa, ...
    indicadores: dict[str, float]
    datas_indicadores: dict[str, dt.date]
    curvas_b3: pd.DataFrame      # du, dc + uma coluna por curva oficial B3 (gabarito)
    ajustes_d1: pd.DataFrame | None = None   # ajustes do pregão anterior, se disponíveis
    data_d1: dt.date | None = None


def _exige(df: pd.DataFrame, colunas, arquivo: str) -> None:
    faltando = [c for c in colunas if c not in df.columns]
    if faltando:
        raise FormatoInvalido(f"{arquivo}: coluna(s) não encontrada(s): {', '.join(faltando)}. "
                              f"Colunas presentes: {', '.join(map(str, df.columns))}")


def _data_unica(serie: pd.Series, arquivo: str) -> dt.date:
    datas = pd.to_datetime(serie.dropna()).dt.date.unique()
    if len(datas) != 1:
        raise FormatoInvalido(f"{arquivo}: esperava uma única data de referência, encontrei {sorted(datas)}.")
    return datas[0]


def ler_ajustes(caminho: Path) -> tuple[dt.date, pd.DataFrame]:
    df = pd.read_excel(caminho, sheet_name="CONSOLIDADO").rename(columns=_RENOMEAR_AJUSTES)
    _exige(df, _COLUNAS_AJUSTES, caminho.name)
    data_ref = _data_unica(df["TradeDate"], caminho.name)
    df = df[list(_COLUNAS_AJUSTES)].rename(columns=_COLUNAS_AJUSTES)
    df["ticker"] = df["ticker"].astype(str).str.strip()
    df["familia"] = df["familia"].astype(str).str.strip()
    df = df[df["familia"].isin(FAMILIAS)].copy()
    df["vencimento"] = [cal.vencimento(t) for t in df["ticker"]]
    df["du"] = [cal.du(data_ref, v) for v in df["vencimento"]]
    df["dc"] = [cal.dc(data_ref, v) for v in df["vencimento"]]
    df = df.drop(columns="data_ref").sort_values(["familia", "dc"]).reset_index(drop=True)
    return data_ref, df


def ler_indicadores(caminho: Path) -> tuple[dict[str, float], dict[str, dt.date]]:
    df = pd.read_excel(caminho)
    _exige(df, ["codigo_indicador", "data_referencia", "valor"], caminho.name)
    df["codigo_indicador"] = df["codigo_indicador"].astype(str).str.strip()
    df = df.set_index("codigo_indicador")
    faltando = [c for c in INDICADORES if c not in df.index]
    if faltando:
        raise FormatoInvalido(f"{caminho.name}: indicador(es) não encontrado(s): {', '.join(faltando)}.")
    valores = {nome: float(df.at[cod, "valor"]) for cod, nome in INDICADORES.items()}
    datas = {nome: pd.Timestamp(df.at[cod, "data_referencia"]).date() for cod, nome in INDICADORES.items()}
    return valores, datas


def ler_taxas(caminho: Path) -> tuple[dt.date, pd.DataFrame]:
    df = pd.read_excel(caminho, sheet_name="Consolidado")
    if "DCL" not in df.columns and "DCP" in df.columns:   # B3 renomeou DCL -> DCP após o Manual V21
        df = df.rename(columns={"DCP": "DCL"})
    _exige(df, ["data_geracao_dt", "dias_corridos_num", "dias_uteis_num", *CURVAS_B3], caminho.name)
    data_ref = _data_unica(df["data_geracao_dt"], caminho.name)
    df = df.rename(columns={"dias_uteis_num": "du", "dias_corridos_num": "dc"})
    return data_ref, df[["du", "dc", *CURVAS_B3]].reset_index(drop=True)


def ler_pasta(pasta: Path) -> DadosD0:
    pasta = Path(pasta)
    for nome in (ARQ_AJUSTES, ARQ_INDICADORES, ARQ_TAXAS):
        if not (pasta / nome).exists():
            raise FileNotFoundError(f"Arquivo não encontrado: {pasta / nome}")
    data_aj, ajustes = ler_ajustes(pasta / ARQ_AJUSTES)
    indicadores, datas_ind = ler_indicadores(pasta / ARQ_INDICADORES)
    data_tx, curvas = ler_taxas(pasta / ARQ_TAXAS)
    data_ind = datas_ind["cdi"]
    if not (data_aj == data_tx == data_ind):
        raise DatasDivergentes(
            f"Os arquivos não são do mesmo pregão: ajustes {data_aj:%d/%m/%Y}, "
            f"taxas {data_tx:%d/%m/%Y}, indicadores {data_ind:%d/%m/%Y}.")
    aj_d1, data_d1 = None, None
    if (pasta / ARQ_AJUSTES_D1).exists():
        data_d1, aj_d1 = ler_ajustes(pasta / ARQ_AJUSTES_D1)
        if data_d1 >= data_aj:          # D1 tem de ser um pregão anterior; senão é ignorado
            aj_d1, data_d1 = None, None
    return DadosD0(data_aj, ajustes, indicadores, datas_ind, curvas, aj_d1, data_d1)
