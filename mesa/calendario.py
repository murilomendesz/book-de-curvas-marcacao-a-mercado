"""Calendário de dias úteis e vencimentos dos contratos B3.

Dias úteis: feriados nacionais ANBIMA (`data/calendario/feriados_anbima.csv`).
Contagem de DU: dias úteis entre a data de referência (exclusive) e o vencimento
(inclusive) — mesma contagem dos ajustes da B3 (validada em 207 contratos).

Vencimentos (especificações dos contratos B3 — ofícios circulares 034/2025 e 015/2025, Futuro de Euro):
- DI1, DDI, FRC, DOL: primeiro dia útil do mês de vencimento.
- DAP: dia 15 do mês; se não for dia útil, o dia útil seguinte.
- EUR: fixing = 2º dia útil em Chicago e Nova York imediatamente anterior à 3ª
  quarta-feira do mês; vencimento = dia de sessão B3 imediatamente posterior ao
  fixing ou, se o fixing não for dia de sessão B3, o 2º dia de sessão posterior
  (especificação vigente do Futuro de Euro, cláusula 1).
"""
from __future__ import annotations

import datetime as dt
from functools import lru_cache
from pathlib import Path

import numpy as np
import pandas as pd
from pandas.tseries.holiday import USFederalHolidayCalendar

ARQ_FERIADOS = Path(__file__).resolve().parent.parent / "data" / "calendario" / "feriados_anbima.csv"

CODIGO_MES = {c: i for i, c in enumerate("FGHJKMNQUVXZ", start=1)}
FAMILIAS_1O_DU = {"DI1", "DDI", "FRC", "DOL"}

Data = dt.date | np.datetime64 | pd.Timestamp | str


def _d(x: Data) -> np.datetime64:
    return np.datetime64(pd.Timestamp(x).date(), "D")


def _date(x: np.datetime64) -> dt.date:
    return pd.Timestamp(x).date()


@lru_cache(maxsize=1)
def feriados() -> np.ndarray:
    """Feriados nacionais ANBIMA (2001–2099) como datetime64[D]."""
    df = pd.read_csv(ARQ_FERIADOS, parse_dates=["data"])
    return df["data"].values.astype("datetime64[D]")


@lru_cache(maxsize=1)
def _feriados_eua() -> np.ndarray:
    h = USFederalHolidayCalendar().holidays("2001-01-01", "2099-12-31")
    return np.asarray(h.values, dtype="datetime64[D]")


def _checa_intervalo(*datas: np.datetime64) -> None:
    for d in datas:
        if not (np.datetime64("2001-01-01") <= d <= np.datetime64("2099-12-31")):
            raise ValueError(f"Data {d} fora do calendário ANBIMA (2001–2099).")


def ajusta_dia_util(d: Data) -> dt.date:
    """A própria data se for dia útil; senão, o próximo dia útil."""
    x = _d(d)
    _checa_intervalo(x)
    return _date(np.busday_offset(x, 0, roll="forward", holidays=feriados()))


def soma_du(d: Data, n: int) -> dt.date:
    """Data `n` dias úteis após `d` (`d` precisa ser dia útil)."""
    return _date(np.busday_offset(_d(d), n, roll="raise", holidays=feriados()))


def du(inicio: Data, fim: Data) -> int:
    """Dias úteis entre `inicio` (exclusive) e `fim` (inclusive)."""
    a, b = _d(inicio), _d(fim)
    _checa_intervalo(a, b)
    # busday_count conta [a, b); para a e b dias úteis equivale a (a, b].
    return int(np.busday_count(a, b, holidays=feriados()))


def dc(inicio: Data, fim: Data) -> int:
    """Dias corridos entre `inicio` e `fim`."""
    return int((_d(fim) - _d(inicio)).astype(int))


def _mes_ano(ticker: str) -> tuple[int, int]:
    codigo, ano = ticker[-3], ticker[-2:]
    if codigo not in CODIGO_MES or not ano.isdigit():
        raise ValueError(f"Ticker inválido: {ticker!r} (esperado FAMILIA + letra do mês + AA).")
    return CODIGO_MES[codigo], 2000 + int(ano)


def fixing_eur(mes: int, ano: int) -> dt.date:
    """2º dia útil EUA estritamente antes da 3ª quarta-feira do mês.

    roll="forward": se a própria 3ª quarta for feriado nos EUA (Juneteenth, 19/06), conta-se a
    partir do dia útil seguinte — os 2 dias úteis anteriores à quarta continuam sendo os de antes
    dela (com roll="backward" o fixing recuaria um dia a mais).
    """
    primeiro = dt.date(ano, mes, 1)
    terceira_quarta = primeiro + dt.timedelta(days=(2 - primeiro.weekday()) % 7 + 14)
    return _date(np.busday_offset(_d(terceira_quarta), -2, roll="forward", holidays=_feriados_eua()))


def _vencimento_eur(mes: int, ano: int) -> dt.date:
    fixing = _d(fixing_eur(mes, ano))
    if np.is_busday(fixing, holidays=feriados()):
        venc = np.busday_offset(fixing, 1, holidays=feriados())
    else:   # rola para a 1ª sessão após o fixing e anda mais uma → 2ª sessão posterior ao fixing
        venc = np.busday_offset(fixing, 1, roll="forward", holidays=feriados())
    return _date(venc)


def vencimento(ticker: str) -> dt.date:
    """Data de vencimento a partir do ticker (ex.: 'DI1F27', 'DAPK35', 'EURZ26')."""
    familia = ticker[:-3]
    mes, ano = _mes_ano(ticker)
    if familia in FAMILIAS_1O_DU:
        return ajusta_dia_util(dt.date(ano, mes, 1))
    if familia == "DAP":
        return ajusta_dia_util(dt.date(ano, mes, 15))
    if familia == "EUR":
        return _vencimento_eur(mes, ano)
    raise ValueError(f"Família sem regra de vencimento: {familia!r} ({ticker}).")


def grade_du(data_ref: Data, max_dc: int) -> pd.DataFrame:
    """Todos os dias úteis após `data_ref` até `max_dc` dias corridos: colunas data, du, dc."""
    ref = _d(data_ref)
    fim = ref + np.timedelta64(max_dc, "D")
    _checa_intervalo(ref, fim)
    dias = np.arange(ref + np.timedelta64(1, "D"), fim + np.timedelta64(1, "D"), dtype="datetime64[D]")
    dias = dias[np.is_busday(dias, holidays=feriados())]
    return pd.DataFrame({
        "data": pd.to_datetime(dias).date,
        "du": np.arange(1, len(dias) + 1),
        "dc": (dias - ref).astype(int),
    })


def ultimo_pregao_esperado(hoje: Data) -> dt.date:
    """Pregão mais recente que pode existir em `hoje`: o próprio dia, se útil; senão o dia útil anterior."""
    x = _d(hoje)
    return _date(np.busday_offset(x, 0, roll="backward", holidays=feriados()))
