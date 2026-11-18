"""Paleta e formatação (padrão visual do PROJETO DI / Amostra do Projeto.xlsx)."""
from __future__ import annotations

FUNDO = "0F1B33"       # fundo das abas
CARTAO = "1E2D4F"      # KPIs / caixas
CABECALHO = "243356"   # cabeçalho de tabela
LINHA = "162240"       # linhas de tabela
BRANCO = "FFFFFF"
CINZA = "7A8BA8"       # rótulos
CINZA_CLARO = "B8C5D9" # valores secundários
AZUL = "4A90D9"        # spot
AZUL_CLARO = "5BA3E6"
LARANJA = "F39C12"     # forward / destaque
VERDE = "2ECC71"
VERMELHO = "E74C3C"
INPUT = "2B3F66"       # células editáveis
FONTE = "Arial"


def rgb(hex6: str) -> int:
    """Cor no formato da API do Excel (BGR em inteiro)."""
    r, g, b = int(hex6[0:2], 16), int(hex6[2:4], 16), int(hex6[4:6], 16)
    return r + g * 256 + b * 65536


def pinta(rng, fundo: str | None = None, cor: str = BRANCO, negrito: bool = False, tamanho: float = 10,
          alinhar: str | None = None, formato: str | None = None) -> None:
    """Aplica fundo/fonte/alinhamento/formato numérico a um range xlwings."""
    if fundo:
        rng.color = "#" + fundo
    f = rng.api.Font
    f.Name, f.Size, f.Bold, f.Color = FONTE, tamanho, negrito, rgb(cor)
    if alinhar:
        rng.api.HorizontalAlignment = {"esq": -4131, "centro": -4108, "dir": -4152}[alinhar]
    if formato:
        aplica_formato(rng, formato)


_INTL: dict[int, tuple[str, str, str]] = {}


def formato_local(fmt: str, decimal: str, milhar: str, ano: str) -> str:
    """Converte um formato escrito em convenção inglesa ("#,##0.00", "dd/mm/yyyy") para a do Excel
    do usuário (ex.: pt-BR "#.##0,00", "dd/mm/aaaa"). Texto entre aspas é preservado."""
    out, aspas = [], False
    for ch in fmt:
        if ch == '"':
            aspas = not aspas
            out.append(ch)
        elif aspas:
            out.append(ch)
        elif ch == ".":
            out.append(decimal)
        elif ch == ",":
            out.append(milhar)
        elif ch in "yY":
            out.append(ano)
        else:
            out.append(ch)
    return "".join(out)


def aplica_formato(rng, fmt: str) -> None:
    """Via COM, o Excel interpreta o formato no idioma local; traduz antes de aplicar."""
    rng.api.NumberFormatLocal = formato_local(fmt, *convencoes(rng.sheet.book.app))


def convencoes(app) -> tuple[str, str, str]:
    """(decimal, milhar, código de ano) do Excel em uso — cache por processo do Excel."""
    if app.pid not in _INTL:
        i = app.api.International          # tupla: (3) decimal, (4) milhar, (19) código de ano
        _INTL[app.pid] = (i[2], i[3], i[18])
    return _INTL[app.pid]
