"""Download dos arquivos diários (site IdeaFin, dados B3/BCB), sem login.

Como o site funciona (verificado em 16–18/11/2026): cada arquivo é um arquivo fixo no Google
Drive (um ID para o "D0", outro para o "D1" = pregão anterior...) cujo conteúdo é trocado a cada
pregão; os links curtos (s.gy) do site redirecionam para esses mesmos IDs.

Rota principal: download direto do Google Drive pelo ID.
Rota de reserva: link curto (s.gy) → redirect HTTP → ID → Drive. Usada quando o ID falha e também
quando o pregão baixado está atrasado (se o site passar a usar arquivos novos, o ID antigo
poderia continuar no ar com um pregão velho). O ID novo só é aceito se vier da cadeia de redirects
do link curto e apontar para o Google Drive.

Os arquivos são baixados para uma pasta temporária; só depois de lidos e validados (mesma data de
referência nos três) vão para `data/raw/AAAA-MM-DD/`. Se o site ainda não publicou o pregão do dia
(ex.: ajustes ainda em D-1), nada é alterado e o erro explica o que aconteceu.
"""
from __future__ import annotations

import datetime as dt
import re
import shutil
import tempfile
from pathlib import Path
from urllib.parse import urlparse

import requests

from mesa.dados import leitura

RAIZ = Path(__file__).resolve().parents[2]
PASTA_RAW = RAIZ / "data" / "raw"

FONTES = {   # arquivo: (ID no Google Drive, link curto de reserva) — arquivos "D0" do site
    leitura.ARQ_AJUSTES: ("1q8UC2Tdij9A8oLq_6_HFAB2qiY0Wd4sK", "https://d3te7j2.s.gy/yDzr9X"),
    leitura.ARQ_TAXAS: ("1VOZC5ERRvTAqmOUhHz78obN_zlukXp5r", "https://d3te7j2.s.gy/Xlz7xl"),
    leitura.ARQ_INDICADORES: ("14yDodU3C9fpDhRuqBNzFVnDqo1Zm0afO", "https://d3te7j2.s.gy/XUUl1Z"),
}
FONTE_D1 = {  # ajustes do pregão anterior ("D1" do site): fator de correção do PA de abertura (DAP)
    leitura.ARQ_AJUSTES_D1: ("1z6UH_XGo70Yqc-RSlqJHJrBNT16XucXG", "https://d3te7j2.s.gy/nEbBYc"),
}
URL_DRIVE = "https://drive.google.com/uc?export=download&id={}"
_ID_RE = re.compile(r"/d/([A-Za-z0-9_-]{20,})")
_HOSTS_DRIVE = {"drive.google.com", "docs.google.com", "drive.usercontent.google.com"}
TIMEOUT = 60


class FalhaDownload(Exception):
    """Não foi possível baixar um arquivo por nenhuma das rotas."""


class PregaoNaoPublicado(Exception):
    """Os arquivos disponíveis ainda não são todos do mesmo pregão."""


def _sessao() -> requests.Session:
    s = requests.Session()
    s.headers["User-Agent"] = "mesa-curvas-b3"
    return s


def _baixar_id(s: requests.Session, file_id: str) -> bytes | None:
    try:
        r = s.get(URL_DRIVE.format(file_id), timeout=TIMEOUT)
    except requests.RequestException:
        return None
    return r.content if r.ok and r.content[:2] == b"PK" else None   # .xlsx é um zip ("PK")


def _resolver_link_curto(s: requests.Session, url: str) -> str | None:
    """ID do Drive para onde o link curto redireciona: só da cadeia de redirects (nunca do corpo da
    página) e só se o destino for o Google Drive."""
    try:
        r = s.get(url, timeout=TIMEOUT, allow_redirects=True)
    except requests.RequestException:
        return None
    for u in [h.headers.get("Location", "") for h in getattr(r, "history", [])] + [r.url]:
        m = _ID_RE.search(u)
        if m and (urlparse(u).hostname or "") in _HOSTS_DRIVE:
            return m.group(1)
    return None


def baixar_arquivo(s: requests.Session, nome: str, avisos: list[str], fontes: dict | None = None) -> bytes:
    file_id, curto = (fontes or FONTES)[nome]
    conteudo = _baixar_id(s, file_id)
    if conteudo is None:
        novo = _resolver_link_curto(s, curto)
        if novo and novo != file_id:
            avisos.append(f"{nome}: ID do arquivo mudou ({file_id} → {novo}); atualize FONTES em download.py.")
        conteudo = _baixar_id(s, novo) if novo else None
    if conteudo is None:
        raise FalhaDownload(f"Não consegui baixar {nome} (Drive e link curto falharam). Verifique a internet ou o site.")
    return conteudo


def _ler(pasta: Path) -> leitura.DadosD0:
    try:
        return leitura.ler_pasta(pasta)
    except leitura.DatasDivergentes as e:
        raise PregaoNaoPublicado(f"O site ainda não publicou todos os arquivos do mesmo pregão. {e} "
                                 "Nada foi alterado; tente novamente mais tarde.") from e


def _baixar_d1(s: requests.Session, pasta: Path) -> None:
    """Ajustes do pregão anterior (opcional: sem ele o valor do ponto do DAP usa o IPCA do mês)."""
    try:
        (pasta / leitura.ARQ_AJUSTES_D1).write_bytes(baixar_arquivo(s, leitura.ARQ_AJUSTES_D1, [], FONTE_D1))
    except FalhaDownload:
        pass


def _ids_dos_links_curtos(s: requests.Session) -> dict[str, str]:
    """Arquivos cujo link curto aponta hoje para um ID diferente do configurado em FONTES."""
    novos = {nome: _resolver_link_curto(s, curto) for nome, (_, curto) in FONTES.items()}
    return {nome: i for nome, i in novos.items() if i and i != FONTES[nome][0]}


def baixar_d0(destino_raiz: Path = PASTA_RAW, sessao: requests.Session | None = None,
              esperado: dt.date | None = None) -> tuple[leitura.DadosD0, list[str]]:
    """Baixa, valida e grava os arquivos do último pregão. Retorna os dados lidos e avisos.

    `esperado`: pregão mais recente que pode existir agora. Se o baixado for mais antigo, confere
    se os links curtos passaram a apontar para outros arquivos e, se sim, usa-os quando forem mais novos.
    """
    s = sessao or _sessao()
    avisos: list[str] = []
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        pasta = tmp / "ids"
        pasta.mkdir()
        for nome in FONTES:
            (pasta / nome).write_bytes(baixar_arquivo(s, nome, avisos))
        _baixar_d1(s, pasta)
        dados = _ler(pasta)
        if esperado and dados.data_ref < esperado:
            mudaram = _ids_dos_links_curtos(s)
            if mudaram:
                alt = tmp / "links"
                shutil.copytree(pasta, alt)
                for nome, novo in mudaram.items():
                    conteudo = _baixar_id(s, novo)
                    if conteudo is not None:
                        (alt / nome).write_bytes(conteudo)
                try:
                    dados_alt = leitura.ler_pasta(alt)
                except (leitura.DatasDivergentes, leitura.FormatoInvalido):
                    dados_alt = None
                if dados_alt is not None and dados_alt.data_ref > dados.data_ref:
                    dados, pasta = dados_alt, alt
                    avisos.append("Os links do site passaram a apontar para outros arquivos: "
                                  + "; ".join(f"{n} → {i}" for n, i in mudaram.items())
                                  + ". Usei os novos; atualize FONTES em download.py.")
        destino = Path(destino_raiz) / f"{dados.data_ref:%Y-%m-%d}"
        destino.mkdir(parents=True, exist_ok=True)
        for arq in pasta.iterdir():
            shutil.copy2(arq, destino / arq.name)
    return dados, avisos
