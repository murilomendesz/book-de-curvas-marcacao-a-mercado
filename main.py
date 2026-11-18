"""Mesa de Curvas B3 — ponto de entrada.

Uso:
    python main.py              # roda o pregão do dia (baixa, monta curvas, atualiza a planilha)
    python main.py pregao       # idem
    python main.py book         # só recalcula o Book (sem baixar dados)

Opções:
    --arquivo CAMINHO   planilha .xlsm (padrão: Mesa_Curvas_B3.xlsm na pasta do projeto). Os botões passam
                        o FullName do Excel, que numa pasta do OneDrive é uma URL (https://d.docs.live.net/...)
    --pasta CAMINHO     usa uma pasta local com os 3 arquivos em vez de baixar (uso de desenvolvimento)

Cada execução grava um resumo em data/ultima_execucao.log (útil quando roda pelo botão, cuja janela fecha sozinha).
Só uma execução por vez: se o botão for clicado enquanto o terminal (ou outro clique) ainda está rodando,
a segunda execução é recusada sem tocar na planilha — duas escrevendo ao mesmo tempo quebravam uma à outra.
"""
from __future__ import annotations

import argparse
import contextlib
import datetime as dt
import os
import shutil
import sys
import tempfile
import traceback
from pathlib import Path

import xlwings as xw

from mesa import calendario as cal
from mesa.curvas import montagem
from mesa.dados import download, leitura
from mesa.excel import escrita
from mesa.excel.modelo import ARQ_MODELO

try:
    import msvcrt          # Windows (o projeto só roda no Windows: Excel via COM)
except ImportError:        # pragma: no cover
    msvcrt = None

RAIZ = Path(__file__).resolve().parent
ARQ_PADRAO = RAIZ / "Mesa_Curvas_B3.xlsm"
ARQ_LOG = RAIZ / "data" / "ultima_execucao.log"
# fora do OneDrive (o sincronizador não deve mexer num arquivo travado); um por usuário do Windows
ARQ_TRAVA = Path(os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()) / "MesaCurvasB3" / "execucao.lock"


class ExecucaoEmAndamento(Exception):
    """Já existe outra execução da Mesa rodando (terminal ou botão)."""


@contextlib.contextmanager
def _execucao_unica():
    """Trava de arquivo do Windows (msvcrt.locking): exclusiva entre processos e liberada sozinha se o
    processo morrer — não sobra trava presa depois de um erro ou de fechar o terminal."""
    if msvcrt is None:
        yield
        return
    ARQ_TRAVA.parent.mkdir(parents=True, exist_ok=True)
    f = open(ARQ_TRAVA, "a+")
    try:
        f.seek(0)
        msvcrt.locking(f.fileno(), msvcrt.LK_NBLCK, 1)
    except OSError:
        f.close()
        raise ExecucaoEmAndamento("Já existe uma execução da Mesa em andamento (terminal ou botão). "
                                  "Aguarde ela terminar (~40 s) e tente de novo.") from None
    try:
        yield
    finally:
        f.seek(0)
        msvcrt.locking(f.fileno(), msvcrt.LK_UNLCK, 1)
        f.close()


@contextlib.contextmanager
def _excel_ocupado(app: xw.App):
    """Enquanto escreve: sem redesenhar a tela e sem aceitar teclado/mouse (um clique no meio — no botão,
    numa célula — fazia o Excel recusar ou embaralhar as chamadas do Python). Restaura sempre no fim."""
    app.screen_updating = False
    app.interactive = False
    try:
        yield
    finally:
        app.interactive = True
        app.screen_updating = True


def _log(texto: str, novo: bool = False) -> None:
    """Mostra no console e grava no log da última execução (a janela do botão fecha sozinha)."""
    print(texto)
    try:
        ARQ_LOG.parent.mkdir(parents=True, exist_ok=True)
        with open(ARQ_LOG, "w" if novo else "a", encoding="utf-8") as f:
            f.write(f"{dt.datetime.now():%d/%m/%Y %H:%M:%S}  {texto}\n")
    except OSError:
        pass


def _eh_url(alvo: str) -> bool:
    return "://" in alvo


def corresponde(alvo: str, fullname_excel: str, fullname_local: str | None) -> bool:
    """A planilha aberta é a pedida em `alvo`?

    `alvo` vem do terminal (caminho local) ou dos botões (`ThisWorkbook.FullName`). Numa pasta
    sincronizada pelo OneDrive, o Excel devolve no FullName uma URL (https://d.docs.live.net/...) em
    vez do caminho; o xlwings (`Book.fullname`) converte essa URL para o caminho local. Por isso:
    URL → compara com o FullName do Excel; caminho → compara com o caminho local do xlwings.
    """
    if _eh_url(alvo):
        return fullname_excel.casefold() == alvo.casefold()
    if fullname_local is None:
        return False
    normaliza = lambda p: os.path.normcase(os.path.abspath(p))   # Windows não diferencia maiúsculas
    return normaliza(fullname_local) == normaliza(alvo)


def _procura_aberta(alvo: str) -> xw.Book | None:
    for app in xw.apps:
        for wb in app.books:
            try:
                bruto = wb.api.FullName
            except Exception:
                continue
            try:
                local = wb.fullname
            except Exception:          # ex.: SharePoint sem caminho local configurado no xlwings
                local = None
            if corresponde(alvo, bruto, local):
                return wb
    return None


def abrir_planilha(alvo: str | Path) -> xw.Book:
    alvo = str(alvo)
    wb = _procura_aberta(alvo)
    if wb is not None:
        return wb
    if _eh_url(alvo):
        raise FileNotFoundError(f"A planilha {alvo} não está aberta no Excel.")
    caminho = Path(alvo)
    if not caminho.exists():
        shutil.copy(ARQ_MODELO, caminho)
    app = xw.apps.active if xw.apps.count else xw.App(visible=True, add_book=False)
    app.visible = True
    return app.books.open(str(caminho))


def _status(wb: xw.Book, texto: str, erro: bool = False) -> None:
    try:
        c = wb.sheets["Painel"].range("C4")
        c.value = texto
        c.api.Font.Color = escrita.E.rgb(escrita.E.VERMELHO if erro else escrita.E.VERDE)
    except Exception:
        pass


def rodar_pregao(wb: xw.Book, pasta: Path | None) -> None:
    agora = dt.datetime.now()
    esperado = cal.ultimo_pregao_esperado(agora.date())
    if pasta:
        dados, avisos = leitura.ler_pasta(pasta), []
    else:
        dados, avisos = download.baixar_d0(esperado=esperado)
    curvas = montagem.montar(dados)
    with _excel_ocupado(wb.app):
        entrada = escrita.ler_input_book(wb)
        escrita.escrever(wb, dados, curvas, agora, sys.executable, str(Path(__file__).resolve()))
        escrita.escrever_book(wb, dados, curvas, entrada, agora)
        wb.sheets["Painel"].activate()
    if not pasta and dados.data_ref < esperado:
        aviso = (f"⚠ O site ainda não publicou o pregão de {esperado:%d/%m/%Y} (sai após ~22h). "
                 f"Exibindo o último disponível: {dados.data_ref:%d/%m/%Y}.")
        avisos.append(aviso)
        _status(wb, aviso, erro=True)
        wb.sheets["Config"].range("C7").value = aviso
    for a in avisos:
        _log(f"AVISO: {a}")
    wb.save()
    _log(f"Pregão {dados.data_ref:%d/%m/%Y} processado.")


def rodar_book(wb: xw.Book) -> None:
    ref = wb.sheets["Config"].range("C6").value
    if isinstance(ref, (dt.date, dt.datetime)):
        ref = ref.strftime("%Y-%m-%d")
    pasta = download.PASTA_RAW / str(ref)
    if not ref or not pasta.exists():
        raise FileNotFoundError("Nenhum pregão carregado na planilha. Clique em 'Rodar pregão do dia' primeiro.")
    dados = leitura.ler_pasta(pasta)
    curvas = montagem.montar(dados)
    with _excel_ocupado(wb.app):
        entrada = escrita.ler_input_book(wb)
        escrita.escrever_book(wb, dados, curvas, entrada, dt.datetime.now())
    base = wb.sheets["Config"].range("C7").value or ""
    _status(wb, f"{base}  ·  book recalculado em {dt.datetime.now():%H:%M}")
    wb.save()
    _log("Book recalculado.")


def main(argv=None) -> int:
    for fluxo in (sys.stdout, sys.stderr):        # console do Windows (cp1252) não quebra com ✔/→
        try:
            fluxo.reconfigure(errors="replace")
        except (AttributeError, ValueError):
            pass
    p = argparse.ArgumentParser(description="Mesa de Curvas B3")
    p.add_argument("acao", nargs="?", default="pregao", choices=["pregao", "book"])
    p.add_argument("--arquivo", default=str(ARQ_PADRAO))   # str, não Path: Path estragaria uma URL do OneDrive
    p.add_argument("--pasta", type=Path, default=None)
    a = p.parse_args(argv)
    comando = f"python main.py {a.acao} --arquivo {a.arquivo}" + (f" --pasta {a.pasta}" if a.pasta else "")
    try:
        with _execucao_unica():
            return _executar(a, comando)
    except ExecucaoEmAndamento as e:     # não toca na planilha nem apaga o log da execução que está rodando
        _log(f"RECUSADO ({comando}): {e}")
        return 3


def _executar(a: argparse.Namespace, comando: str) -> int:
    _log(comando, novo=True)
    wb = None
    try:
        wb = abrir_planilha(a.arquivo)
        _status(wb, f"Executando ({a.acao})... aguarde ~40 s sem mexer na planilha.")
        if a.acao == "pregao":
            rodar_pregao(wb, a.pasta)
        else:
            rodar_book(wb)
        return 0
    except (download.PregaoNaoPublicado, download.FalhaDownload, leitura.FormatoInvalido,
            leitura.DatasDivergentes, FileNotFoundError) as e:
        if wb is not None:
            _status(wb, f"⚠ {e}", erro=True)
        _log(f"ERRO: {e}")
        return 1
    except Exception as e:
        if wb is not None:
            _status(wb, f"⚠ Erro inesperado: {e}", erro=True)
        _log(f"ERRO INESPERADO:\n{traceback.format_exc()}")
        return 2


if __name__ == "__main__":
    sys.exit(main())
