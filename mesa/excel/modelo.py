"""Gera o modelo .xlsm com as macros dos botões (executado uma vez pelo desenvolvedor).

O modelo fica versionado em `modelo/Mesa_Curvas_B3_modelo.xlsm`; o usuário nunca
precisa rodar isto (exigiria habilitar "Confiar no acesso ao modelo de objeto do
projeto VBA"). O `main.py` copia o modelo e preenche as abas.

As macros só chamam o Python gravado na aba Config pelo `main.py`
(`sys.executable`), garantindo que o botão use o mesmo Python/bibliotecas.
"""
from __future__ import annotations

from pathlib import Path

import xlwings as xw

RAIZ = Path(__file__).resolve().parents[2]
ARQ_MODELO = RAIZ / "modelo" / "Mesa_Curvas_B3_modelo.xlsm"

VBA = r'''
Option Explicit

Public Sub RodarPregaoDoDia()
    ExecutarPython "pregao"
End Sub

Public Sub RecalcularBook()
    ExecutarPython "book"
End Sub

Private Sub ExecutarPython(acao As String)
    Dim py As String, script As String, cmd As String
    py = ThisWorkbook.Worksheets("Config").Range("C4").Value
    script = ThisWorkbook.Worksheets("Config").Range("C5").Value
    If Len(Dir(py)) = 0 Or Len(Dir(script)) = 0 Then
        MsgBox "Python ou main.py não encontrados (aba Config)." & vbCrLf & _
               "Rode uma vez no terminal: python main.py", vbExclamation, "Mesa de Curvas B3"
        Exit Sub
    End If
    ThisWorkbook.Worksheets("Painel").Range("C4").Value = "Executando (" & acao & ")... aguarde."
    cmd = """" & py & """ """ & script & """ " & acao & " --arquivo """ & ThisWorkbook.FullName & """"
    Shell cmd, vbMinimizedNoFocus
End Sub
'''


def gerar(destino: Path = ARQ_MODELO) -> Path:
    destino.parent.mkdir(parents=True, exist_ok=True)
    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.add()
        modulo = wb.api.VBProject.VBComponents.Add(1)   # 1 = módulo padrão
        modulo.Name = "MesaCurvas"
        modulo.CodeModule.AddFromString(VBA)
        wb.sheets[0].name = "Painel"
        wb.api.SaveAs(str(destino), FileFormat=52)      # 52 = .xlsm
        wb.close()
    finally:
        app.quit()
    return destino


if __name__ == "__main__":
    print(gerar())
