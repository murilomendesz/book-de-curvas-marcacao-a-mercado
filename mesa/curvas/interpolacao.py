"""Interpolação Flat Forward do Manual de Curvas B3 (V21).

- 1.4.2 Flat Forward 252 (taxas exponenciais, base DU/252): ln(fator) linear em DU.
- 1.4.3 Flat Forward 252 com convenção linear (taxas lineares, base DC/360):
  ln(1 + i·DC/36000) linear em DU.
- Após o último vértice: mesma reta usando o penúltimo e o último vértice
  (1.4.6 para exponencial, 1.4.10 para linear) — marcado como "Extrapolado".
"""
from __future__ import annotations

from dataclasses import dataclass, field

import numpy as np

EXP252 = "exp252"
LIN360 = "lin360"

ORIGINAL, INTERPOLADO, EXTRAPOLADO = "Original", "Interpolado", "Extrapolado"


def ln_fator(taxa, du, dc, convencao: str):
    taxa, du, dc = (np.asarray(x, float) for x in (taxa, du, dc))
    if convencao == EXP252:
        return np.log1p(taxa / 100) * du / 252
    if convencao == LIN360:
        return np.log1p(taxa * dc / 36000)
    raise ValueError(convencao)


def taxa_de_ln_fator(lnf, du, dc, convencao: str):
    lnf, du, dc = (np.asarray(x, float) for x in (lnf, du, dc))
    if convencao == EXP252:
        return np.expm1(lnf * 252 / du) * 100
    if convencao == LIN360:
        return np.expm1(lnf) * 36000 / dc
    raise ValueError(convencao)


@dataclass
class Curva:
    """Curva definida por vértices (pivôs) e interpolada por flat forward em DU."""
    nome: str
    convencao: str
    du: np.ndarray
    dc: np.ndarray
    taxa: np.ndarray
    origem: list[str] = field(default_factory=list)   # ex.: "CDI", "DI1F27"
    casas: int | None = None                          # arredondamento B3 da curva publicada

    def __post_init__(self):
        ordem = np.argsort(self.du, kind="stable")
        self.du = np.asarray(self.du, float)[ordem]
        self.dc = np.asarray(self.dc, float)[ordem]
        self.taxa = np.asarray(self.taxa, float)[ordem]
        if self.origem:
            self.origem = [self.origem[i] for i in ordem]
        if np.any(np.diff(self.du) <= 0):
            raise ValueError(f"{self.nome}: vértices com DU repetido ou fora de ordem: {self.du}")
        if len(self.du) < 2:
            raise ValueError(f"{self.nome}: são necessários ao menos 2 vértices.")
        self._lnf = ln_fator(self.taxa, self.du, self.dc, self.convencao)

    def lnf_em(self, du):
        """ln do fator acumulado no prazo `du` (flat forward em DU)."""
        du = np.atleast_1d(np.asarray(du, float))
        x, y = self.du, self._lnf
        out = np.interp(du, x, y)
        fora = du > x[-1]
        out[fora] = y[-2] + (y[-1] - y[-2]) * (du[fora] - x[-2]) / (x[-1] - x[-2])
        antes = du < x[0]
        if antes.any():
            raise ValueError(f"{self.nome}: prazo antes do primeiro vértice ({x[0]:.0f} DU).")
        return out

    def taxa_em(self, du, dc, arredondar: bool = True):
        du, dc = np.atleast_1d(np.asarray(du, float)), np.atleast_1d(np.asarray(dc, float))
        t = taxa_de_ln_fator(self.lnf_em(du), du, dc, self.convencao)
        return np.round(t, self.casas) if (arredondar and self.casas is not None) else t

    def pesos_vertices(self, du, dc) -> np.ndarray:
        """∂taxa(du)/∂taxa(vértice j) para cada vértice j (bump & reprice de 1 bp).

        Usado para distribuir o DV01 de um prazo qualquer entre os vértices da curva
        (DV01 por vértice / key-rate). Com flat forward, só os dois vértices vizinhos
        (ou os dois últimos, na extrapolação) têm peso não nulo.
        """
        du, dc = np.atleast_1d(np.asarray(du, float)), np.atleast_1d(np.asarray(dc, float))
        base = self.taxa_em(du, dc, arredondar=False)
        pesos = np.zeros((len(self.du), len(du)))
        for j in range(len(self.du)):
            taxa = self.taxa.copy()
            taxa[j] += 0.01
            bump = Curva(self.nome, self.convencao, self.du, self.dc, taxa)
            pesos[j] = (bump.taxa_em(du, dc, arredondar=False) - base) / 0.01
        return pesos

    def tipo_em(self, du):
        du = np.atleast_1d(np.asarray(du, float))
        tipo = np.where(du > self.du[-1], EXTRAPOLADO, INTERPOLADO).astype(object)
        tipo[np.isin(du, self.du)] = ORIGINAL
        return tipo
