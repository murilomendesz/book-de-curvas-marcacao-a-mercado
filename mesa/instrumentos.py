"""Preço e sensibilidades dos futuros B3.

Preço (especificações B3 / Manual de Apreçamento V64):
- DI1, DAP: PU = 100.000 / (1 + i)^(DU/252)                     (taxa exponencial)
- DDI:      PU = 100.000 / (1 + i·DC/360)                       (taxa linear)
- DOL, EUR: F  = spot × (1 + PRE)^(DU/252) / (1 + cupom·DC/360)  (paridade; pontos por 1.000 da moeda)

Valor do ponto (R$): DI1 = 1 · DAP = 0,00025 × IPCA pro rata · DDI = US$ 0,50 × Ptax D-1
· DOL = EUR = 50.

Sensibilidades (algébricas — derivadas das fórmulas de preço acima; validadas por
bump & reprice nos testes). Convenção: bump de +1 bp (0,01 p.p.) na taxa, na
convenção de cada curva; resultados em pontos de preço por contrato.
"""
from __future__ import annotations

from dataclasses import dataclass

BP = 0.01        # 1 bp em pontos percentuais
PU_FACE = 100_000.0


# ---------------- juros: PU a partir da taxa ----------------
def pu_exp(taxa, du):
    return PU_FACE / (1 + taxa / 100) ** (du / 252)


def taxa_de_pu_exp(pu, du):
    return ((PU_FACE / pu) ** (252 / du) - 1) * 100


def pu_lin(taxa, dc):
    return PU_FACE / (1 + taxa * dc / 36000)


def taxa_de_pu_lin(pu, dc):
    return (PU_FACE / pu - 1) * 36000 / dc


@dataclass(frozen=True)
class Sensibilidade:
    preco: float          # PU (juros) ou F (moedas), em pontos
    dv01: float           # Δpreço para +1 bp na taxa (pontos, 1ª ordem)
    duration_mod: float   # −(1/P)·dP/di, com i em decimal  (anos)
    convexidade: float    # (1/P)·d²P/di², com i em decimal

    def variacao(self, choque_bp: float) -> float:
        """Δpreço aproximado (duration + convexidade) para um choque em bp."""
        di = choque_bp * BP / 100
        return self.preco * (-self.duration_mod * di + 0.5 * self.convexidade * di ** 2)


def sens_exp(taxa, du) -> Sensibilidade:
    """PU exponencial base 252 (DI1, DAP; também o fator PRE das moedas)."""
    p, t, r = pu_exp(taxa, du), du / 252, taxa / 100
    dmod = t / (1 + r)
    conv = t * (t + 1) / (1 + r) ** 2
    return Sensibilidade(p, -p * dmod * BP / 100, dmod, conv)


def sens_lin(taxa, dc) -> Sensibilidade:
    """PU linear base 360 (DDI; também o fator cupom das moedas)."""
    p, tau, c = pu_lin(taxa, dc), dc / 360, taxa / 100
    dmod = tau / (1 + c * tau)
    conv = 2 * tau ** 2 / (1 + c * tau) ** 2
    return Sensibilidade(p, -p * dmod * BP / 100, dmod, conv)


# ---------------- moedas: forward por paridade ----------------
def forward(spot, pre, du, cupom, dc):
    """Preço futuro em pontos (R$ por 1.000 unidades da moeda)."""
    return spot * 1000 * (1 + pre / 100) ** (du / 252) / (1 + cupom * dc / 36000)


@dataclass(frozen=True)
class SensibilidadeFx:
    preco: float
    dv01_pre: float      # ΔF para +1 bp na PRE
    dv01_cupom: float    # ΔF para +1 bp no cupom
    delta_1pct: float    # ΔF para +1% no spot (F é linear no spot: exato)
    conv_pre: float      # (1/F)·∂²F/∂pre², pre em decimal
    conv_cupom: float    # (1/F)·∂²F/∂cupom², cupom em decimal

    def variacao(self, choque_pre_bp: float = 0.0, choque_cupom_bp: float = 0.0, choque_spot_pct: float = 0.0) -> float:
        """ΔF aproximado: 1ª + 2ª ordem em cada taxa (um fator por vez nos cenários) e linear no spot."""
        dp, dc = choque_pre_bp * BP / 100, choque_cupom_bp * BP / 100
        return (self.dv01_pre * choque_pre_bp + self.dv01_cupom * choque_cupom_bp
                + 0.5 * self.preco * (self.conv_pre * dp ** 2 + self.conv_cupom * dc ** 2)
                + self.delta_1pct * choque_spot_pct)


def sens_forward(spot, pre, du, cupom, dc) -> SensibilidadeFx:
    f = forward(spot, pre, du, cupom, dc)
    t, p = du / 252, pre / 100
    tau, c = dc / 360, cupom / 100
    d_pre = f * t / (1 + p) * BP / 100
    d_cup = -f * (dc / 36000) / (1 + cupom * dc / 36000) * BP
    conv_pre = t * (t - 1) / (1 + p) ** 2           # ∂²[(1+p)^t]/∂p² ÷ (1+p)^t
    conv_cup = 2 * tau ** 2 / (1 + c * tau) ** 2     # ∂²[1/(1+cτ)]/∂c² ÷ [1/(1+cτ)]
    return SensibilidadeFx(f, d_pre, d_cup, f * 0.01, conv_pre, conv_cup)
