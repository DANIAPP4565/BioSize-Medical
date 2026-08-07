from __future__ import annotations

import io
import json
import math
import os
import re
import shutil
import subprocess
import tempfile
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Callable

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
import yaml
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from pypdf import PdfReader
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle, PageBreak
from scipy.stats import norm
from sqlalchemy import create_engine, inspect


APP_NAME = "BioSize Clinical"
APP_VERSION = "1.1.0"
MIN_POWER = 0.80


# -----------------------------------------------------------------------------
# Modelos de datos y utilidades generales
# -----------------------------------------------------------------------------
@dataclass
class SampleSizeResult:
    design_code: str
    design_label: str
    n_raw_total: float
    n_final_total: int
    groups_raw: dict[str, float] = field(default_factory=dict)
    groups_final: dict[str, int] = field(default_factory=dict)
    formula: str = ""
    method: str = ""
    assumptions: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    extra: dict[str, Any] = field(default_factory=dict)
    uses_power: bool = True


DESIGNS: dict[str, str] = {
    "mean_estimation": "Descriptivo · Estimación de una media",
    "proportion_estimation": "Descriptivo · Estimación de una proporción",
    "means_independent": "RCT/Comparativo · Dos medias independientes",
    "means_paired": "RCT/Comparativo · Dos medias apareadas",
    "proportions_independent": "RCT/Comparativo · Dos proporciones independientes",
    "mcnemar": "RCT/Comparativo · Proporciones apareadas (McNemar)",
    "superiority_continuous": "Superioridad · Variable continua",
    "superiority_binary": "Superioridad · Variable dicotómica",
    "noninferiority_continuous": "No inferioridad · Variable continua",
    "noninferiority_binary": "No inferioridad · Variable dicotómica",
    "equivalence_continuous": "Equivalencia · Variable continua (TOST)",
    "equivalence_binary": "Equivalencia · Variable dicotómica (TOST)",
    "odds_ratio": "Asociación · Odds Ratio",
    "risk_ratio": "Asociación · Riesgo Relativo",
    "diagnostic_accuracy": "Validación diagnóstica · Sensibilidad y especificidad",
    "survival": "Supervivencia · Log-rank / Mantel–Cox",
}


ALIASES = {
    "media": "mean_estimation",
    "estimacion_media": "mean_estimation",
    "estimacion_de_media": "mean_estimation",
    "proporcion": "proportion_estimation",
    "estimacion_proporcion": "proportion_estimation",
    "medias_independientes": "means_independent",
    "medias_pareadas": "means_paired",
    "proporciones_independientes": "proportions_independent",
    "mcnemar": "mcnemar",
    "superioridad_continua": "superiority_continuous",
    "superioridad_dicotomica": "superiority_binary",
    "no_inferioridad_continua": "noninferiority_continuous",
    "no_inferioridad_dicotomica": "noninferiority_binary",
    "equivalencia_continua": "equivalence_continuous",
    "equivalencia_dicotomica": "equivalence_binary",
    "odds_ratio": "odds_ratio",
    "or": "odds_ratio",
    "riesgo_relativo": "risk_ratio",
    "rr": "risk_ratio",
    "validacion": "diagnostic_accuracy",
    "validacion_diagnostica": "diagnostic_accuracy",
    "diagnostico": "diagnostic_accuracy",
    "supervivencia": "survival",
    "mantel_cox": "survival",
    "log_rank": "survival",
}


def normalize_token(value: Any) -> str:
    text = str(value or "").strip().lower()
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]+", "_", text).strip("_")


def clamp_probability(value: float, name: str) -> float:
    value = float(value)
    if not 0 < value < 1:
        raise ValueError(f"{name} debe estar estrictamente entre 0 y 1.")
    return value


def positive(value: float, name: str) -> float:
    value = float(value)
    if value <= 0:
        raise ValueError(f"{name} debe ser mayor que cero.")
    return value


def z_alpha(alpha: float, sided: str) -> float:
    alpha = clamp_probability(alpha, "Error alfa")
    return float(norm.ppf(1 - alpha / 2)) if sided == "Bilateral" else float(norm.ppf(1 - alpha))


def z_beta(power: float) -> float:
    power = clamp_probability(power, "Poder estadístico")
    return float(norm.ppf(power))


def finite_population_correction(n: float, population: int | None) -> float:
    """Corrección exacta para muestreo aleatorio simple sin reemplazo."""
    if not population:
        return n
    if population <= 1:
        raise ValueError("La población finita debe ser mayor que 1.")
    return (population * n) / (population + n - 1)


def inflate_loss(n: float, loss_rate: float) -> int:
    loss_rate = float(loss_rate)
    if not 0 <= loss_rate < 0.95:
        raise ValueError("Las pérdidas esperadas deben estar entre 0% y menos de 95%.")
    return int(math.ceil(n / (1 - loss_rate)))


def continuity_correct_equal_groups(n_uncorrected: float, absolute_difference: float) -> float:
    """Aproximación de Casagrande–Pike–Smith/Fleiss para continuidad."""
    if absolute_difference <= 0:
        raise ValueError("La diferencia entre proporciones debe ser distinta de cero.")
    return (n_uncorrected / 4.0) * (
        1.0 + math.sqrt(1.0 + 2.0 / (n_uncorrected * absolute_difference))
    ) ** 2


def two_proportion_n_per_group(
    p1: float,
    p2: float,
    alpha: float,
    power: float,
    sided: str,
    yates: bool = False,
    null_difference: float = 0.0,
) -> float:
    """Aproximación normal con varianza agrupada bajo H0 y no agrupada bajo H1."""
    p1 = clamp_probability(p1, "Proporción 1")
    p2 = clamp_probability(p2, "Proporción 2")
    za = z_alpha(alpha, sided)
    zb = z_beta(power)
    observed_difference = p1 - p2
    effective_difference = abs(observed_difference - null_difference)
    if effective_difference <= 0:
        raise ValueError("La diferencia efectiva respecto de la hipótesis nula debe ser mayor que cero.")

    # Aproximación estable: promedio de las proporciones esperadas para la varianza nula.
    pbar = (p1 + p2) / 2
    numerator = (
        za * math.sqrt(2 * pbar * (1 - pbar))
        + zb * math.sqrt(p1 * (1 - p1) + p2 * (1 - p2))
    ) ** 2
    n = numerator / (effective_difference**2)
    if yates:
        n = continuity_correct_equal_groups(n, effective_difference)
    return n


def finalize_groups(
    design_code: str,
    design_label: str,
    raw_groups: dict[str, float],
    loss_rate: float,
    formula: str,
    method: str,
    assumptions: list[str],
    warnings: list[str] | None = None,
    extra: dict[str, Any] | None = None,
    uses_power: bool = True,
) -> SampleSizeResult:
    final_groups = {name: inflate_loss(n, loss_rate) for name, n in raw_groups.items()}
    return SampleSizeResult(
        design_code=design_code,
        design_label=design_label,
        n_raw_total=float(sum(raw_groups.values())),
        n_final_total=int(sum(final_groups.values())),
        groups_raw=raw_groups,
        groups_final=final_groups,
        formula=formula,
        method=method,
        assumptions=assumptions,
        warnings=warnings or [],
        extra=extra or {},
        uses_power=uses_power,
    )


# -----------------------------------------------------------------------------
# Motor estadístico
# -----------------------------------------------------------------------------
def calculate_sample_size(params: dict[str, Any], power_override: float | None = None) -> SampleSizeResult:
    code = params["design_code"]
    label = DESIGNS[code]
    alpha = float(params.get("alpha", 0.05))
    power = float(power_override if power_override is not None else params.get("power", 0.80))
    sided = params.get("sided", "Bilateral")
    loss = float(params.get("loss_rate", 0.0))
    za = z_alpha(alpha, sided)
    zb = z_beta(power)
    assumptions = [
        f"α={alpha:.3f}",
        f"contraste {sided.lower()}",
        f"pérdidas esperadas={loss:.1%}",
    ]
    warnings: list[str] = []

    if power < MIN_POWER and code not in {"mean_estimation", "proportion_estimation", "diagnostic_accuracy"}:
        warnings.append("El poder seleccionado es inferior al mínimo convencional de 80%.")

    if code == "mean_estimation":
        sd = positive(params["sd"], "Desviación estándar")
        precision = positive(params["precision"], "Precisión")
        n0 = (za * sd / precision) ** 2
        population = params.get("population") if params.get("finite_population") else None
        n = finite_population_correction(n0, population)
        final = inflate_loss(n, loss)
        assumptions += [f"DE={sd:g}", f"error máximo={precision:g}"]
        if population:
            assumptions.append(f"población finita N={population}")
        return SampleSizeResult(
            code,
            label,
            n,
            final,
            {"Muestra": n},
            {"Muestra": final},
            "n₀=(Z·σ/d)²; nFPC=N·n₀/(N+n₀−1); nfinal=ceil(nFPC/(1−R))",
            "Estimación de una media con precisión absoluta y corrección opcional por población finita.",
            assumptions,
            warnings,
            uses_power=False,
        )

    if code == "proportion_estimation":
        p = clamp_probability(params["p"], "Proporción esperada")
        precision = positive(params["precision"], "Precisión")
        n0 = (za**2) * p * (1 - p) / (precision**2)
        population = params.get("population") if params.get("finite_population") else None
        n = finite_population_correction(n0, population)
        final = inflate_loss(n, loss)
        assumptions += [f"p={p:.3f}", f"error máximo={precision:.3f}"]
        if population:
            assumptions.append(f"población finita N={population}")
        return SampleSizeResult(
            code,
            label,
            n,
            final,
            {"Muestra": n},
            {"Muestra": final},
            "n₀=Z²·p·(1−p)/d²; nFPC=N·n₀/(N+n₀−1); nfinal=ceil(nFPC/(1−R))",
            "Estimación de una proporción con precisión absoluta y corrección opcional por población finita.",
            assumptions,
            warnings,
            uses_power=False,
        )

    if code in {"means_independent", "superiority_continuous"}:
        delta = positive(abs(params["delta"]), "Diferencia clínicamente importante")
        sd = positive(params["sd"], "Desviación estándar")
        n = 2 * (sd**2) * ((za + zb) ** 2) / (delta**2)
        assumptions += [f"DMCI={delta:g}", f"DE común={sd:g}", f"poder={power:.1%}"]
        return finalize_groups(
            code,
            label,
            {"Grupo intervención": n, "Grupo control": n},
            loss,
            "n/grupo=2·σ²·(Zα+Zβ)²/Δ²",
            "Comparación bilateral/unilateral de dos medias independientes con asignación 1:1.",
            assumptions,
            warnings,
        )

    if code == "means_paired":
        delta = positive(abs(params["delta"]), "Diferencia media")
        sd_diff = positive(params["sd_diff"], "DE de las diferencias")
        n = ((za + zb) * sd_diff / delta) ** 2
        final = inflate_loss(n, loss)
        assumptions += [f"diferencia media={delta:g}", f"DE diferencias={sd_diff:g}", f"poder={power:.1%}"]
        return SampleSizeResult(
            code,
            label,
            n,
            final,
            {"Pares/sujetos": n},
            {"Pares/sujetos": final},
            "n=(Zα+Zβ)²·σd²/Δ²",
            "Comparación de una diferencia media dentro de los mismos sujetos.",
            assumptions,
            warnings,
        )

    if code in {"proportions_independent", "superiority_binary"}:
        p1 = params["p1"]
        p2 = params["p2"]
        yates = bool(params.get("yates", False))
        n = two_proportion_n_per_group(p1, p2, alpha, power, sided, yates=yates)
        assumptions += [f"p1={p1:.3f}", f"p2={p2:.3f}", f"poder={power:.1%}"]
        if yates:
            assumptions.append("corrección de continuidad de Yates/Fleiss")
        return finalize_groups(
            code,
            label,
            {"Grupo intervención/expuesto": n, "Grupo control/no expuesto": n},
            loss,
            "n/grupo=[Zα√(2p̄(1−p̄))+Zβ√(p1(1−p1)+p2(1−p2))]²/(p1−p2)²",
            "Comparación de dos proporciones independientes mediante aproximación normal.",
            assumptions,
            warnings,
        )

    if code == "mcnemar":
        p01 = clamp_probability(params["p01"], "Proporción discordante 0→1")
        p10 = clamp_probability(params["p10"], "Proporción discordante 1→0")
        q = p01 + p10
        if q >= 1:
            raise ValueError("La suma de las proporciones discordantes debe ser menor que 1.")
        diff = abs(p01 - p10)
        if diff <= 0:
            raise ValueError("Las proporciones discordantes deben diferir para calcular potencia.")
        n = ((za * math.sqrt(q) + zb * math.sqrt(q - diff**2)) ** 2) / (diff**2)
        if params.get("yates", False):
            n = continuity_correct_equal_groups(n, diff)
            assumptions.append("corrección de continuidad de Yates/Fleiss")
        final = inflate_loss(n, loss)
        assumptions += [f"p01={p01:.3f}", f"p10={p10:.3f}", f"poder={power:.1%}"]
        return SampleSizeResult(
            code,
            label,
            n,
            final,
            {"Pares/sujetos": n},
            {"Pares/sujetos": final},
            "n=[Zα√(p01+p10)+Zβ√(p01+p10−(p01−p10)²)]²/(p01−p10)²",
            "Prueba de McNemar basada en las proporciones discordantes esperadas.",
            assumptions,
            warnings,
        )

    if code == "noninferiority_continuous":
        sd = positive(params["sd"], "Desviación estándar")
        margin = positive(params["margin"], "Margen de no inferioridad")
        expected_diff = float(params.get("expected_diff", 0.0))
        distance = margin + expected_diff  # H0: T-C <= -margin
        if distance <= 0:
            raise ValueError("La diferencia esperada debe estar por encima del límite de no inferioridad (−margen).")
        za1 = float(norm.ppf(1 - alpha))
        n = 2 * sd**2 * (za1 + zb) ** 2 / distance**2
        assumptions += [f"margen={margin:g}", f"diferencia esperada T−C={expected_diff:g}", f"DE={sd:g}"]
        return finalize_groups(
            code,
            label,
            {"Tratamiento": n, "Control": n},
            loss,
            "n/grupo=2·σ²·(Z1−α+Z1−β)²/(Δesperada+M)²",
            "No inferioridad para diferencia de medias, contraste unilateral y asignación 1:1.",
            assumptions,
            warnings,
        )

    if code == "equivalence_continuous":
        sd = positive(params["sd"], "Desviación estándar")
        margin = positive(params["margin"], "Margen de equivalencia")
        expected_diff = abs(float(params.get("expected_diff", 0.0)))
        distance = margin - expected_diff
        if distance <= 0:
            raise ValueError("La diferencia esperada debe quedar estrictamente dentro de ±margen.")
        za1 = float(norm.ppf(1 - alpha))
        n = 2 * sd**2 * (za1 + zb) ** 2 / distance**2
        assumptions += [f"margen simétrico=±{margin:g}", f"|diferencia esperada|={expected_diff:g}", f"DE={sd:g}"]
        return finalize_groups(
            code,
            label,
            {"Tratamiento": n, "Control": n},
            loss,
            "n/grupo=2·σ²·(Z1−α+Z1−β)²/(M−|Δesperada|)²",
            "Equivalencia mediante dos pruebas unilaterales (TOST), asignación 1:1.",
            assumptions,
            warnings,
        )

    if code == "noninferiority_binary":
        p_t = clamp_probability(params["p1"], "Proporción tratamiento")
        p_c = clamp_probability(params["p2"], "Proporción control")
        margin = positive(params["margin"], "Margen de no inferioridad")
        n = two_proportion_n_per_group(
            p_t, p_c, alpha, power, "Unilateral", yates=bool(params.get("yates", False)), null_difference=-margin
        )
        assumptions += [f"pT={p_t:.3f}", f"pC={p_c:.3f}", f"margen absoluto={margin:.3f}"]
        warnings.append("Aproximación normal. En protocolos regulatorios se recomienda confirmar con Farrington–Manning o simulación.")
        return finalize_groups(
            code,
            label,
            {"Tratamiento": n, "Control": n},
            loss,
            "n/grupo≈[Zα√V0+Zβ√V1]²/[(pT−pC)−(−M)]²",
            "No inferioridad de dos proporciones con margen absoluto y contraste unilateral.",
            assumptions,
            warnings,
        )

    if code == "equivalence_binary":
        p_t = clamp_probability(params["p1"], "Proporción tratamiento")
        p_c = clamp_probability(params["p2"], "Proporción control")
        margin = positive(params["margin"], "Margen de equivalencia")
        expected_diff = abs(p_t - p_c)
        distance = margin - expected_diff
        if distance <= 0:
            raise ValueError("La diferencia esperada debe quedar estrictamente dentro de ±margen.")
        pbar = (p_t + p_c) / 2
        za1 = float(norm.ppf(1 - alpha))
        n = (
            za1 * math.sqrt(2 * pbar * (1 - pbar))
            + zb * math.sqrt(p_t * (1 - p_t) + p_c * (1 - p_c))
        ) ** 2 / distance**2
        if params.get("yates", False):
            n = continuity_correct_equal_groups(n, distance)
        assumptions += [f"pT={p_t:.3f}", f"pC={p_c:.3f}", f"margen=±{margin:.3f}"]
        warnings.append("Aproximación TOST normal. En protocolos regulatorios se recomienda confirmar con un método binomial dedicado.")
        return finalize_groups(
            code,
            label,
            {"Tratamiento": n, "Control": n},
            loss,
            "n/grupo≈[Z1−α√V0+Z1−β√V1]²/(M−|pT−pC|)²",
            "Equivalencia de dos proporciones mediante TOST aproximado.",
            assumptions,
            warnings,
        )

    if code == "odds_ratio":
        p_control = clamp_probability(params["p_control"], "Proporción de exposición/evento en controles")
        or_value = positive(params["effect_ratio"], "Odds Ratio")
        if math.isclose(or_value, 1.0, rel_tol=1e-12):
            raise ValueError("El OR debe ser distinto de 1.")
        p_case = (or_value * p_control) / (1 - p_control + or_value * p_control)
        if not 0 < p_case < 1:
            raise ValueError("La combinación de OR y proporción basal produce una probabilidad inválida.")
        n = two_proportion_n_per_group(
            p_case, p_control, alpha, power, sided, yates=bool(params.get("yates", False))
        )
        assumptions += [f"OR={or_value:g}", f"p controles={p_control:.3f}", f"p casos derivada={p_case:.3f}"]
        return finalize_groups(
            code,
            label,
            {"Casos": n, "Controles": n},
            loss,
            "pcasos=(OR·pcontroles)/(1−pcontroles+OR·pcontroles); luego comparación de dos proporciones",
            "Diseño caso-control 1:1 basado en OR mínimo detectable.",
            assumptions,
            warnings,
            extra={"derived_probability": p_case},
        )

    if code == "risk_ratio":
        p_control = clamp_probability(params["p_control"], "Riesgo basal en no expuestos/control")
        rr_value = positive(params["effect_ratio"], "Riesgo Relativo")
        if math.isclose(rr_value, 1.0, rel_tol=1e-12):
            raise ValueError("El RR debe ser distinto de 1.")
        p_exposed = rr_value * p_control
        if not 0 < p_exposed < 1:
            raise ValueError("RR × riesgo basal debe quedar entre 0 y 1.")
        n = two_proportion_n_per_group(
            p_exposed, p_control, alpha, power, sided, yates=bool(params.get("yates", False))
        )
        assumptions += [f"RR={rr_value:g}", f"riesgo basal={p_control:.3f}", f"riesgo expuesto derivado={p_exposed:.3f}"]
        return finalize_groups(
            code,
            label,
            {"Expuestos/intervención": n, "No expuestos/control": n},
            loss,
            "pexpuestos=RR·pcontrol; luego comparación de dos proporciones",
            "Cohorte o ensayo 1:1 basado en RR mínimo detectable.",
            assumptions,
            warnings,
            extra={"derived_probability": p_exposed},
        )

    if code == "diagnostic_accuracy":
        sensitivity = clamp_probability(params["sensitivity"], "Sensibilidad esperada")
        specificity = clamp_probability(params["specificity"], "Especificidad esperada")
        prevalence = clamp_probability(params["prevalence"], "Prevalencia")
        d_se = positive(params["precision_se"], "Precisión para sensibilidad")
        d_sp = positive(params["precision_sp"], "Precisión para especificidad")
        z = float(norm.ppf(1 - alpha / 2))
        n_se_total = z**2 * sensitivity * (1 - sensitivity) / (d_se**2 * prevalence)
        n_sp_total = z**2 * specificity * (1 - specificity) / (d_sp**2 * (1 - prevalence))
        selected_raw = max(n_se_total, n_sp_total)
        selected_final = inflate_loss(selected_raw, loss)
        se_final = inflate_loss(n_se_total, loss)
        sp_final = inflate_loss(n_sp_total, loss)
        assumptions += [
            f"Se={sensitivity:.3f}",
            f"Sp={specificity:.3f}",
            f"prevalencia={prevalence:.3f}",
            f"precisión Se=±{d_se:.3f}",
            f"precisión Sp=±{d_sp:.3f}",
        ]
        warnings.append("Este método estima precisión de intervalos de confianza; el parámetro de poder no interviene.")
        return SampleSizeResult(
            code,
            label,
            selected_raw,
            selected_final,
            {"Requerimiento por sensibilidad": n_se_total, "Requerimiento por especificidad": n_sp_total},
            {"Sensibilidad": se_final, "Especificidad": sp_final, "Recomendado (mayor)": selected_final},
            "NSe=Z²·Se·(1−Se)/(dSe²·Prev); NSp=Z²·Sp·(1−Sp)/(dSp²·(1−Prev)); N=max(NSe,NSp)",
            "Método de precisión diagnóstica ajustado por prevalencia (enfoque de Buderer).",
            assumptions,
            warnings,
            extra={"n_sensitivity": se_final, "n_specificity": sp_final},
            uses_power=False,
        )

    if code == "survival":
        hr = positive(params["hazard_ratio"], "Hazard Ratio")
        if math.isclose(hr, 1.0, rel_tol=1e-12):
            raise ValueError("El HR debe ser distinto de 1.")
        mortality_control = clamp_probability(params["mortality_control"], "Mortalidad acumulada control")
        mortality_treatment = clamp_probability(params["mortality_treatment"], "Mortalidad acumulada tratamiento")
        allocation_t = clamp_probability(params.get("allocation_treatment", 0.5), "Fracción de asignación al tratamiento")
        event_probability = allocation_t * mortality_treatment + (1 - allocation_t) * mortality_control
        events = (za + zb) ** 2 / ((math.log(hr) ** 2) * allocation_t * (1 - allocation_t))
        total_raw = events / event_probability
        n_t = total_raw * allocation_t
        n_c = total_raw * (1 - allocation_t)
        assumptions += [
            f"HR={hr:g}",
            f"mortalidad control={mortality_control:.1%}",
            f"mortalidad tratamiento={mortality_treatment:.1%}",
            f"fracción tratamiento={allocation_t:.2f}",
            f"eventos requeridos≈{math.ceil(events)}",
        ]
        warnings.append("La conversión de eventos a sujetos supone mortalidad acumulada promedio y riesgos proporcionales.")
        return finalize_groups(
            code,
            label,
            {"Tratamiento": n_t, "Control": n_c},
            loss,
            "D=(Zα+Zβ)²/[q(1−q)·ln(HR)²]; N=D/probabilidad promedio de evento",
            "Número de eventos de Schoenfeld para log-rank/Mantel–Cox y conversión aproximada a participantes.",
            assumptions,
            warnings,
            extra={"required_events": int(math.ceil(events)), "event_probability": event_probability},
        )

    raise ValueError("Diseño no implementado.")


# -----------------------------------------------------------------------------
# Importación de datos y protocolos
# -----------------------------------------------------------------------------
PROTOCOL_FIELD_ALIASES: dict[str, tuple[str, ...]] = {
    "tipo_diseno": (
        "tipo de diseño",
        "tipo de diseno",
        "diseño del estudio",
        "diseno del estudio",
        "diseño estadístico",
        "diseno estadistico",
        "diseño metodológico",
        "diseno metodologico",
    ),
    "error_alfa": (
        "error alfa",
        "nivel alfa",
        "nivel de significación",
        "nivel de significacion",
        "significancia alfa",
        "alpha",
        "alfa",
    ),
    "poder_estadistico": (
        "poder estadístico",
        "poder estadistico",
        "potencia estadística",
        "potencia estadistica",
        "potencia del estudio",
        "poder del estudio",
        "1-beta",
        "1 - beta",
    ),
    "efecto_esperado": (
        "efecto esperado",
        "tamaño del efecto",
        "tamano del efecto",
        "diferencia mínima clínicamente importante",
        "diferencia minima clinicamente importante",
        "diferencia clínicamente importante",
        "diferencia clinicamente importante",
        "diferencia esperada",
        "dmci",
        "delta esperado",
    ),
    "variabilidad_estimada": (
        "variabilidad estimada",
        "desviación estándar estimada",
        "desviacion estandar estimada",
        "desvío estándar estimado",
        "desvio estandar estimado",
        "desviación estándar",
        "desviacion estandar",
        "desvío estándar",
        "desvio estandar",
        "sigma estimada",
        "sigma",
    ),
    "perdidas_esperadas": (
        "pérdidas esperadas",
        "perdidas esperadas",
        "pérdidas de seguimiento",
        "perdidas de seguimiento",
        "tasa de pérdidas",
        "tasa de perdidas",
        "abandono esperado",
    ),
    "precision": (
        "precisión absoluta",
        "precision absoluta",
        "error máximo aceptable",
        "error maximo aceptable",
        "precisión esperada",
        "precision esperada",
    ),
    "margen": (
        "margen de no inferioridad",
        "margen de equivalencia",
        "margen clínico",
        "margen clinico",
        "margen absoluto",
    ),
    "proporcion_grupo_1": (
        "proporción grupo 1",
        "proporcion grupo 1",
        "proporción tratamiento",
        "proporcion tratamiento",
        "tasa tratamiento",
    ),
    "proporcion_grupo_2": (
        "proporción grupo 2",
        "proporcion grupo 2",
        "proporción control",
        "proporcion control",
        "tasa control",
    ),
    "proporcion_esperada": ("proporción esperada", "proporcion esperada"),
    "proporcion_control": (
        "riesgo basal en control",
        "proporción basal en control",
        "proporcion basal en control",
        "proporción de expuestos en controles",
        "proporcion de expuestos en controles",
    ),
    "prevalencia": ("prevalencia", "prevalencia de la enfermedad"),
    "sensibilidad": ("sensibilidad esperada", "sensibilidad"),
    "especificidad": ("especificidad esperada", "especificidad"),
    "precision_sensibilidad": (
        "precisión para sensibilidad",
        "precision para sensibilidad",
        "error de sensibilidad",
    ),
    "precision_especificidad": (
        "precisión para especificidad",
        "precision para especificidad",
        "error de especificidad",
    ),
    "odds_ratio": ("odds ratio", "razón de momios", "razon de momios"),
    "riesgo_relativo": ("riesgo relativo", "relative risk"),
    "hazard_ratio": ("hazard ratio", "razón de riesgos", "razon de riesgos"),
    "mortalidad_control": (
        "mortalidad acumulada control",
        "mortalidad en control",
        "tasa de mortalidad control",
    ),
    "mortalidad_tratamiento": (
        "mortalidad acumulada tratamiento",
        "mortalidad en tratamiento",
        "tasa de mortalidad tratamiento",
    ),
    "poblacion": ("tamaño de la población", "tamano de la poblacion", "población finita", "poblacion finita"),
    "p01": ("discordantes 0 a 1", "discordantes 0→1", "p01"),
    "p10": ("discordantes 1 a 0", "discordantes 1→0", "p10"),
}

NUMERIC_PROTOCOL_FIELDS = set(PROTOCOL_FIELD_ALIASES) - {"tipo_diseno"}


def _accentless_text(value: Any) -> str:
    text = str(value or "").replace("\u00a0", " ")
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    return text.lower()


def _parse_locale_number(value: Any) -> float:
    if isinstance(value, (int, float, np.integer, np.floating)):
        number = float(value)
        if not math.isfinite(number):
            raise ValueError("El valor numérico no es finito.")
        return number

    text = str(value or "").strip().replace("\u00a0", " ")
    # Evita interpretar la notación de potencia (1−β) como el valor numérico.
    text = re.sub(r"^\s*\([^)]*\)\s*", "", text)
    text = re.sub(r"^\s*1\s*[-−–]\s*(?:beta|β)\s*[:=]?\s*", "", text, flags=re.IGNORECASE)
    match = re.search(r"[-+]?\d(?:[\d\s.,]*\d)?|[-+]?\d", text)
    if not match:
        raise ValueError(f"No se encontró un número en {value!r}.")
    token = match.group(0).replace(" ", "")
    if "," in token and "." in token:
        decimal_sep = "," if token.rfind(",") > token.rfind(".") else "."
        thousands_sep = "." if decimal_sep == "," else ","
        token = token.replace(thousands_sep, "").replace(decimal_sep, ".")
    elif "," in token:
        parts = token.split(",")
        token = "".join(parts[:-1]) + "." + parts[-1] if len(parts[-1]) <= 4 else "".join(parts)
    elif token.count(".") > 1:
        parts = token.split(".")
        token = "".join(parts[:-1]) + "." + parts[-1]
    number = float(token)
    if not math.isfinite(number):
        raise ValueError("El valor numérico no es finito.")
    return number


def _extract_docx_text(raw: bytes) -> str:
    document = Document(io.BytesIO(raw))
    blocks: list[str] = []
    for paragraph in document.paragraphs:
        if paragraph.text.strip():
            blocks.append(paragraph.text.strip())
    for table in document.tables:
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            if any(cells):
                blocks.append("\t".join(cells))
    for section in document.sections:
        for paragraph in section.header.paragraphs:
            if paragraph.text.strip():
                blocks.append(paragraph.text.strip())
        for paragraph in section.footer.paragraphs:
            if paragraph.text.strip():
                blocks.append(paragraph.text.strip())
    return "\n".join(blocks)


def _extract_pdf_text(raw: bytes) -> str:
    reader = PdfReader(io.BytesIO(raw))
    if reader.is_encrypted:
        try:
            reader.decrypt("")
        except Exception as exc:
            raise ValueError("El PDF está cifrado y no pudo abrirse sin contraseña.") from exc
    pages = [(page.extract_text() or "").strip() for page in reader.pages]
    return "\n\n".join(page for page in pages if page)


def _extract_legacy_doc_text(raw: bytes) -> tuple[str, list[str]]:
    warnings: list[str] = []
    with tempfile.NamedTemporaryFile(suffix=".doc", delete=False) as tmp:
        tmp.write(raw)
        temp_path = tmp.name
    try:
        for command_name in ("antiword", "catdoc"):
            executable = shutil.which(command_name)
            if executable:
                completed = subprocess.run(
                    [executable, temp_path],
                    check=False,
                    capture_output=True,
                    timeout=30,
                )
                if completed.returncode == 0 and completed.stdout.strip():
                    for encoding in ("utf-8", "latin-1"):
                        try:
                            text = completed.stdout.decode(encoding)
                            if text.strip():
                                return text, warnings
                        except UnicodeDecodeError:
                            continue

        # Rescate heurístico para Word binario cuando antiword/catdoc no están instalados.
        ascii_chunks = [m.decode("latin-1", errors="ignore") for m in re.findall(rb"[\x20-\x7e\xa0-\xff]{5,}", raw)]
        utf16_chunks: list[str] = []
        for match in re.findall(rb"(?:[\x20-\x7e\xa0-\xff]\x00){5,}", raw):
            try:
                utf16_chunks.append(match.decode("utf-16le", errors="ignore"))
            except Exception:
                pass
        text = "\n".join(dict.fromkeys(chunk.strip() for chunk in utf16_chunks + ascii_chunks if chunk.strip()))
        warnings.append(
            "El archivo .DOC legado se leyó mediante recuperación heurística. Para máxima confiabilidad, guárdelo como DOCX."
        )
        return text, warnings
    finally:
        try:
            os.unlink(temp_path)
        except OSError:
            pass


def extract_protocol_text(raw: bytes, suffix: str) -> tuple[str, list[str]]:
    warnings: list[str] = []
    if suffix == ".pdf":
        text = _extract_pdf_text(raw)
        if len(text.strip()) < 30:
            raise ValueError(
                "No se detectó texto utilizable en el PDF. Si es un documento escaneado, aplique OCR o conviértalo a PDF con texto seleccionable."
            )
        return text, warnings
    if suffix in {".docx", ".docm"}:
        text = _extract_docx_text(raw)
        if len(text.strip()) < 20:
            raise ValueError("El archivo Word no contiene texto utilizable.")
        return text, warnings
    if suffix == ".doc":
        text, legacy_warnings = _extract_legacy_doc_text(raw)
        warnings.extend(legacy_warnings)
        if len(text.strip()) < 20:
            raise ValueError("No fue posible recuperar texto del archivo .DOC. Conviértalo a DOCX.")
        return text, warnings
    raise ValueError("Formato documental no compatible.")


def _value_after_alias(lines: list[str], aliases: tuple[str, ...]) -> str | None:
    normalized_aliases = sorted({_accentless_text(alias).strip() for alias in aliases}, key=len, reverse=True)
    for index, line in enumerate(lines):
        normalized_line = _accentless_text(line)
        for alias in normalized_aliases:
            position = normalized_line.find(alias)
            if position < 0:
                continue
            tail = line[position + len(alias):].strip()
            tail = re.sub(r"^[\s\t:=\-–—·•()\[\]]+", "", tail).strip()
            if tail:
                return tail
            if index + 1 < len(lines):
                next_line = lines[index + 1].strip()
                if next_line:
                    return next_line
    return None


def infer_design_from_text(text: str) -> str | None:
    token = normalize_token(text)

    def has(*parts: str) -> bool:
        return any(normalize_token(part) in token for part in parts)

    binary_hint = has("dicotomica", "binaria", "proporcion", "evento", "respuesta si no", "tasa")
    continuous_hint = has("continua", "media", "promedio", "desviacion estandar")

    if has("mcnemar"):
        return "mcnemar"
    if has("validacion diagnostica", "precision diagnostica") or (has("sensibilidad") and has("especificidad")):
        return "diagnostic_accuracy"
    if has("mantel cox", "log rank", "logrank", "curvas de supervivencia", "hazard ratio"):
        return "survival"
    if has("odds ratio", "razon de momios"):
        return "odds_ratio"
    if has("riesgo relativo", "relative risk"):
        return "risk_ratio"
    if has("no inferioridad", "no-inferioridad"):
        return "noninferiority_binary" if binary_hint and not continuous_hint else "noninferiority_continuous"
    if has("equivalencia", "tost"):
        return "equivalence_binary" if binary_hint and not continuous_hint else "equivalence_continuous"
    if has("superioridad"):
        return "superiority_binary" if binary_hint and not continuous_hint else "superiority_continuous"
    if has("medias apareadas", "media apareada", "antes y despues", "medidas pareadas"):
        return "means_paired"
    if has("proporciones apareadas", "datos apareados dicotomicos"):
        return "mcnemar"
    if has("medias independientes", "dos medias independientes", "comparacion de medias independientes"):
        return "means_independent"
    if has("proporciones independientes", "dos proporciones independientes", "comparacion de proporciones"):
        return "proportions_independent"
    if has("estimacion de una proporcion", "estimacion de proporcion", "prevalencia con precision"):
        return "proportion_estimation"
    if has("estimacion de una media", "estimacion de media", "media con precision"):
        return "mean_estimation"
    return None


def extract_protocol_mapping_from_text(text: str) -> dict[str, Any]:
    cleaned = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = [re.sub(r"\s+", " ", line).strip() for line in cleaned.split("\n") if line.strip()]
    protocol: dict[str, Any] = {}

    for field_name, aliases in PROTOCOL_FIELD_ALIASES.items():
        raw_value = _value_after_alias(lines, aliases)
        if raw_value is None:
            continue
        if field_name == "tipo_diseno":
            protocol[field_name] = raw_value
        else:
            try:
                protocol[field_name] = _parse_locale_number(raw_value)
            except ValueError:
                continue

    if "tipo_diseno" not in protocol:
        inferred = infer_design_from_text(cleaned)
        if inferred:
            protocol["tipo_diseno"] = inferred

    # Parámetros específicos pueden completar los cinco campos mínimos.
    design_code: str | None = None
    if protocol.get("tipo_diseno") is not None:
        try:
            design_code = resolve_design(protocol["tipo_diseno"])
        except ValueError:
            inferred = infer_design_from_text(str(protocol["tipo_diseno"]))
            if inferred:
                protocol["tipo_diseno"] = inferred
                design_code = inferred

    if "efecto_esperado" not in protocol:
        effect_candidates: dict[str, tuple[str, ...]] = {
            "diagnostic_accuracy": ("sensibilidad", "especificidad"),
            "odds_ratio": ("odds_ratio",),
            "risk_ratio": ("riesgo_relativo",),
            "survival": ("hazard_ratio",),
            "noninferiority_continuous": ("margen",),
            "noninferiority_binary": ("margen",),
            "equivalence_continuous": ("margen",),
            "equivalence_binary": ("margen",),
        }
        for key in effect_candidates.get(design_code or "", ("margen",)):
            if key in protocol:
                protocol["efecto_esperado"] = protocol[key]
                break

    if "variabilidad_estimada" not in protocol:
        variability_candidates: dict[str, tuple[str, ...]] = {
            "diagnostic_accuracy": ("prevalencia",),
            "odds_ratio": ("proporcion_control", "proporcion_grupo_2"),
            "risk_ratio": ("proporcion_control", "proporcion_grupo_2"),
            "survival": ("mortalidad_control",),
            "proportion_estimation": ("precision", "proporcion_esperada"),
            "noninferiority_binary": ("proporcion_grupo_2",),
            "equivalence_binary": ("proporcion_grupo_2",),
        }
        for key in variability_candidates.get(design_code or "", ()):
            if key in protocol:
                protocol["variabilidad_estimada"] = protocol[key]
                break

    return protocol


def canonicalize_protocol_mapping(data: dict[str, Any]) -> dict[str, Any]:
    canonical: dict[str, Any] = {}
    alias_lookup: dict[str, str] = {}
    for canonical_name, aliases in PROTOCOL_FIELD_ALIASES.items():
        alias_lookup[normalize_token(canonical_name)] = canonical_name
        for alias in aliases:
            alias_lookup[normalize_token(alias)] = canonical_name

    for key, value in data.items():
        canonical_key = alias_lookup.get(normalize_token(key), normalize_token(key))
        canonical[canonical_key] = value
    return canonical


def parse_protocol(uploaded_file: Any) -> dict[str, Any]:
    raw = uploaded_file.getvalue()
    suffix = os.path.splitext(uploaded_file.name)[1].lower()
    extraction_warnings: list[str] = []
    extracted_text = ""

    if suffix == ".json":
        data = json.loads(raw.decode("utf-8-sig"))
    elif suffix in {".yaml", ".yml"}:
        data = yaml.safe_load(raw.decode("utf-8-sig"))
    elif suffix in {".pdf", ".docx", ".docm", ".doc"}:
        extracted_text, extraction_warnings = extract_protocol_text(raw, suffix)
        data = extract_protocol_mapping_from_text(extracted_text)
    else:
        raise ValueError("El protocolo debe ser JSON, YAML, PDF o Word (DOCX/DOCM/DOC).")

    if not isinstance(data, dict):
        raise ValueError("El archivo de protocolo debe contener parámetros identificables.")
    data = canonicalize_protocol_mapping(data)

    # Convierte campos numéricos aun cuando JSON/YAML los hayan guardado como texto con coma o porcentaje.
    for key in list(data):
        if key in NUMERIC_PROTOCOL_FIELDS and data[key] not in (None, ""):
            data[key] = _parse_locale_number(data[key])

    if data.get("tipo_diseno") is None and extracted_text:
        inferred = infer_design_from_text(extracted_text)
        if inferred:
            data["tipo_diseno"] = inferred

    required = {"tipo_diseno", "error_alfa", "poder_estadistico", "efecto_esperado", "variabilidad_estimada"}
    missing = sorted(key for key in required if key not in data or data[key] in (None, ""))
    if missing:
        readable = ", ".join(missing)
        raise ValueError(
            "No fue posible identificar todos los parámetros obligatorios. Faltan: "
            + readable
            + ". En PDF/Word use etiquetas claras, por ejemplo: 'Error alfa: 0,05'."
        )

    data["_source_format"] = suffix.lstrip(".").upper()
    data["_extraction_warnings"] = extraction_warnings
    if extracted_text:
        data["_text_preview"] = extracted_text[:1500]
    return data

def resolve_design(value: Any) -> str:
    token = normalize_token(value)
    if token in DESIGNS:
        return token
    if token in ALIASES:
        return ALIASES[token]
    # Búsqueda tolerante por palabras del rótulo.
    for code, label in DESIGNS.items():
        normalized_label = normalize_token(label)
        if token == normalized_label or token in normalized_label or normalized_label in token:
            return code
    inferred = infer_design_from_text(str(value))
    if inferred:
        return inferred
    raise ValueError(f"tipo_diseno no reconocido: {value!r}")


def apply_protocol_to_state(protocol: dict[str, Any]) -> None:
    code = resolve_design(protocol["tipo_diseno"])
    alpha = float(protocol["error_alfa"])
    power = float(protocol["poder_estadistico"])
    if alpha > 1:
        alpha /= 100
    if power > 1:
        power /= 100
    if not 0 < alpha < 1:
        raise ValueError("error_alfa debe expresarse como proporción (0,05) o porcentaje (5).")
    if not 0.50 <= power < 1:
        raise ValueError("poder_estadistico debe estar entre 0,50 y menos de 1.")

    effect = float(protocol["efecto_esperado"])
    variability = float(protocol["variabilidad_estimada"])
    if not math.isfinite(effect) or not math.isfinite(variability):
        raise ValueError("efecto_esperado y variabilidad_estimada deben ser valores numéricos finitos.")

    def bounded_probability(value: float, fallback: float = 0.50) -> float:
        numeric = float(value)
        if 1 < numeric <= 100:
            numeric /= 100
        if 0 <= numeric <= 1:
            return min(max(numeric, 0.001), 0.999)
        return fallback

    def probability_or_points(value: float) -> float:
        numeric = abs(float(value))
        return numeric / 100 if 1 < numeric <= 100 else numeric

    st.session_state.design_code = code
    st.session_state.alpha = alpha
    st.session_state.power = power

    # Interpretación automática de los dos parámetros genéricos según el diseño.
    if code in {"mean_estimation", "means_independent", "means_paired", "superiority_continuous"}:
        st.session_state.delta = max(abs(effect), 0.0001)
        st.session_state.sd = max(abs(variability), 0.0001)
        st.session_state.sd_diff = max(abs(variability), 0.0001)
    elif code == "proportion_estimation":
        st.session_state.p = bounded_probability(effect)
        st.session_state.precision_prop = max(min(probability_or_points(variability), 0.50), 0.001)
    elif code in {"proportions_independent", "superiority_binary"}:
        baseline = bounded_probability(variability)
        absolute_effect = probability_or_points(effect)
        st.session_state.p2 = baseline
        st.session_state.p1 = min(max(baseline + absolute_effect, 0.001), 0.999)
    elif code in {"noninferiority_continuous", "equivalence_continuous"}:
        st.session_state.margin_cont = max(abs(effect), 0.0001)
        st.session_state.sd = max(abs(variability), 0.0001)
    elif code == "noninferiority_binary":
        baseline = bounded_probability(variability, 0.80)
        st.session_state.margin_bin = max(min(probability_or_points(effect), 0.50), 0.001)
        st.session_state.p1_ni = baseline
        st.session_state.p2_ni = baseline
    elif code == "equivalence_binary":
        baseline = bounded_probability(variability, 0.80)
        st.session_state.margin_bin = max(min(probability_or_points(effect), 0.50), 0.001)
        st.session_state.p1_eq = baseline
        st.session_state.p2_eq = baseline
    elif code == "mcnemar":
        discordance_total = min(max(probability_or_points(variability), 0.01), 0.98)
        discordance_difference = min(probability_or_points(effect), discordance_total - 0.001)
        st.session_state.p01 = (discordance_total + discordance_difference) / 2
        st.session_state.p10 = (discordance_total - discordance_difference) / 2
    elif code in {"odds_ratio", "risk_ratio"}:
        st.session_state.effect_ratio = max(abs(effect), 0.01)
        st.session_state.p_control = bounded_probability(variability, 0.20)
    elif code == "diagnostic_accuracy":
        accuracy = bounded_probability(effect, 0.90)
        st.session_state.sensitivity = accuracy
        st.session_state.specificity = accuracy
        st.session_state.prevalence = bounded_probability(variability, 0.20)
    elif code == "survival":
        hr = max(abs(effect), 0.01)
        mortality_control = bounded_probability(variability, 0.30)
        st.session_state.hazard_ratio = hr
        st.session_state.mortality_control = mortality_control
        st.session_state.mortality_treatment = min(max(mortality_control * hr, 0.001), 0.999)

    if "precision" in protocol and protocol["precision"] is not None:
        precision_value = abs(float(protocol["precision"]))
        if code == "mean_estimation":
            st.session_state.precision_mean = max(precision_value, 0.0001)
        elif code == "proportion_estimation":
            st.session_state.precision_prop = max(min(probability_or_points(precision_value), 0.50), 0.001)
    if "margen" in protocol and protocol["margen"] is not None:
        margin_value = abs(float(protocol["margen"]))
        if code in {"noninferiority_continuous", "equivalence_continuous"}:
            st.session_state.margin_cont = max(margin_value, 0.0001)
        elif code in {"noninferiority_binary", "equivalence_binary"}:
            st.session_state.margin_bin = max(min(probability_or_points(margin_value), 0.50), 0.001)
    if "perdidas_esperadas" in protocol and protocol["perdidas_esperadas"] is not None:
        loss_value = float(protocol["perdidas_esperadas"])
        loss_percent = loss_value * 100 if loss_value <= 1 else loss_value
        st.session_state.loss_percent = int(round(min(max(loss_percent, 0), 50)))

    # Los campos específicos, cuando existen, prevalecen sobre la interpretación genérica.
    if "proporcion_grupo_1" in protocol and protocol["proporcion_grupo_1"] is not None:
        target = "p1_eq" if code == "equivalence_binary" else "p1_ni" if code == "noninferiority_binary" else "p1"
        st.session_state[target] = bounded_probability(protocol["proporcion_grupo_1"])
    if "proporcion_grupo_2" in protocol and protocol["proporcion_grupo_2"] is not None:
        target = "p2_eq" if code == "equivalence_binary" else "p2_ni" if code == "noninferiority_binary" else "p2"
        st.session_state[target] = bounded_probability(protocol["proporcion_grupo_2"])

    optional_map = {
        "proporcion_esperada": "p",
        "proporcion_control": "p_control",
        "diferencia_esperada": "expected_diff",
        "prevalencia": "prevalence",
        "sensibilidad": "sensitivity",
        "especificidad": "specificity",
        "precision_sensibilidad": "precision_se",
        "precision_especificidad": "precision_sp",
        "odds_ratio": "effect_ratio",
        "riesgo_relativo": "effect_ratio",
        "hazard_ratio": "hazard_ratio",
        "mortalidad_control": "mortality_control",
        "mortalidad_tratamiento": "mortality_treatment",
        "poblacion": "population",
        "p01": "p01",
        "p10": "p10",
    }
    probability_fields = {
        "proporcion_esperada",
        "proporcion_control",
        "prevalencia",
        "sensibilidad",
        "especificidad",
        "precision_sensibilidad",
        "precision_especificidad",
        "mortalidad_control",
        "mortalidad_tratamiento",
        "p01",
        "p10",
    }
    for source_key, state_key in optional_map.items():
        if source_key in protocol and protocol[source_key] is not None:
            value = protocol[source_key]
            if source_key in probability_fields:
                value = bounded_probability(value)
            st.session_state[state_key] = value


def load_dataset(uploaded_file: Any) -> tuple[pd.DataFrame | None, dict[str, Any]]:
    suffix = os.path.splitext(uploaded_file.name)[1].lower()
    metadata: dict[str, Any] = {"name": uploaded_file.name, "type": suffix}
    if suffix == ".csv":
        raw = uploaded_file.getvalue()
        last_error: Exception | None = None
        for encoding in ("utf-8-sig", "latin-1"):
            for sep in (None, ",", ";", "\t"):
                try:
                    df = pd.read_csv(io.BytesIO(raw), encoding=encoding, sep=sep, engine="python")
                    if df.shape[1] > 0:
                        metadata["rows"] = len(df)
                        metadata["columns"] = len(df.columns)
                        return df, metadata
                except Exception as exc:  # pragma: no cover - depende del archivo
                    last_error = exc
        raise ValueError(f"No se pudo interpretar el CSV: {last_error}")

    if suffix in {".xlsx", ".xlsm"}:
        xls = pd.ExcelFile(io.BytesIO(uploaded_file.getvalue()), engine="openpyxl")
        sheet = st.sidebar.selectbox("Hoja de Excel", xls.sheet_names, key="excel_sheet")
        df = pd.read_excel(xls, sheet_name=sheet, engine="openpyxl")
        metadata.update({"sheet": sheet, "rows": len(df), "columns": len(df.columns)})
        return df, metadata

    if suffix in {".sqlite", ".sqlite3", ".db"}:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        tmp.write(uploaded_file.getvalue())
        tmp.close()
        engine = create_engine(f"sqlite+pysqlite:///{tmp.name}")
        table_names = inspect(engine).get_table_names()
        if not table_names:
            raise ValueError("La base SQLite no contiene tablas visibles.")
        table = st.sidebar.selectbox("Tabla SQLite", table_names, key="sqlite_table")
        df = pd.read_sql_table(table, con=engine)
        metadata.update({"table": table, "rows": len(df), "columns": len(df.columns), "temp_path": tmp.name})
        return df, metadata

    raise ValueError("Formato no compatible. Use XLSX, CSV, SQLite, SQLite3 o DB.")


def numeric_columns(df: pd.DataFrame) -> list[str]:
    return [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]


# -----------------------------------------------------------------------------
# Reportes DOCX y anexos CONSORT
# -----------------------------------------------------------------------------
CONSORT_ITEMS = [
    ("Título y resumen", "Identificar el estudio como aleatorizado y presentar un resumen estructurado."),
    ("Fundamento", "Explicar el contexto científico y los objetivos o hipótesis."),
    ("Diseño", "Describir el diseño, la razón de asignación y cualquier cambio posterior al inicio."),
    ("Participantes", "Precisar criterios de elegibilidad, centros y lugares de recolección."),
    ("Intervenciones", "Detallar cada intervención con información suficiente para reproducirla."),
    ("Desenlaces", "Definir desenlaces primarios/secundarios y cómo/cuándo se evaluaron."),
    ("Tamaño de muestra", "Informar cómo se determinó la muestra y cualquier análisis intermedio o regla de detención."),
    ("Secuencia aleatoria", "Describir el método de generación de la secuencia y restricciones."),
    ("Ocultamiento", "Explicar el mecanismo utilizado para ocultar la asignación."),
    ("Implementación", "Indicar quién generó la secuencia, incorporó participantes y asignó intervenciones."),
    ("Enmascaramiento", "Informar quién estuvo cegado y cómo se mantuvo el cegamiento."),
    ("Métodos estadísticos", "Describir análisis primarios, secundarios y métodos adicionales."),
    ("Flujo de participantes", "Documentar asignación, seguimiento, pérdidas, exclusiones y análisis."),
    ("Reclutamiento", "Informar fechas de reclutamiento y seguimiento, y motivo de finalización."),
    ("Datos basales", "Presentar características demográficas y clínicas por grupo."),
    ("Números analizados", "Indicar denominadores y si se analizó según asignación original."),
    ("Resultados y estimación", "Informar efecto, precisión e intervalos de confianza."),
    ("Análisis adicionales", "Distinguir análisis preespecificados de exploratorios."),
    ("Daños", "Describir eventos adversos o efectos no deseados por grupo."),
    ("Limitaciones", "Discutir sesgos potenciales, imprecisión y multiplicidad."),
    ("Generalización", "Analizar la aplicabilidad externa de los hallazgos."),
    ("Interpretación", "Ofrecer una interpretación consistente con beneficios, daños y evidencia disponible."),
    ("Registro", "Indicar registro del ensayo y número de inscripción."),
    ("Protocolo", "Señalar dónde puede accederse al protocolo completo."),
    ("Financiación", "Informar fuentes de financiación, apoyos y papel de los financiadores."),
]


def set_doc_defaults(doc: Document) -> None:
    section = doc.sections[0]
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)
    styles = doc.styles
    styles["Normal"].font.name = "Aptos"
    styles["Normal"].font.size = Pt(10)
    styles["Title"].font.name = "Aptos Display"
    styles["Title"].font.size = Pt(22)
    styles["Heading 1"].font.name = "Aptos Display"
    styles["Heading 1"].font.size = Pt(15)
    styles["Heading 2"].font.name = "Aptos Display"
    styles["Heading 2"].font.size = Pt(12)


def add_key_value_table(doc: Document, rows: list[tuple[str, Any]]) -> None:
    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    table.rows[0].cells[0].text = "Parámetro"
    table.rows[0].cells[1].text = "Valor"
    for key, value in rows:
        cells = table.add_row().cells
        cells[0].text = str(key)
        cells[1].text = str(value)


def consort_paragraph(result: SampleSizeResult, params: dict[str, Any]) -> str:
    alpha = float(params.get("alpha", 0.05))
    power = float(params.get("power", 0.80))
    sided = params.get("sided", "Bilateral").lower()
    losses = float(params.get("loss_rate", 0.0))
    variability = params.get("sd", params.get("sd_diff", "no aplicable"))
    effect = params.get("delta", params.get("margin", params.get("effect_ratio", "según diseño")))
    power_clause = (
        f"un poder estadístico de {power:.0%}" if result.uses_power else "un criterio de precisión del intervalo de confianza"
    )
    return (
        f"El tamaño de muestra se determinó a priori para el desenlace principal mediante {result.method.lower()} "
        f"Se asumió un error alfa de {alpha:.3f}, un contraste {sided}, {power_clause}, "
        f"un efecto o margen de {effect} y una variabilidad/proporción basal de {variability}. "
        f"El cálculo inicial fue de {math.ceil(result.n_raw_total)} participantes y se ajustó por una pérdida esperada "
        f"de {losses:.1%}, obteniéndose un requerimiento final de {result.n_final_total}. "
        "Estos parámetros, su fuente clínica y cualquier modificación posterior deberán declararse en Métodos, "
        "de acuerdo con el ítem de tamaño de muestra de CONSORT 2010."
    )


def build_word_report(
    result: SampleSizeResult,
    params: dict[str, Any],
    sensitivity_df: pd.DataFrame,
    dataset_metadata: dict[str, Any] | None,
) -> bytes:
    doc = Document()
    set_doc_defaults(doc)
    title = doc.add_paragraph(style="Title")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.add_run("Informe de cálculo del tamaño de muestra")
    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.add_run(f"{APP_NAME} · v{APP_VERSION} · {datetime.now():%d/%m/%Y %H:%M}")

    doc.add_heading("1. Resumen del diseño", level=1)
    doc.add_paragraph(result.design_label)
    add_key_value_table(
        doc,
        [
            ("N total recomendado", result.n_final_total),
            ("N inicial sin pérdidas", math.ceil(result.n_raw_total)),
            ("Método", result.method),
            ("Distribución", "; ".join(f"{k}: {v}" for k, v in result.groups_final.items())),
        ],
    )

    doc.add_heading("2. Justificación matemática", level=1)
    doc.add_paragraph(result.formula)
    doc.add_paragraph("Supuestos utilizados: " + "; ".join(result.assumptions) + ".")
    if result.warnings:
        p = doc.add_paragraph()
        p.add_run("Advertencias metodológicas: ").bold = True
        p.add_run(" ".join(result.warnings))

    doc.add_heading("3. Diferencia mínima clínicamente importante", level=1)
    doc.add_paragraph(
        "El efecto esperado debe representar la diferencia mínima que modificaría una decisión clínica, no la diferencia "
        "que simplemente podría alcanzar significación estadística. Su valor debe justificarse con literatura previa, "
        "datos piloto o consenso clínico-estadístico."
    )

    doc.add_heading("4. Párrafo para reporte CONSORT 2010", level=1)
    doc.add_paragraph(consort_paragraph(result, params))

    doc.add_heading("5. Análisis de sensibilidad", level=1)
    table = doc.add_table(rows=1, cols=len(sensitivity_df.columns))
    table.style = "Table Grid"
    for idx, column in enumerate(sensitivity_df.columns):
        table.rows[0].cells[idx].text = str(column)
    for _, row in sensitivity_df.iterrows():
        cells = table.add_row().cells
        for idx, value in enumerate(row):
            cells[idx].text = str(value)

    if dataset_metadata:
        doc.add_heading("6. Fuente de datos importada", level=1)
        doc.add_paragraph(
            "; ".join(f"{k}: {v}" for k, v in dataset_metadata.items() if k != "temp_path")
        )

    doc.add_heading("7. Trazabilidad y límites", level=1)
    doc.add_paragraph(
        "Este informe documenta un cálculo de planificación. Los diseños con múltiples desenlaces, análisis intermedios, "
        "conglomerados, estratificación, multiplicidad, riesgos no proporcionales o requisitos regulatorios deben validarse "
        "con un bioestadístico y, cuando corresponda, mediante simulación."
    )

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


def build_consort_docx() -> bytes:
    doc = Document()
    set_doc_defaults(doc)
    title = doc.add_paragraph(style="Title")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.add_run("Checklist operativo CONSORT 2010")
    doc.add_paragraph(
        "Versión para trabajo interno, redactada en forma resumida y parafraseada. Debe cotejarse con el documento oficial."
    )
    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    headers = ["N.º", "Dominio", "Verificación", "Página/nota"]
    for i, header in enumerate(headers):
        table.rows[0].cells[i].text = header
    for i, (domain, check) in enumerate(CONSORT_ITEMS, 1):
        cells = table.add_row().cells
        cells[0].text = str(i)
        cells[1].text = domain
        cells[2].text = check
        cells[3].text = ""
    doc.add_paragraph(
        "Fuente normativa: CONSORT 2010 Statement y checklist oficial. Para ensayos nuevos, verificar además las actualizaciones vigentes."
    )
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


def build_consort_pdf() -> bytes:
    bio = io.BytesIO()
    doc = SimpleDocTemplate(
        bio,
        pagesize=A4,
        rightMargin=1.2 * cm,
        leftMargin=1.2 * cm,
        topMargin=1.2 * cm,
        bottomMargin=1.2 * cm,
    )
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="CenterTitle", parent=styles["Title"], alignment=TA_CENTER, fontSize=16, leading=19))
    story = [Paragraph("Checklist operativo CONSORT 2010", styles["CenterTitle"]), Spacer(1, 0.25 * cm)]
    story.append(
        Paragraph(
            "Versión resumida y parafraseada para trabajo interno; cotejar con el checklist oficial.",
            styles["BodyText"],
        )
    )
    story.append(Spacer(1, 0.25 * cm))
    data = [["N.º", "Dominio", "Verificación", "Página/nota"]]
    for i, (domain, check) in enumerate(CONSORT_ITEMS, 1):
        data.append([str(i), domain, Paragraph(check, styles["BodyText"]), ""])
    table = Table(data, colWidths=[1.0 * cm, 3.5 * cm, 11.0 * cm, 2.2 * cm], repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DCEAF7")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#12344D")),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 7.2),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("GRID", (0, 0), (-1, -1), 0.35, colors.grey),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7FAFC")]),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    story.append(table)
    story.append(Spacer(1, 0.25 * cm))
    story.append(
        Paragraph(
            "Fuente normativa: CONSORT 2010 Statement. Para ensayos nuevos, verificar además las actualizaciones vigentes.",
            styles["BodyText"],
        )
    )
    doc.build(story)
    return bio.getvalue()



def _protocol_display_rows(protocol: dict[str, Any]) -> list[tuple[str, str]]:
    labels = {
        "tipo_diseno": "Tipo de diseño",
        "error_alfa": "Error alfa",
        "poder_estadistico": "Poder estadístico",
        "efecto_esperado": "Efecto esperado",
        "variabilidad_estimada": "Variabilidad estimada",
        "perdidas_esperadas": "Pérdidas esperadas",
    }
    return [(labels[key], str(protocol[key])) for key in labels if key in protocol]


def build_protocol_docx_template(protocol: dict[str, Any]) -> bytes:
    document = Document()
    document.add_heading("Archivo de configuración de protocolo", level=1)
    document.add_paragraph(
        "Complete los valores conservando las etiquetas. BioSize Clinical leerá automáticamente este archivo Word."
    )
    table = document.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    table.rows[0].cells[0].text = "Parámetro"
    table.rows[0].cells[1].text = "Valor"
    for label, value in _protocol_display_rows(protocol):
        cells = table.add_row().cells
        cells[0].text = label
        cells[1].text = value
    document.add_paragraph(
        "También pueden agregarse parámetros específicos, por ejemplo: prevalencia, sensibilidad, especificidad, margen, hazard ratio o proporciones por grupo."
    )
    bio = io.BytesIO()
    document.save(bio)
    return bio.getvalue()


def build_protocol_pdf_template(protocol: dict[str, Any]) -> bytes:
    bio = io.BytesIO()
    doc = SimpleDocTemplate(bio, pagesize=A4, rightMargin=2 * cm, leftMargin=2 * cm, topMargin=2 * cm, bottomMargin=2 * cm)
    styles = getSampleStyleSheet()
    story: list[Any] = [
        Paragraph("Archivo de configuración de protocolo", styles["Title"]),
        Spacer(1, 0.3 * cm),
        Paragraph(
            "Complete los valores conservando las etiquetas. BioSize Clinical leerá automáticamente este PDF cuando contenga texto seleccionable.",
            styles["BodyText"],
        ),
        Spacer(1, 0.4 * cm),
    ]
    rows = [["Parámetro", "Valor"]] + [[label, value] for label, value in _protocol_display_rows(protocol)]
    table = Table(rows, colWidths=[9 * cm, 5 * cm])
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#dbe7f0")),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    story.append(table)
    doc.build(story)
    return bio.getvalue()


# -----------------------------------------------------------------------------
# Interfaz Streamlit
# -----------------------------------------------------------------------------
st.set_page_config(page_title=APP_NAME, page_icon="📐", layout="wide")

st.markdown(
    """
    <style>
      .main .block-container {padding-top: 1.5rem; padding-bottom: 3rem; max-width: 1450px;}
      .hero {padding: 1.2rem 1.4rem; border: 1px solid #dbe7f0; border-radius: 18px;
             background: linear-gradient(135deg, #f6fbff 0%, #eef7f4 100%); margin-bottom: 1rem;}
      .hero h1 {margin: 0; color: #12344d; font-size: 2.2rem;}
      .hero p {margin: .35rem 0 0 0; color: #496273;}
      .metric-card {border: 1px solid #dbe7f0; border-radius: 16px; padding: 1rem;
                    background: #ffffff; min-height: 120px; box-shadow: 0 2px 10px rgba(22,52,74,.05);}
      .metric-label {font-size: .86rem; color: #627787; text-transform: uppercase; letter-spacing: .04em;}
      .metric-value {font-size: 2.15rem; font-weight: 750; color: #0f5e62; margin-top: .15rem;}
      .metric-note {font-size: .86rem; color: #6f7f8b;}
      .technical-note {border-left: 4px solid #0f7b78; padding: .65rem .9rem; background: #f1fbfa; border-radius: 8px;}
      div[data-testid="stDownloadButton"] button {border-radius: 10px;}
    </style>
    """,
    unsafe_allow_html=True,
)

# Estado inicial
DEFAULTS = {
    "design_code": "means_independent",
    "alpha": 0.05,
    "power": 0.80,
    "sided": "Bilateral",
    "loss_percent": 10,
    "delta": 5.0,
    "sd": 10.0,
    "sd_diff": 8.0,
    "precision_mean": 2.0,
    "precision_prop": 0.05,
    "p": 0.50,
    "p1": 0.60,
    "p2": 0.40,
    "p1_ni": 0.78,
    "p2_ni": 0.80,
    "p1_eq": 0.80,
    "p2_eq": 0.80,
    "p01": 0.20,
    "p10": 0.08,
    "margin_cont": 5.0,
    "margin_bin": 0.10,
    "expected_diff": 0.0,
    "p_control": 0.20,
    "effect_ratio": 2.0,
    "sensitivity": 0.90,
    "specificity": 0.90,
    "prevalence": 0.20,
    "precision_se": 0.05,
    "precision_sp": 0.05,
    "hazard_ratio": 0.70,
    "mortality_control": 0.30,
    "mortality_treatment": 0.22,
    "allocation_treatment": 0.50,
    "finite_population": False,
    "population": 1000,
    "yates": False,
}
for key, value in DEFAULTS.items():
    st.session_state.setdefault(key, value)

with st.sidebar:
    st.header("📥 Importación")
    protocol_file = st.file_uploader(
        "Archivo de protocolo",
        type=["json", "yaml", "yml", "pdf", "docx", "docm", "doc"],
        key="protocol_upload",
        help="PDF y Word deben contener texto seleccionable o una tabla con los parámetros del protocolo.",
    )
    if protocol_file is not None:
        signature = (protocol_file.name, len(protocol_file.getvalue()), hash(protocol_file.getvalue()))
        if st.session_state.get("protocol_signature") != signature:
            try:
                protocol = parse_protocol(protocol_file)
                apply_protocol_to_state(protocol)
                st.session_state.protocol_signature = signature
                st.session_state.protocol_loaded_message = (
                    f"Protocolo aplicado: {protocol_file.name} · formato {protocol.get('_source_format', '')}"
                )
                st.session_state.protocol_detected = {
                    key: value for key, value in protocol.items() if not key.startswith("_")
                }
                st.session_state.protocol_extraction_warnings = protocol.get("_extraction_warnings", [])
                st.session_state.protocol_text_preview = protocol.get("_text_preview", "")
            except Exception as exc:
                st.error(str(exc))
    if st.session_state.get("protocol_loaded_message"):
        st.success(st.session_state.protocol_loaded_message)
    for warning in st.session_state.get("protocol_extraction_warnings", []):
        st.warning(warning)
    if st.session_state.get("protocol_detected"):
        with st.expander("Parámetros detectados", expanded=False):
            st.json(st.session_state.protocol_detected)
    if st.session_state.get("protocol_text_preview"):
        with st.expander("Vista previa del texto extraído", expanded=False):
            st.text(st.session_state.protocol_text_preview)

    data_file = st.file_uploader("Dataset", type=["xlsx", "xlsm", "csv", "sqlite", "sqlite3", "db"], key="data_upload")
    dataset: pd.DataFrame | None = None
    dataset_metadata: dict[str, Any] | None = None
    if data_file is not None:
        try:
            dataset, dataset_metadata = load_dataset(data_file)
            st.success(f"{len(dataset):,} filas · {len(dataset.columns)} columnas")
        except Exception as exc:
            st.error(f"No se pudo cargar el dataset: {exc}")

    st.divider()
    st.caption("Plantillas de protocolo")
    sample_protocol = {
        "tipo_diseno": "medias_independientes",
        "error_alfa": 0.05,
        "poder_estadistico": 0.80,
        "efecto_esperado": 5.0,
        "variabilidad_estimada": 10.0,
        "perdidas_esperadas": 0.10,
    }
    st.download_button(
        "Descargar JSON de ejemplo",
        data=json.dumps(sample_protocol, indent=2, ensure_ascii=False).encode("utf-8"),
        file_name="protocolo_ejemplo.json",
        mime="application/json",
        use_container_width=True,
    )
    st.download_button(
        "Descargar YAML de ejemplo",
        data=yaml.safe_dump(sample_protocol, allow_unicode=True, sort_keys=False).encode("utf-8"),
        file_name="protocolo_ejemplo.yaml",
        mime="application/x-yaml",
        use_container_width=True,
    )
    try:
        st.download_button(
            "Descargar Word de ejemplo",
            data=build_protocol_docx_template(sample_protocol),
            file_name="protocolo_ejemplo.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
        st.download_button(
            "Descargar PDF de ejemplo",
            data=build_protocol_pdf_template(sample_protocol),
            file_name="protocolo_ejemplo.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
    except Exception as exc:
        st.caption(f"No se pudieron generar las plantillas documentales: {exc}")

st.markdown(
    """
    <div class="hero">
      <h1>📐 BioSize Clinical</h1>
      <p>Cálculo trazable de tamaño de muestra para investigación clínica, con curva de potencia, ajustes metodológicos y reporte Word editable.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# Selector de diseño y parámetros generales
left, right = st.columns([1.4, 1.0], gap="large")
with left:
    design_code = st.selectbox(
        "Diseño del estudio",
        options=list(DESIGNS.keys()),
        format_func=lambda x: DESIGNS[x],
        key="design_code",
    )
with right:
    no_variability = st.checkbox("No dispongo de una estimación confiable de variabilidad", key="no_variability")

if no_variability and design_code in {
    "mean_estimation",
    "means_independent",
    "means_paired",
    "superiority_continuous",
    "noninferiority_continuous",
    "equivalence_continuous",
}:
    st.warning(
        "Se recomienda realizar primero un estudio piloto de aproximadamente 30–50 sujetos para estimar la desviación estándar o la variabilidad de las diferencias."
    )

with st.expander("⚙️ Parámetros estadísticos generales", expanded=True):
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        alpha = st.number_input("Error alfa", min_value=0.001, max_value=0.20, step=0.005, format="%.3f", key="alpha")
    with c2:
        power = st.number_input("Poder estadístico (0–1)", min_value=0.50, max_value=0.99, step=0.01, format="%.2f", key="power")
        if power < 0.80:
            st.caption("⚠️ Inferior al mínimo convencional de 80%.")
        elif math.isclose(power, 0.80, abs_tol=1e-9):
            st.caption("80%: mínimo aceptado habitualmente.")
    with c3:
        sided = st.radio("Tipo de contraste", ["Bilateral", "Unilateral"], horizontal=True, key="sided")
    with c4:
        loss_percent = st.slider("Pérdidas esperadas (%)", 0, 50, step=1, key="loss_percent")
        loss_rate = loss_percent / 100.0

# Parámetros específicos
st.subheader("Parámetros del diseño")
params: dict[str, Any] = {
    "design_code": design_code,
    "alpha": alpha,
    "power": power,
    "sided": sided,
    "loss_rate": loss_rate,
}

if design_code == "mean_estimation":
    c1, c2, c3 = st.columns(3)
    with c1:
        params["sd"] = st.number_input("Desviación estándar esperada", min_value=0.0001, key="sd")
    with c2:
        params["precision"] = st.number_input("Error máximo aceptable (±d)", min_value=0.0001, key="precision_mean")
    with c3:
        params["finite_population"] = st.checkbox("Aplicar población finita", key="finite_population")
        if params["finite_population"]:
            params["population"] = st.number_input("Tamaño de la población", min_value=2, step=1, key="population")

elif design_code == "proportion_estimation":
    c1, c2, c3 = st.columns(3)
    with c1:
        params["p"] = st.number_input("Proporción esperada", min_value=0.001, max_value=0.999, step=0.01, key="p")
        st.caption("Use 0,50 cuando sea desconocida para un cálculo conservador.")
    with c2:
        params["precision"] = st.number_input("Error máximo aceptable (±d)", min_value=0.001, max_value=0.50, step=0.005, key="precision_prop")
    with c3:
        params["finite_population"] = st.checkbox("Aplicar población finita", key="finite_population")
        if params["finite_population"]:
            params["population"] = st.number_input("Tamaño de la población", min_value=2, step=1, key="population")

elif design_code in {"means_independent", "superiority_continuous"}:
    c1, c2 = st.columns(2)
    with c1:
        params["delta"] = st.number_input("Diferencia mínima clínicamente importante", min_value=0.0001, key="delta")
    with c2:
        params["sd"] = st.number_input("Desviación estándar común", min_value=0.0001, key="sd")

elif design_code == "means_paired":
    c1, c2 = st.columns(2)
    with c1:
        params["delta"] = st.number_input("Diferencia media esperada", min_value=0.0001, key="delta")
    with c2:
        params["sd_diff"] = st.number_input("DE de las diferencias intraindividuo", min_value=0.0001, key="sd_diff")

elif design_code in {"proportions_independent", "superiority_binary"}:
    c1, c2, c3 = st.columns(3)
    with c1:
        params["p1"] = st.number_input("Proporción grupo 1", min_value=0.001, max_value=0.999, step=0.01, key="p1")
    with c2:
        params["p2"] = st.number_input("Proporción grupo 2", min_value=0.001, max_value=0.999, step=0.01, key="p2")
    with c3:
        params["yates"] = st.checkbox("Corrección de continuidad", key="yates")

elif design_code == "mcnemar":
    c1, c2, c3 = st.columns(3)
    with c1:
        params["p01"] = st.number_input("Discordantes 0→1", min_value=0.001, max_value=0.999, step=0.01, key="p01")
    with c2:
        params["p10"] = st.number_input("Discordantes 1→0", min_value=0.001, max_value=0.999, step=0.01, key="p10")
    with c3:
        params["yates"] = st.checkbox("Corrección de continuidad", key="yates")

elif design_code in {"noninferiority_continuous", "equivalence_continuous"}:
    c1, c2, c3 = st.columns(3)
    with c1:
        params["sd"] = st.number_input("Desviación estándar común", min_value=0.0001, key="sd")
    with c2:
        params["margin"] = st.number_input("Margen clínico", min_value=0.0001, key="margin_cont")
    with c3:
        params["expected_diff"] = st.number_input("Diferencia esperada T−C", key="expected_diff")
    params["sided"] = "Unilateral"

elif design_code == "noninferiority_binary":
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        params["p1"] = st.number_input("Proporción tratamiento", min_value=0.001, max_value=0.999, step=0.01, key="p1_ni")
    with c2:
        params["p2"] = st.number_input("Proporción control", min_value=0.001, max_value=0.999, step=0.01, key="p2_ni")
    with c3:
        params["margin"] = st.number_input("Margen absoluto", min_value=0.001, max_value=0.50, step=0.01, key="margin_bin")
    with c4:
        params["yates"] = st.checkbox("Corrección de continuidad", key="yates")
    params["sided"] = "Unilateral"

elif design_code == "equivalence_binary":
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        params["p1"] = st.number_input("Proporción tratamiento", min_value=0.001, max_value=0.999, step=0.01, key="p1_eq")
    with c2:
        params["p2"] = st.number_input("Proporción control", min_value=0.001, max_value=0.999, step=0.01, key="p2_eq")
    with c3:
        params["margin"] = st.number_input("Margen absoluto", min_value=0.001, max_value=0.50, step=0.01, key="margin_bin")
    with c4:
        params["yates"] = st.checkbox("Corrección de continuidad", key="yates")
    params["sided"] = "Unilateral"

elif design_code in {"odds_ratio", "risk_ratio"}:
    c1, c2, c3 = st.columns(3)
    with c1:
        params["p_control"] = st.number_input("Proporción/riesgo basal en control", min_value=0.001, max_value=0.999, step=0.01, key="p_control")
    with c2:
        label = "Odds Ratio mínimo" if design_code == "odds_ratio" else "Riesgo Relativo mínimo"
        params["effect_ratio"] = st.number_input(label, min_value=0.01, step=0.05, key="effect_ratio")
    with c3:
        params["yates"] = st.checkbox("Corrección de continuidad", key="yates")

elif design_code == "diagnostic_accuracy":
    c1, c2, c3 = st.columns(3)
    with c1:
        params["sensitivity"] = st.number_input("Sensibilidad esperada", min_value=0.001, max_value=0.999, step=0.01, key="sensitivity")
        params["precision_se"] = st.number_input("Precisión para sensibilidad (±d)", min_value=0.001, max_value=0.30, step=0.005, key="precision_se")
    with c2:
        params["specificity"] = st.number_input("Especificidad esperada", min_value=0.001, max_value=0.999, step=0.01, key="specificity")
        params["precision_sp"] = st.number_input("Precisión para especificidad (±d)", min_value=0.001, max_value=0.30, step=0.005, key="precision_sp")
    with c3:
        params["prevalence"] = st.number_input("Prevalencia de la enfermedad", min_value=0.001, max_value=0.999, step=0.01, key="prevalence")
        st.caption("Se calcularán NSe y NSp por separado y se recomendará el mayor.")

elif design_code == "survival":
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        params["hazard_ratio"] = st.number_input("Hazard Ratio objetivo", min_value=0.01, step=0.05, key="hazard_ratio")
    with c2:
        params["mortality_control"] = st.number_input("Mortalidad acumulada control", min_value=0.001, max_value=0.999, step=0.01, key="mortality_control")
    with c3:
        params["mortality_treatment"] = st.number_input("Mortalidad acumulada tratamiento", min_value=0.001, max_value=0.999, step=0.01, key="mortality_treatment")
    with c4:
        params["allocation_treatment"] = st.number_input("Fracción asignada a tratamiento", min_value=0.05, max_value=0.95, step=0.05, key="allocation_treatment")

st.markdown(
    """
    <div class="technical-note"><b>Nota clínica:</b> la diferencia mínima clínicamente importante debe definirse antes del análisis. Un efecto puede ser estadísticamente significativo y, aun así, ser clínicamente irrelevante.</div>
    """,
    unsafe_allow_html=True,
)

# Asistente de estimación desde dataset
if dataset is not None:
    with st.expander("🧪 Estimar parámetros desde el dataset", expanded=False):
        st.dataframe(dataset.head(100), use_container_width=True, height=260)
        cols = numeric_columns(dataset)
        if cols:
            ac1, ac2 = st.columns(2)
            with ac1:
                selected_col = st.selectbox("Variable numérica", cols, key="dataset_numeric_col")
                clean = pd.to_numeric(dataset[selected_col], errors="coerce").dropna()
                if len(clean):
                    st.write(
                        {
                            "n": int(clean.size),
                            "media": round(float(clean.mean()), 4),
                            "DE": round(float(clean.std(ddof=1)), 4) if clean.size > 1 else None,
                            "mínimo": round(float(clean.min()), 4),
                            "máximo": round(float(clean.max()), 4),
                        }
                    )
                    if st.button("Usar DE en el cálculo", disabled=clean.size < 2):
                        st.session_state.sd = float(clean.std(ddof=1))
                        st.session_state.sd_diff = float(clean.std(ddof=1))
                        st.rerun()
            with ac2:
                binary_col = st.selectbox("Variable para proporción/prevalencia", dataset.columns, key="dataset_binary_col")
                values = dataset[binary_col].dropna()
                unique_values = list(values.unique())[:50]
                if unique_values:
                    positive_value = st.selectbox("Valor considerado positivo", unique_values, key="positive_value")
                    estimated_p = float((values == positive_value).mean())
                    st.metric("Proporción estimada", f"{estimated_p:.3f}")
                    if st.button("Usar como proporción y prevalencia"):
                        bounded = min(max(estimated_p, 0.001), 0.999)
                        st.session_state.p = bounded
                        st.session_state.prevalence = bounded
                        st.session_state.p_control = bounded
                        st.rerun()
        else:
            st.info("El dataset no contiene columnas numéricas detectables.")

# Cálculo y dashboard
try:
    result = calculate_sample_size(params)
except Exception as exc:
    st.error(f"No se pudo completar el cálculo: {exc}")
    st.stop()

st.subheader("Dashboard de resumen")
metric_cols = st.columns(4)
metric_cols[0].markdown(
    f'<div class="metric-card"><div class="metric-label">N total final</div><div class="metric-value">{result.n_final_total:,}</div><div class="metric-note">Incluye {loss_rate:.0%} de pérdidas</div></div>',
    unsafe_allow_html=True,
)
metric_cols[1].markdown(
    f'<div class="metric-card"><div class="metric-label">N inicial</div><div class="metric-value">{math.ceil(result.n_raw_total):,}</div><div class="metric-note">Antes del ajuste por pérdidas</div></div>',
    unsafe_allow_html=True,
)
first_group = next(iter(result.groups_final.items())) if result.groups_final else ("Grupo", result.n_final_total)
metric_cols[2].markdown(
    f'<div class="metric-card"><div class="metric-label">{first_group[0]}</div><div class="metric-value">{first_group[1]:,}</div><div class="metric-note">Asignación calculada</div></div>',
    unsafe_allow_html=True,
)
if len(result.groups_final) > 1:
    second_group = list(result.groups_final.items())[1]
    metric_cols[3].markdown(
        f'<div class="metric-card"><div class="metric-label">{second_group[0]}</div><div class="metric-value">{second_group[1]:,}</div><div class="metric-note">Asignación calculada</div></div>',
        unsafe_allow_html=True,
    )
else:
    metric_cols[3].markdown(
        f'<div class="metric-card"><div class="metric-label">Poder / criterio</div><div class="metric-value">{power:.0%}</div><div class="metric-note">{"Usado" if result.uses_power else "No aplicable: diseño por precisión"}</div></div>',
        unsafe_allow_html=True,
    )

if result.extra.get("required_events"):
    st.info(f"Eventos requeridos para supervivencia: **{result.extra['required_events']:,}**.")
if design_code == "diagnostic_accuracy":
    st.info(
        f"Requerimiento por sensibilidad: **{result.extra['n_sensitivity']:,}** · por especificidad: **{result.extra['n_specificity']:,}** · se adopta el mayor."
    )
for warning in result.warnings:
    st.warning(warning)

# Curva de potencia o precisión
plot_col, detail_col = st.columns([1.55, 1.0], gap="large")
with plot_col:
    if result.uses_power:
        powers = np.round(np.linspace(0.50, 0.95, 19), 3)
        n_values: list[int] = []
        for pwr in powers:
            try:
                n_values.append(calculate_sample_size(params, power_override=float(pwr)).n_final_total)
            except Exception:
                n_values.append(np.nan)
        fig = go.Figure()
        fig.add_trace(
            go.Scatter(
                x=powers * 100,
                y=n_values,
                mode="lines+markers",
                name="N total",
                hovertemplate="Poder %{x:.0f}%<br>N %{y:,.0f}<extra></extra>",
            )
        )
        fig.add_vline(x=power * 100, line_dash="dash", annotation_text=f"Seleccionado: {power:.0%}")
        fig.update_layout(
            title="Curva de potencia: tamaño de muestra vs. poder",
            xaxis_title="Poder estadístico (%)",
            yaxis_title="N total final",
            height=430,
            margin=dict(l=20, r=20, t=60, b=20),
            hovermode="x unified",
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        if design_code == "diagnostic_accuracy":
            base_precision = min(params["precision_se"], params["precision_sp"])
            precision_values = np.linspace(max(0.01, base_precision * 0.55), min(0.25, base_precision * 1.60), 20)
            n_values = []
            for d in precision_values:
                pcopy = dict(params)
                pcopy["precision_se"] = float(d)
                pcopy["precision_sp"] = float(d)
                n_values.append(calculate_sample_size(pcopy).n_final_total)
        else:
            base_precision = params["precision"]
            precision_values = np.linspace(max(0.001, base_precision * 0.55), base_precision * 1.60, 20)
            n_values = []
            for d in precision_values:
                pcopy = dict(params)
                pcopy["precision"] = float(d)
                n_values.append(calculate_sample_size(pcopy).n_final_total)
        fig = go.Figure()
        fig.add_trace(
            go.Scatter(
                x=precision_values,
                y=n_values,
                mode="lines+markers",
                name="N total",
                hovertemplate="Precisión ±%{x:.3f}<br>N %{y:,.0f}<extra></extra>",
            )
        )
        fig.add_vline(x=base_precision, line_dash="dash", annotation_text="Precisión seleccionada")
        fig.update_layout(
            title="Curva de precisión: tamaño de muestra vs. error máximo",
            xaxis_title="Error máximo aceptable (±d)",
            yaxis_title="N total final",
            height=430,
            margin=dict(l=20, r=20, t=60, b=20),
        )
        st.plotly_chart(fig, use_container_width=True)

with detail_col:
    st.markdown("#### Justificación matemática")
    st.code(result.formula, language=None)
    st.write(result.method)
    st.markdown("**Supuestos:**")
    st.write(" · ".join(result.assumptions))
    group_df = pd.DataFrame(
        [
            {
                "Componente": name,
                "N sin pérdidas": math.ceil(result.groups_raw.get(name, 0)),
                "N final": value,
            }
            for name, value in result.groups_final.items()
        ]
    )
    st.dataframe(group_df, hide_index=True, use_container_width=True)

# Análisis de sensibilidad y descargas
if result.uses_power:
    scenarios = sorted(set([power, 0.80, 0.90, 0.95]))
    sensitivity_rows = []
    for scenario_power in scenarios:
        scenario_result = calculate_sample_size(params, power_override=scenario_power)
        sensitivity_rows.append(
            {
                "Escenario": f"Poder {scenario_power:.0%}",
                "Poder": f"{scenario_power:.0%}",
                "N total": scenario_result.n_final_total,
                "Variación vs actual": scenario_result.n_final_total - result.n_final_total,
            }
        )
else:
    sensitivity_rows = [
        {
            "Escenario": "Diseño basado en precisión",
            "Poder": "No aplica",
            "N total": result.n_final_total,
            "Variación vs actual": 0,
        },
        {
            "Escenario": "Poder 90%",
            "Poder": "No aplica a esta fórmula",
            "N total": result.n_final_total,
            "Variación vs actual": 0,
        },
    ]
sensitivity_df = pd.DataFrame(sensitivity_rows)

st.subheader("Análisis de sensibilidad")
st.dataframe(sensitivity_df, hide_index=True, use_container_width=True)

word_report = build_word_report(result, params, sensitivity_df, dataset_metadata)
consort_docx = build_consort_docx()
consort_pdf = build_consort_pdf()

st.subheader("Descargas")
d1, d2, d3 = st.columns(3)
with d1:
    st.download_button(
        "⬇️ Informe Word editable",
        data=word_report,
        file_name=f"informe_tamano_muestra_{design_code}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )
with d2:
    st.download_button(
        "⬇️ Checklist CONSORT 2010 · Word",
        data=consort_docx,
        file_name="checklist_CONSORT_2010_operativo.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )
with d3:
    st.download_button(
        "⬇️ Checklist CONSORT 2010 · PDF",
        data=consort_pdf,
        file_name="checklist_CONSORT_2010_operativo.pdf",
        mime="application/pdf",
        use_container_width=True,
    )

with st.expander("Texto CONSORT generado", expanded=False):
    st.write(consort_paragraph(result, params))

st.divider()
st.caption(
    "Herramienta de planificación metodológica. Verifique los supuestos, la fuente del efecto esperado y la adecuación del método al protocolo. "
    "Los estudios regulatorios o de diseño complejo requieren revisión bioestadística independiente."
)
