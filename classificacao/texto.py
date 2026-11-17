"""Normalização de texto usada em todo o projeto."""

import re
import unicodedata

import pandas as pd


def sem_acento(texto: str) -> str:
    return unicodedata.normalize("NFKD", texto).encode("ASCII", errors="ignore").decode("ASCII")


def limpa_texto_procv(texto):
    """
    Limpeza ORIGINAL do projeto (remove acentos, números e caracteres
    especiais). Mantida igual para continuar batendo com a base de
    referência (Transacoes_base.xlsx) no PROCV.
    """
    if pd.isna(texto):
        return texto
    texto = sem_acento(str(texto))
    texto = re.sub(r"[0-9\-]", " ", texto)
    texto = re.sub(r"[^A-Za-z\s]", " ", texto)
    texto = re.sub(r"\s+", " ", texto).strip()
    texto = re.sub(r"\bS\s+NF\b", "", texto)
    return re.sub(r"\s+", " ", texto).strip().upper()


def limpa_texto_regra(texto) -> str:
    """
    Limpeza para as regras de palavra-chave: tira acentos e pontuação, mas
    MANTÉM os números (para regras como "ARLA 32" ou "DIESEL S10").
    """
    if texto is None or (not isinstance(texto, str) and pd.isna(texto)):
        return ""
    texto = sem_acento(str(texto)).upper()
    texto = re.sub(r"[^A-Z0-9]", " ", texto)
    return re.sub(r"\s+", " ", texto).strip()


def normaliza_chave(texto) -> str:
    """Chave para comparar nomes de conta/categoria/grupo: sem acento, minúsculo, espaços compactados."""
    if texto is None or (not isinstance(texto, str) and pd.isna(texto)):
        return ""
    return re.sub(r"\s+", " ", sem_acento(str(texto))).strip().lower()
