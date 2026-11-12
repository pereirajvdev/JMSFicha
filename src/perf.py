"""Cronometragem opcional (--perf) com tempo total e tempo próprio por etapa.

Etapas aninhadas se sobrepõem no tempo total; por isso o relatório mostra
também o tempo PRÓPRIO (total menos o tempo das etapas filhas). A soma da
coluna "Próprio" é o tempo medido de verdade, sem dupla contagem, e a ordem
do relatório é por tempo próprio: o primeiro item é o maior gargalo.
Desligado, o custo é uma verificação de atributo por chamada.
"""

from __future__ import annotations

import time
from contextlib import contextmanager
from functools import wraps


class _Perfil:
    def __init__(self) -> None:
        self.ativo = False
        self._dados: dict[str, list[float]] = {}  # nome -> [n, total, proprio, max]
        self._pilha: list[list[float]] = []

    def ativar(self, ativo: bool = True) -> None:
        self.ativo = ativo

    def limpar(self) -> None:
        self._dados.clear()
        self._pilha.clear()

    @contextmanager
    def medir(self, nome: str):
        if not self.ativo:
            yield
            return
        filhos = [0.0]
        self._pilha.append(filhos)
        inicio = time.perf_counter()
        try:
            yield
        finally:
            total = time.perf_counter() - inicio
            self._pilha.pop()
            if self._pilha:
                self._pilha[-1][0] += total
            reg = self._dados.setdefault(nome, [0, 0.0, 0.0, 0.0])
            reg[0] += 1
            reg[1] += total
            reg[2] += total - filhos[0]
            reg[3] = max(reg[3], total)

    def relatorio(self) -> str:
        if not self.ativo or not self._dados:
            return ""
        soma = sum(r[2] for r in self._dados.values()) or 1.0
        linhas = [
            "=" * 82,
            " RELATÓRIO DE DESEMPENHO (--perf)",
            " Total = inclui etapas filhas | Próprio = só a etapa (soma sem dupla contagem)",
            "=" * 82,
            f"{'Etapa':<34}{'Chamadas':>9}{'Total(s)':>10}{'Próprio(s)':>11}"
            f"{'%':>6}{'Média(s)':>9}{'Máx(s)':>8}",
        ]
        for nome, (n, total, proprio, maximo) in sorted(
            self._dados.items(), key=lambda kv: kv[1][2], reverse=True
        ):
            linhas.append(
                f"{nome:<34}{int(n):>9}{total:>10.2f}{proprio:>11.2f}"
                f"{100 * proprio / soma:>5.0f}%{total / n:>9.3f}{maximo:>8.2f}"
            )
        linhas.append("=" * 82)
        return "\n".join(linhas)


PERFIL = _Perfil()


def ativar(ativo: bool = True) -> None:
    PERFIL.ativar(ativo)


def medir(nome: str):
    """Context manager: mede um bloco (sem efeito com --perf desligado)."""
    return PERFIL.medir(nome)


def cronometrar(nome: str | None = None):
    """Decorator: mede o tempo de uma função (sem efeito com --perf desligado)."""

    def decorador(func):
        rotulo = nome or func.__name__

        @wraps(func)
        def wrapper(*args, **kwargs):
            if not PERFIL.ativo:
                return func(*args, **kwargs)
            with PERFIL.medir(rotulo):
                return func(*args, **kwargs)

        return wrapper

    return decorador


def imprimir_relatorio_tempos() -> None:
    texto = PERFIL.relatorio()
    if texto:
        print("\n" + texto)
