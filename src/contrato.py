"""Contrato de entrada e saída do JMSFicha, igual ao do AutoFicha (chamado pelo AutoArquivo).

Entrada (linha de comando)::

    python main.py --servidores "<dir>" --registros "<json>" [--pasta "<dir>"] [--retorno "<arquivo.json>"]

``--registros``: ``[{"matricula":"1234","inicio":2007,"fim":2008}, ...]`` (anos).

Saída: em ``--retorno`` o script grava (de forma atômica) o JSON lido pelo painel::

    {"sucesso": true, "pastas_processadas": ["C:/.../Servidor A"],
     "total_pastas": 1, "total_fichas_salvas": 2, "falhas": [], "erro": ""}

Sem dependências do Windows: pode ser testado em qualquer sistema.
"""

from __future__ import annotations

import argparse
import json
import os
from pathlib import Path


def adicionar_argumentos(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--servidores", help="Pasta Servidores (cada servidor ganha uma subpasta).")
    parser.add_argument("--registros", help="JSON com matrícula, ano inicial e ano final.")
    parser.add_argument("--pasta", help="Pasta específica: os PDFs são salvos direto nela.")
    parser.add_argument("--retorno", help="Arquivo JSON onde o resultado é gravado.")


def pasta_existente(parser: argparse.ArgumentParser, valor: str | None, nome: str) -> Path | None:
    if not valor:
        return None
    pasta = Path(valor)
    if not pasta.is_dir():
        parser.error(f"{nome}: pasta não encontrada: {pasta}")
    return pasta.resolve()


def registros_para_dados(texto: str, parser: argparse.ArgumentParser) -> list[tuple[str, list[int]]]:
    """Converte o JSON de ``--registros`` no formato interno ``[(matricula, [anos...]), ...]``.

    Registros da mesma matrícula são somados sem repetir anos; a ordem das matrículas é mantida.
    """
    try:
        itens = json.loads(texto)
    except ValueError as exc:
        parser.error(f"--registros não é um JSON válido: {exc}")
    if not isinstance(itens, list) or not itens:
        parser.error("--registros deve ser uma lista não vazia.")

    anos_por_matricula: dict[str, set[int]] = {}
    for n, item in enumerate(itens, start=1):
        ok = isinstance(item, dict)
        mat = str(item.get("matricula", "")).strip() if ok else ""
        ini, fim = (item.get("inicio"), item.get("fim")) if ok else (None, None)
        if (not mat.isdigit() or isinstance(ini, bool) or isinstance(fim, bool)
                or not isinstance(ini, int) or not isinstance(fim, int) or ini > fim):
            parser.error(f"--registros: item {n} inválido ({item!r}).")
        anos_por_matricula.setdefault(mat, set()).update(range(ini, fim + 1))
    return [(mat, sorted(anos)) for mat, anos in anos_por_matricula.items()]


class Retorno:
    """Acumula o resultado da execução e grava o JSON de ``--retorno`` (se informado)."""

    def __init__(self, caminho: str | os.PathLike | None) -> None:
        self.caminho = Path(caminho) if caminho else None
        self.pastas: list[str] = []
        self.total_fichas = 0
        self.falhas: list[dict] = []

    def ficha_salva(self, pasta: Path) -> None:
        texto = str(Path(pasta).resolve())  # o painel exige caminho absoluto
        if all(os.path.normcase(texto) != os.path.normcase(p) for p in self.pastas):
            self.pastas.append(texto)
        self.total_fichas += 1

    def falha(self, mensagem: str) -> None:
        self.falhas.append({"mensagem": mensagem})

    def gravar(self, sucesso: bool, erro: str = "") -> None:
        if self.caminho is None:
            return
        dados = {
            "sucesso": bool(sucesso),
            "pastas_processadas": self.pastas,
            "total_pastas": len(self.pastas),
            "total_fichas_salvas": self.total_fichas,
            "falhas": self.falhas,
            "erro": erro,
        }
        tmp = self.caminho.with_name(self.caminho.name + ".tmp")
        tmp.write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(tmp, self.caminho)  # troca atômica: o painel nunca lê um arquivo pela metade
