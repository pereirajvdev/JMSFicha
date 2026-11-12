"""JMSFicha: gera fichas financeiras (2008 e anteriores) no sistema Folha.

Uso::

    python main.py --servidores "<dir>" --registros "<json>" [--pasta "<dir>"] [--retorno "<arq.json>"]
    python main.py planilha.xlsx
    python main.py --mat MATRICULA ANO_INICIAL-ANO_FINAL
    (qualquer forma aceita --perf para o relatório de desempenho)

Códigos de saída: 0 sucesso | 1 falha | 2 argumentos inválidos | 130 interrompido.
"""

from __future__ import annotations

import argparse
import logging
import re
import sys
import zipfile
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

import perf
from config import (
    CAMINHO_TEMPORARIO,
    CRED,
    EXE,
    MAX_TENTATIVAS,
    PASTA_DOWNLOADS,
    PASTA_SERVIDORES_PADRAO,
    ErroConfiguracao,
    ErroJMSFicha,
)
from contrato import Retorno, adicionar_argumentos, pasta_existente, registros_para_dados
from pdf_ficha import Arquivador, deletar_temporario

log = logging.getLogger("jmsficha")


# ======================================================================
# LOG
# ======================================================================
def configurar_log() -> None:
    """INFO em stdout; ERROR em stderr. Nunca registra credenciais."""
    for fluxo in (sys.stdout, sys.stderr):
        try:
            fluxo.reconfigure(errors="replace")
        except AttributeError:
            pass
    formato = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S")
    saida = logging.StreamHandler(sys.stdout)
    saida.setLevel(logging.INFO)
    saida.addFilter(lambda r: r.levelno < logging.ERROR)
    erro = logging.StreamHandler(sys.stderr)
    erro.setLevel(logging.ERROR)
    for h in (saida, erro):
        h.setFormatter(formato)
    raiz = logging.getLogger()
    raiz.handlers[:] = [saida, erro]
    raiz.setLevel(logging.INFO)


# ======================================================================
# ENTRADA
# ======================================================================
def _ano_de(valor) -> int | None:
    if isinstance(valor, (datetime, date)):
        return valor.year
    if isinstance(valor, int) and not isinstance(valor, bool) and 1900 <= valor <= 2100:
        return valor
    if isinstance(valor, str):
        achado = re.search(r"\d{4}", valor)
        if achado:
            return int(achado.group())
    return None


def _matricula_de(valor) -> str | None:
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)  # Excel entrega 1234 como 1234.0
    texto = str(valor).strip() if valor is not None else ""
    return texto if texto.isdigit() else None


def ler_planilha(caminho: Path) -> tuple[list[tuple[str, list[int]]], list[str]]:
    """Lê (matrícula, data inicial, data final). Devolve (dados, avisos).

    Anos de linhas repetidas da mesma matrícula são unidos; linhas inválidas viram aviso.
    """
    try:
        wb = load_workbook(caminho, data_only=True, read_only=True)
    except (OSError, zipfile.BadZipFile, InvalidFileException, KeyError) as exc:
        raise ErroConfiguracao(f"Não foi possível abrir a planilha {caminho}: {exc}") from exc

    anos_por_matricula: dict[str, set[int]] = {}
    avisos: list[str] = []
    try:
        for numero, linha in enumerate(wb.active.iter_rows(min_row=2, values_only=True), start=2):
            mat_bruta, ini_bruto, fim_bruto = (list(linha) + [None] * 3)[:3]
            if mat_bruta is None and ini_bruto is None and fim_bruto is None:
                continue
            matricula, inicio, fim = _matricula_de(mat_bruta), _ano_de(ini_bruto), _ano_de(fim_bruto)
            if matricula is None or inicio is None or fim is None or inicio > fim:
                avisos.append(f"Linha {numero} ignorada: {mat_bruta!r}, {ini_bruto!r}, {fim_bruto!r}")
                continue
            anos_por_matricula.setdefault(matricula, set()).update(range(inicio, fim + 1))
    finally:
        wb.close()

    return [(m, sorted(a)) for m, a in anos_por_matricula.items()], avisos


def obter_dados(args, parser) -> list[tuple[str, list[int]]]:
    """Devolve [(matricula, [anos...]), ...] a partir de --registros, --mat ou da planilha."""
    if args.registros is not None:
        if args.mat or args.planilha:
            parser.error("Use --registros, --mat ou a planilha, apenas um deles.")
        return registros_para_dados(args.registros, parser)

    if args.mat:
        if args.planilha:
            parser.error("Use --mat ou o caminho da planilha, não os dois.")
        matricula, periodo = args.mat
        if not matricula.isdigit():
            parser.error("A matrícula deve conter apenas dígitos.")
        resultado = re.fullmatch(r"(\d{4})-(\d{4})", periodo)
        if not resultado:
            parser.error(
                "Período inválido. Use o formato ANO_INICIAL-ANO_FINAL, por exemplo: 2000-2005."
            )
        inicial, final = int(resultado.group(1)), int(resultado.group(2))
        if inicial > final:
            parser.error("O ano inicial não pode ser maior que o ano final.")
        return [(matricula, list(range(inicial, final + 1)))]

    if args.planilha:
        planilha = Path(args.planilha)
        if not planilha.is_file():
            parser.error(f"Planilha não encontrada: {planilha}")
        try:
            dados, avisos = ler_planilha(planilha)
        except ErroConfiguracao as exc:
            parser.error(str(exc))
        for aviso in avisos:
            log.warning(aviso)
        return dados

    parser.error("Informe --registros, uma planilha ou --mat MATRICULA ANO_INICIAL-ANO_FINAL.")


def criar_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Gera fichas financeiras (2008 e anteriores) por registros, planilha ou matrícula."
    )
    parser.add_argument("planilha", nargs="?", help="Planilha Excel com as matrículas e os períodos.")
    parser.add_argument(
        "--mat", nargs=2, metavar=("MATRICULA", "PERIODO"),
        help="Matrícula e intervalo de anos. Ex.: --mat 1234 2000-2005",
    )
    parser.add_argument("--perf", action="store_true", help="Imprime o relatório de desempenho ao final.")
    adicionar_argumentos(parser)
    return parser


# ======================================================================
# PRÉ-REQUISITOS
# ======================================================================
def verificar_pre_requisitos(arquivador: Arquivador) -> None:
    """Falha cedo, antes de tocar na interface, se algo essencial faltar."""
    if not EXE.is_file():
        raise ErroConfiguracao(f"Executável do Folha não encontrado: {EXE}")
    if not CRED.is_file():
        raise ErroConfiguracao(f"Arquivo de credenciais não encontrado: {CRED}")
    try:
        PASTA_DOWNLOADS.mkdir(parents=True, exist_ok=True)
        if arquivador.pasta_fixa is None:
            arquivador.pasta_servidores.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise ErroConfiguracao(f"Não foi possível preparar as pastas de trabalho: {exc}") from exc


# ======================================================================
# EXECUÇÃO
# ======================================================================
@dataclass
class _Estado:
    reposicionar: bool = False        # há ficha anterior: o formulário precisa de TABs
    ultima_matricula: str | None = None


def _processar_ficha(matricula: str, ano: int, estado: _Estado, arquivador: Arquivador) -> Path:
    import automacao  # Windows-only: importado só quando há de fato automação

    if estado.reposicionar:
        automacao.reposicionar_formulario(mudou_matricula=matricula != estado.ultima_matricula)
        estado.reposicionar = False
    automacao.gerar_e_salvar_pdf(matricula, ano)
    destino = arquivador.arquivar(CAMINHO_TEMPORARIO, matricula, ano)
    estado.reposicionar = True
    estado.ultima_matricula = matricula
    return destino


def executar(dados, arquivador: Arquivador, retorno: Retorno) -> tuple[bool, str]:
    """Processa todos os registros. Uma ficha com falha não derruba as demais."""
    import automacao
    from janelas import EXC_AUTOMACAO

    erros_previstos = (ErroJMSFicha, *EXC_AUTOMACAO)
    total = sum(len(anos) for _, anos in dados)
    log.info("%d matrícula(s), %d ficha(s) a gerar.", len(dados), total)

    automacao.configurar_pyautogui()
    estado = _Estado()
    feitas = 0
    try:
        try:
            automacao.garantir_impressora_pdf()
            automacao.iniciar_sessao()
        except erros_previstos as exc:
            retorno.falha(str(exc))
            return False, str(exc)

        for im, (matricula, anos) in enumerate(dados, start=1):
            for ia, ano in enumerate(anos, start=1):
                log.info("=" * 70)
                log.info("Matrícula %s (%d/%d) | ano %d (%d/%d)", matricula, im, len(dados), ano, ia, len(anos))

                ultimo_erro = ""
                for tentativa in range(1, MAX_TENTATIVAS + 1):
                    try:
                        destino = _processar_ficha(matricula, ano, estado, arquivador)
                    except erros_previstos as exc:
                        ultimo_erro = str(exc)
                        log.error("Matrícula %s, ano %d, tentativa %d/%d: %s",
                                  matricula, ano, tentativa, MAX_TENTATIVAS, exc)
                        # Estado da tela desconhecido: recomeça a sessão do zero.
                        estado = _Estado()
                        try:
                            automacao.iniciar_sessao()
                        except erros_previstos as exc2:
                            msg = f"Sessão não pôde ser reiniciada após falha (matrícula {matricula}, ano {ano}): {exc2}"
                            retorno.falha(msg)
                            return False, msg
                        continue
                    retorno.ficha_salva(destino.parent)
                    feitas += 1
                    break
                else:
                    retorno.falha(f"Matrícula {matricula}, ano {ano}: {ultimo_erro}")
    finally:
        _limpar(automacao)

    falhas = len(retorno.falhas)
    log.info("Concluído: %d ficha(s) salva(s), %d falha(s).", feitas, falhas)
    return falhas == 0, ("" if falhas == 0 else f"{falhas} ficha(s) não geradas; veja 'falhas'.")


def _limpar(automacao) -> None:
    """Encerra o Folha e remove o temporário; falhas aqui só são registradas."""
    try:
        automacao.encerrar_folha()
    except ErroJMSFicha as exc:
        log.warning("Não foi possível encerrar o Folha: %s", exc)
    try:
        deletar_temporario(CAMINHO_TEMPORARIO)
    except ErroJMSFicha as exc:
        log.warning("%s", exc)


def main(argv: list[str] | None = None) -> int:
    configurar_log()
    parser = criar_parser()
    args = parser.parse_args(argv)  # entrada inválida: sai com código 2

    servidores = pasta_existente(parser, args.servidores, "--servidores")
    pasta = pasta_existente(parser, args.pasta, "--pasta")
    dados = obter_dados(args, parser)
    retorno = Retorno(args.retorno)

    if not dados:
        log.error("Nenhum registro encontrado.")
        retorno.falha("Nenhum registro encontrado.")
        retorno.gravar(False, "Nenhum registro encontrado.")
        return 1

    perf.ativar(args.perf)
    arquivador = Arquivador(servidores or PASTA_SERVIDORES_PADRAO, pasta)
    codigo = 1
    sucesso, erro = False, ""
    try:
        verificar_pre_requisitos(arquivador)
        sucesso, erro = executar(dados, arquivador, retorno)
        codigo = 0 if sucesso else 1
    except KeyboardInterrupt:
        erro, codigo = "Interrompido pelo usuário.", 130
        retorno.falha(erro)
    except ErroJMSFicha as exc:
        erro = str(exc)
        retorno.falha(erro)
    except Exception as exc:  # último recurso: sempre gravar o retorno com o motivo
        if type(exc).__name__ == "FailSafeException":  # pyautogui: mouse no canto da tela
            erro, codigo = "Interrompido pelo failsafe do pyautogui.", 130
        else:
            log.exception("Erro inesperado")
            erro = f"Erro inesperado: {exc}"
        retorno.falha(erro)
    finally:
        retorno.gravar(sucesso, erro)
        perf.imprimir_relatorio_tempos()

    if erro:
        log.error("%s", erro)
    return codigo


if __name__ == "__main__":
    sys.exit(main())
