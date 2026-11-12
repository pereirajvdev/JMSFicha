"""Condução da interface do sistema Folha (login, navegação, impressão, salvamento).

Depende de Windows, pywinauto e do sistema Folha instalado: não é testável
fora da máquina de produção. A sequência de teclas do legado foi preservada.
"""

from __future__ import annotations

import logging
import subprocess
import sys
import time
from pathlib import Path

import psutil
import pyautogui
import pyperclip
import pywintypes
import win32print

from config import (
    AUTOFICHA_DIR,
    AUTOFICHA_SCRIPT,
    CAMINHO_TEMPORARIO,
    ESPERA_FORMULARIO,
    ESPERA_REPOSICIONAR,
    EXE,
    IMPRESSORA_PDF,
    INTERVALO_TAB_ANO,
    INTERVALO_TAB_MATRICULA,
    MODO_INTERNO_RAPIDO,
    PYAUTOGUI_PAUSE,
    REPORT_CLIQUE,
    TEMPO_GRACE_MODO,
    TEMPO_JANELA_INTERNA,
    TEMPO_JMS,
    TEMPO_JMS_INICIAL,
    TEMPO_LIMITE,
    TEMPO_REPORT,
    TEMPO_REPORT_OPCIONAL,
    TITULO_EXTERNO_IMPRIMIR,
    TITULO_EXTERNO_PRINT,
    TITULO_IMPRESSAO,
    TITULO_INICIAL,
    TITULO_INTERNO,
    TITULO_JMS,
    TITULO_REPORT,
    TITULO_REPORT_VARIACAO,
    TRECHO_REPORT,
    ErroAutomacao,
    ErroConfiguracao,
    JanelaNaoEncontrada,
    ler_credenciais,
)
from janelas import (
    EXC_AUTOMACAO,
    aguardar_janela,
    exigir_controle,
    exigir_foco,
    exigir_janela,
    hwnds_com_trecho,
    janela_viva,
    listar_janelas_visiveis,
    pid_da_janela,
    procurar_controle,
    propriedades_controle,
    titulo_em_primeiro_plano,
)
from pdf_ficha import aguardar_pdf_pronto, deletar_temporario
from perf import cronometrar

log = logging.getLogger(__name__)

MODO_EXTERNO_IMPRIMIR = "externo_imprimir"
MODO_EXTERNO_PRINT = "externo_print"
MODO_INTERNO = "interno_print"

_TITULOS_FOLHA = (TITULO_INICIAL, TITULO_JMS, TITULO_REPORT, TITULO_REPORT_VARIACAO)


def configurar_pyautogui() -> None:
    """Fixa PAUSE explicitamente (0,1 s era o valor implícito) e mantém o failsafe."""
    pyautogui.PAUSE = PYAUTOGUI_PAUSE
    pyautogui.FAILSAFE = True  # mouse no canto superior esquerdo aborta


# ======================================================================
# IMPRESSORA PDF
# ======================================================================
def _impressora_padrao() -> str:
    try:
        return win32print.GetDefaultPrinter()
    except (RuntimeError, pywintypes.error):
        return ""


def _e_pdf(nome: str) -> bool:
    return nome.strip().casefold() == IMPRESSORA_PDF


def _aguardar_impressora_pdf(timeout: float) -> bool:
    limite = time.monotonic() + timeout
    while True:
        if _e_pdf(_impressora_padrao()):
            return True
        if time.monotonic() >= limite:
            return False
        time.sleep(0.25)


@cronometrar()
def garantir_impressora_pdf() -> None:
    """Garante 'Microsoft Print to PDF' como padrão, chamando o AutoFicha --padrao."""
    atual = _impressora_padrao()
    if _e_pdf(atual):
        log.info("Impressora padrão já é Microsoft Print to PDF.")
        return

    log.info("Impressora padrão atual: %r. Acionando AutoFicha --padrao...", atual)
    if AUTOFICHA_SCRIPT.is_file():
        try:
            subprocess.Popen(  # sem esperar: como na versão anterior; o estado é verificado abaixo
                [sys.executable, str(AUTOFICHA_SCRIPT), "--padrao"],
                cwd=str(AUTOFICHA_DIR),
            )
        except OSError as exc:
            log.warning("Não foi possível executar o AutoFicha: %s", exc)
    else:
        log.warning("AutoFicha não encontrado em %s", AUTOFICHA_SCRIPT)

    if _aguardar_impressora_pdf(15):
        return

    log.warning("AutoFicha não definiu a impressora; tentando diretamente pelo Windows.")
    try:
        win32print.SetDefaultPrinter("Microsoft Print to PDF")
    except pywintypes.error as exc:
        log.warning("SetDefaultPrinter falhou: %s", exc)
    if not _aguardar_impressora_pdf(3):
        raise ErroConfiguracao(
            f"A impressora padrão é {_impressora_padrao()!r}, não Microsoft Print to PDF."
        )


# ======================================================================
# PROCESSO DA FOLHA
# ======================================================================
def _processo_da_folha(proc: psutil.Process) -> bool:
    try:
        if proc.name().casefold() == EXE.name.casefold():
            return True
        return Path(proc.exe()).parent == EXE.parent
    except (psutil.NoSuchProcess, psutil.AccessDenied, OSError):
        return False


def _processos_folha() -> list[psutil.Process]:
    """Processos do Folha: por nome/pasta do executável e donos das janelas conhecidas."""
    achados: dict[int, psutil.Process] = {}
    for proc in psutil.process_iter(["name"]):
        if _processo_da_folha(proc):
            achados[proc.pid] = proc

    titulos = {t.strip().casefold() for t in _TITULOS_FOLHA}
    for hwnd, texto in listar_janelas_visiveis():
        if texto.strip().casefold() not in titulos:
            continue
        pid = pid_da_janela(hwnd)
        if pid in achados:
            continue
        try:
            proc = psutil.Process(pid)
        except psutil.NoSuchProcess:
            continue
        if _processo_da_folha(proc):
            achados[pid] = proc
        else:
            log.warning("Janela '%s' pertence a outro programa (PID %d); não será encerrada.", texto, pid)
    return list(achados.values())


@cronometrar()
def encerrar_folha() -> None:
    """Encerra instâncias anteriores do Folha (só processos do próprio Folha)."""
    for _ in range(3):
        alvos = _processos_folha()
        if not alvos:
            return
        for proc in alvos:
            try:
                log.info("Encerrando %s (PID %d)...", proc.name(), proc.pid)
                proc.terminate()
            except psutil.NoSuchProcess:
                continue
            except psutil.AccessDenied as exc:
                raise ErroAutomacao(
                    "Acesso negado ao processo da folha. Verifique as permissões do Windows."
                ) from exc
        _, vivos = psutil.wait_procs(alvos, timeout=2)
        for proc in vivos:
            try:
                proc.kill()
            except psutil.NoSuchProcess:
                pass
        psutil.wait_procs(vivos, timeout=2)
    if _processos_folha():
        raise ErroAutomacao("Não foi possível encerrar a instância anterior do Folha.")


# ======================================================================
# LOGIN E NAVEGAÇÃO INICIAL
# ======================================================================
def _digitar(texto: str) -> None:
    """Digita no campo em foco; textos não ASCII vão por colar (pyautogui.write ignora-os)."""
    if texto.isascii():
        pyautogui.write(texto, interval=0.001)
        return
    pyperclip.copy(texto)
    try:
        pyautogui.hotkey("ctrl", "v")
    finally:
        pyperclip.copy("")


@cronometrar()
def login() -> None:
    """Abre o Folha e preenche usuário/senha. Levanta erro se não for possível."""
    if not EXE.is_file():
        raise ErroConfiguracao(f"Executável do Folha não encontrado: {EXE}")
    usuario, senha = ler_credenciais()

    encerrar_folha()

    log.info("Abrindo Folha (%s)...", EXE)
    try:
        subprocess.Popen([str(EXE)], cwd=str(EXE.parent))
    except OSError as exc:
        raise ErroAutomacao(f"Não foi possível iniciar {EXE}: {exc}") from exc

    janela = exigir_janela(TITULO_INICIAL)
    exigir_foco(janela)

    _digitar(usuario)
    pyautogui.press("tab")
    _digitar(senha)
    pyautogui.press("tab")
    pyautogui.press("enter")
    log.info("Credenciais enviadas.")


@cronometrar()
def abrir_janela_interna() -> None:
    """Abre a tela de fichas pelo menu (sequência de teclas do legado preservada)."""
    janela = aguardar_janela(TITULO_JMS, time_limit=TEMPO_JMS_INICIAL)
    if janela is None:
        # aviso intermediário com o mesmo título do login: Form1
        aviso = aguardar_janela(TITULO_INICIAL, time_limit=2)
        if aviso is not None:
            exigir_foco(aviso)
            pyautogui.press("enter")
        janela = exigir_janela(TITULO_JMS, time_limit=TEMPO_JMS)
    exigir_foco(janela)

    pyautogui.press("alt")
    pyautogui.press("r")
    pyautogui.press("down", presses=6, interval=0.001)
    pyautogui.press("right")
    pyautogui.press("down", presses=9, interval=0.001)
    pyautogui.press("enter")


@cronometrar()
def verificar_janela_interna(timeout: float = TEMPO_JANELA_INTERNA) -> None:
    """Espera (com polling) o controle 'Form1' dentro da janela principal."""
    janela = exigir_janela(TITULO_JMS)
    limite = time.monotonic() + timeout
    while True:
        if procurar_controle(janela, titulo=TITULO_INTERNO) is not None:
            log.info("Janela interna identificada.")
            return
        if time.monotonic() >= limite:
            raise JanelaNaoEncontrada(f"Janela interna '{TITULO_INTERNO}' não apareceu em {timeout:.0f}s.")
        time.sleep(0.3)
        if not janela_viva(int(janela.handle)):
            janela = exigir_janela(TITULO_JMS)


@cronometrar("iniciar_sessao")
def iniciar_sessao() -> None:
    """Login + abertura + verificação da janela interna (também usado na recuperação)."""
    login()
    abrir_janela_interna()
    verificar_janela_interna()


# ======================================================================
# FORMULÁRIO (janela desenhada: só teclado)
# ======================================================================
def _garantir_foco_formulario() -> None:
    """Se o primeiro plano não for uma janela conhecida do Folha, refocaliza a principal."""
    conhecidos = {t.strip().casefold() for t in (*_TITULOS_FOLHA, TITULO_INTERNO)}
    if titulo_em_primeiro_plano().strip().casefold() in conhecidos:
        return
    log.warning("Primeiro plano inesperado (%r); refocalizando a janela principal.", titulo_em_primeiro_plano())
    exigir_foco(exigir_janela(TITULO_JMS, time_limit=5))


def reposicionar_formulario(mudou_matricula: bool) -> None:
    """3 TABs para voltar ao campo do ano (mesmos intervalos da versão anterior)."""
    time.sleep(ESPERA_REPOSICIONAR)
    pyautogui.press("tab", presses=3, interval=INTERVALO_TAB_MATRICULA if mudou_matricula else INTERVALO_TAB_ANO)


@cronometrar()
def gerar_ficha(matricula: str, ano: int) -> None:
    """Preenche o formulário e dispara a impressão (sequência do legado preservada)."""
    time.sleep(ESPERA_FORMULARIO)  # janela desenhada: não há controle para consultar o estado

    pyautogui.write(str(ano), interval=0.001)
    pyautogui.press("tab")
    pyautogui.write(matricula, interval=0.001)
    pyautogui.press("tab")
    pyautogui.write(matricula, interval=0.001)
    pyautogui.press("tab", presses=9, interval=0.001)
    pyautogui.press("enter")
    pyautogui.press("tab")
    pyautogui.press("enter")


# ======================================================================
# SAÍDA DE IMPRESSÃO / SALVAMENTO
# ======================================================================
def _titulo(janela) -> str:
    try:
        return janela.window_text().strip()
    except EXC_AUTOMACAO as exc:
        raise ErroAutomacao(f"A janela desapareceu: {exc}") from exc


def _classificar(janela) -> str | None:
    titulo = _titulo(janela).casefold()
    if titulo == TITULO_EXTERNO_IMPRIMIR.casefold():
        return MODO_EXTERNO_IMPRIMIR
    if titulo == TITULO_EXTERNO_PRINT.casefold():
        return MODO_EXTERNO_PRINT
    return None


@cronometrar()
def detectar_modo(report_antigos: frozenset[int] = frozenset()):
    """Descobre o modo de impressão: (modo, janela). Timeout sem diálogo = modo interno (como antes)."""
    trechos = (TRECHO_REPORT,) if MODO_INTERNO_RAPIDO else ()
    janela = aguardar_janela(
        TITULO_EXTERNO_IMPRIMIR,
        TITULO_EXTERNO_PRINT,
        trechos=trechos,
        ignorar=report_antigos,
        time_limit=TEMPO_LIMITE,
    )
    if janela is None:
        return MODO_INTERNO, None

    modo = _classificar(janela)
    if modo is not None:
        return modo, janela

    # Surgiu um Report Designer novo: dá uma folga para o diálogo externo aparecer.
    externa = aguardar_janela(
        TITULO_EXTERNO_IMPRIMIR, TITULO_EXTERNO_PRINT, time_limit=TEMPO_GRACE_MODO
    )
    if externa is not None:
        modo = _classificar(externa)
        if modo is not None:
            return modo, externa
    return MODO_INTERNO, None


@cronometrar()
def salvar_temporario() -> int:
    """Preenche a janela 'Salvar Saída de Impressão como' e espera o PDF válido. Devolve nº de páginas."""
    deletar_temporario(CAMINHO_TEMPORARIO)

    janela = exigir_janela(TITULO_IMPRESSAO)
    exigir_foco(janela)

    # UMA enumeração de controles serve ao campo de nome e ao botão.
    try:
        controles = janela.descendants()
    except EXC_AUTOMACAO as exc:
        raise ErroAutomacao(f"Falha ao listar controles da janela de salvamento: {exc}") from exc

    campo = None
    botao = None
    for controle in controles:
        props = propriedades_controle(controle)
        if props is None:
            continue
        texto, classe = props
        if campo is None and classe == "Edit":
            try:
                if controle.rectangle().width() > 500:
                    campo = controle
            except EXC_AUTOMACAO:
                pass
        elif botao is None and texto == "Sa&lvar":
            botao = controle
        if campo is not None and botao is not None:
            break

    if campo is None:
        raise ErroAutomacao("Campo do nome do arquivo não encontrado.")
    if botao is None:
        raise ErroAutomacao("Botão 'Salvar' não encontrado.")

    _preencher_nome(campo, str(CAMINHO_TEMPORARIO))
    botao.click()

    return aguardar_pdf_pronto(CAMINHO_TEMPORARIO)


def _preencher_nome(campo, valor: str) -> None:
    """Cola o caminho no campo e confere o resultado; se falhar, define o texto direto."""

    def colar() -> None:
        campo.set_focus()
        pyperclip.copy(valor)
        campo.type_keys("^a", set_foreground=False)
        campo.type_keys("^v", set_foreground=False)

    def definir() -> None:
        campo.set_edit_text(valor)

    for metodo in (colar, definir):
        try:
            metodo()
            if campo.window_text().casefold() == valor.casefold():
                return
        except EXC_AUTOMACAO as exc:
            log.warning("Preenchimento do nome falhou (%s): %s", metodo.__name__, exc)
    raise ErroAutomacao("Não foi possível preencher o nome do arquivo na janela de salvamento.")


def tratar_report_designer(obrigatorio: bool) -> bool:
    """Modo interno: duplo clique no Report Designer abre a janela de salvamento."""
    janela = aguardar_janela(
        TITULO_REPORT,
        TITULO_REPORT_VARIACAO,
        trechos=(TRECHO_REPORT,),
        time_limit=TEMPO_REPORT if obrigatorio else TEMPO_REPORT_OPCIONAL,
    )
    if janela is None:
        if obrigatorio:
            raise JanelaNaoEncontrada("Report Designer não encontrado.")
        log.info("Report Designer ausente (opcional neste modo); seguindo.")
        return False

    exigir_foco(janela)
    x, y = REPORT_CLIQUE
    try:
        retangulo = janela.rectangle()
    except EXC_AUTOMACAO as exc:
        raise ErroAutomacao(f"Report Designer desapareceu: {exc}") from exc
    if not (retangulo.left <= x < retangulo.right and retangulo.top <= y < retangulo.bottom):
        raise ErroAutomacao(f"Ponto de clique {REPORT_CLIQUE} fora do Report Designer {retangulo}.")

    pyautogui.doubleClick(x, y, interval=0.1)
    salvar_temporario()  # espera a própria janela de salvamento: dispensa sleep
    return True


def _tratar_externo_imprimir(janela) -> None:
    exigir_foco(janela)
    exigir_controle(janela, "&Sim").click()
    salvar_temporario()
    tratar_report_designer(obrigatorio=False)

    # Confirmação final: o wrapper anterior pode ter sido destruído; localiza de novo.
    janela = exigir_janela(TITULO_EXTERNO_IMPRIMIR, time_limit=10)
    exigir_foco(janela)
    exigir_controle(janela, "&Sim").click()


def _tratar_externo_print(janela) -> None:
    exigir_foco(janela)
    exigir_controle(janela, "OK").click()
    salvar_temporario()


@cronometrar()
def gerar_e_salvar_pdf(matricula: str, ano: int) -> str:
    """Gera a ficha e deixa o PDF validado em CAMINHO_TEMPORARIO. Devolve o modo usado."""
    report_antigos = hwnds_com_trecho(TRECHO_REPORT) if MODO_INTERNO_RAPIDO else frozenset()
    _garantir_foco_formulario()
    gerar_ficha(matricula, ano)

    modo, janela = detectar_modo(report_antigos)
    log.info("Modo de saída de impressão: %s", modo)
    if modo == MODO_EXTERNO_IMPRIMIR:
        _tratar_externo_imprimir(janela)
    elif modo == MODO_EXTERNO_PRINT:
        _tratar_externo_print(janela)
    else:
        tratar_report_designer(obrigatorio=True)
    return modo
