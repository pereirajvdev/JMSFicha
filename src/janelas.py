"""Acesso a janelas e controles do Windows (Win32 via ctypes + wrappers pywinauto).

Regras deste módulo:
* uma enumeração de janelas por verificação, sem criar wrapper até haver acerto;
* janelas travadas (IsHungAppWindow) são ignoradas;
* o wrapper só é criado para um handle ainda válido, e o foco é verificado
  antes de qualquer digitação.
"""

from __future__ import annotations

import ctypes
import logging
import time
from ctypes import wintypes
from typing import Iterable, Sequence

import pywintypes
from pywinauto import Desktop
from pywinauto.controls.hwndwrapper import InvalidWindowHandle
from pywinauto.findwindows import ElementAmbiguousError, ElementNotFoundError

from config import (
    INTERVALO_VERIFICACAO,
    TEMPO_LIMITE,
    ControleNaoEncontrado,
    ErroAutomacao,
    JanelaNaoEncontrada,
)
from perf import cronometrar, medir

log = logging.getLogger(__name__)

# Falhas esperadas ao conversar com janelas que podem sumir a qualquer momento.
EXC_AUTOMACAO = (
    ElementNotFoundError,
    ElementAmbiguousError,
    InvalidWindowHandle,
    pywintypes.error,
    OSError,
)

# DLL privada: não altera os argtypes da user32 compartilhada com pywinauto/pyautogui.
_user32 = ctypes.WinDLL("user32")
_WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
GA_ROOTOWNER = 3

_user32.EnumWindows.argtypes = [_WNDENUMPROC, wintypes.LPARAM]
_user32.EnumWindows.restype = wintypes.BOOL
_user32.IsWindow.argtypes = [wintypes.HWND]
_user32.IsWindow.restype = wintypes.BOOL
_user32.IsWindowVisible.argtypes = [wintypes.HWND]
_user32.IsWindowVisible.restype = wintypes.BOOL
_user32.IsHungAppWindow.argtypes = [wintypes.HWND]
_user32.IsHungAppWindow.restype = wintypes.BOOL
_user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
_user32.GetWindowTextLengthW.restype = ctypes.c_int
_user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
_user32.GetWindowTextW.restype = ctypes.c_int
_user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
_user32.GetWindowThreadProcessId.restype = wintypes.DWORD
_user32.GetForegroundWindow.argtypes = []
_user32.GetForegroundWindow.restype = wintypes.HWND
_user32.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]
_user32.GetAncestor.restype = wintypes.HWND

_desktop: Desktop | None = None
_avisos_travada: set[int] = set()


# ======================================================================
# ENUMERAÇÃO
# ======================================================================
def listar_janelas_visiveis() -> list[tuple[int, str]]:
    """Janelas de nível superior visíveis e com título: [(hwnd, titulo), ...]."""
    resultado: list[tuple[int, str]] = []

    def callback(hwnd, _lparam):
        try:
            if _user32.IsWindowVisible(hwnd):
                tamanho = _user32.GetWindowTextLengthW(hwnd)
                if tamanho:
                    buffer = ctypes.create_unicode_buffer(tamanho + 1)
                    _user32.GetWindowTextW(hwnd, buffer, tamanho + 1)
                    resultado.append((int(hwnd), buffer.value))
        except OSError:
            pass  # uma janela problemática não pode interromper a enumeração
        return True

    proc = _WNDENUMPROC(callback)  # precisa continuar referenciado durante EnumWindows
    _user32.EnumWindows(proc, 0)
    return resultado


def janela_viva(hwnd: int) -> bool:
    return bool(_user32.IsWindow(hwnd))


def janela_visivel(hwnd: int) -> bool:
    return janela_viva(hwnd) and bool(_user32.IsWindowVisible(hwnd))


def pid_da_janela(hwnd: int) -> int:
    pid = wintypes.DWORD(0)
    _user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return int(pid.value)


def hwnds_com_trecho(trecho: str) -> frozenset[int]:
    """hwnds das janelas visíveis cujo título contém ``trecho``."""
    chave = trecho.casefold()
    return frozenset(h for h, t in listar_janelas_visiveis() if chave in t.casefold())


def titulo_em_primeiro_plano() -> str:
    hwnd = _user32.GetForegroundWindow()
    if not hwnd:
        return ""
    tamanho = _user32.GetWindowTextLengthW(hwnd)
    buffer = ctypes.create_unicode_buffer(tamanho + 1)
    _user32.GetWindowTextW(hwnd, buffer, tamanho + 1)
    return buffer.value


def _responde(hwnd: int, texto: str) -> bool:
    if _user32.IsHungAppWindow(hwnd):
        if hwnd not in _avisos_travada:
            _avisos_travada.add(hwnd)
            log.warning("Janela sem resposta ignorada: %s", texto)
        return False
    return True


def buscar_janela(
    titulos: Sequence[str],
    trechos: Sequence[str] = (),
    ignorar: Iterable[int] = (),
) -> tuple[int, str] | None:
    """UMA enumeração; devolve (hwnd, título) da primeira janela válida.

    Prioridade: títulos exatos (sem diferenciar maiúsculas), na ordem dada;
    depois, janelas cujo título contém algum trecho (exceto hwnds em ``ignorar``).
    """
    janelas = listar_janelas_visiveis()
    por_titulo: dict[str, list[tuple[int, str]]] = {}
    for hwnd, texto in janelas:
        texto = texto.strip()
        por_titulo.setdefault(texto.casefold(), []).append((hwnd, texto))

    for titulo in titulos:
        for hwnd, texto in por_titulo.get(titulo.strip().casefold(), ()):
            if _responde(hwnd, texto):
                return hwnd, texto

    ignorados = frozenset(ignorar)
    for trecho in trechos:
        chave = trecho.casefold()
        for hwnd, texto in janelas:
            if hwnd in ignorados or chave not in texto.casefold():
                continue
            if _responde(hwnd, texto):
                return hwnd, texto.strip()
    return None


# ======================================================================
# WRAPPERS E FOCO
# ======================================================================
def criar_wrapper(hwnd: int):
    """Wrapper pywinauto do handle, ou None se a janela já sumiu."""
    global _desktop
    if not janela_visivel(hwnd):
        return None
    if _desktop is None:
        _desktop = Desktop(backend="win32")
    try:
        return _desktop.window(handle=hwnd).wrapper_object()
    except EXC_AUTOMACAO as exc:
        log.debug("Wrapper indisponível para %s: %s", hwnd, exc)
        return None


@cronometrar()
def aguardar_janela(
    *titulos: str,
    time_limit: float = TEMPO_LIMITE,
    trechos: Sequence[str] = (),
    ignorar: Iterable[int] = (),
):
    """Espera uma janela com um dos títulos (ou trechos). Devolve o wrapper ou None."""
    ignorar = frozenset(ignorar)
    log.debug("Aguardando janela: %s", ", ".join((*titulos, *trechos)))
    limite = time.monotonic() + time_limit

    while True:
        with medir("aguardar_janela:enumerar"):
            achado = buscar_janela(titulos, trechos, ignorar)
        if achado is not None:
            wrapper = criar_wrapper(achado[0])
            if wrapper is not None:
                log.info("Janela identificada: %s", achado[1])
                return wrapper
        restante = limite - time.monotonic()
        if restante <= 0:
            break
        with medir("aguardar_janela:espera"):
            time.sleep(min(INTERVALO_VERIFICACAO, restante))

    log.warning("Nenhuma janela válida em %.1fs: %s", time_limit, ", ".join((*titulos, *trechos)))
    return None


def exigir_janela(*titulos: str, time_limit: float = TEMPO_LIMITE, trechos: Sequence[str] = ()):
    janela = aguardar_janela(*titulos, time_limit=time_limit, trechos=trechos)
    if janela is None:
        raise JanelaNaoEncontrada("Janela não encontrada: " + ", ".join((*titulos, *trechos)))
    return janela


def _foreground_pertence(hwnd: int) -> bool:
    primeiro = _user32.GetForegroundWindow()
    if not primeiro:
        return False
    primeiro = int(primeiro)
    return primeiro == hwnd or int(_user32.GetAncestor(primeiro, GA_ROOTOWNER) or 0) == hwnd


def focar(janela, tentativas: int = 3, espera: float = 1.0) -> bool:
    """Coloca a janela em primeiro plano e CONFIRMA que o foco chegou nela."""
    hwnd = int(janela.handle)
    for _ in range(tentativas):
        if not janela_viva(hwnd):
            return False
        try:
            janela.set_focus()
        except EXC_AUTOMACAO as exc:
            log.warning("set_focus falhou: %s", exc)
        limite = time.monotonic() + espera
        while True:
            if _foreground_pertence(hwnd):
                return True
            if time.monotonic() >= limite:
                break
            time.sleep(0.05)
    return False


def exigir_foco(janela) -> None:
    """Como focar(), mas levanta erro: nada é digitado/clicado sem foco confirmado."""
    try:
        titulo = janela.window_text()
    except EXC_AUTOMACAO:
        titulo = "?"
    if not focar(janela):
        raise ErroAutomacao(f"Não foi possível dar foco à janela: {titulo}")


# ======================================================================
# CONTROLES
# ======================================================================
def propriedades_controle(controle) -> tuple[str, str] | None:
    """(texto, classe) do controle, ou None se ele sumiu/não responde."""
    try:
        return controle.window_text(), controle.class_name()
    except EXC_AUTOMACAO:
        return None


def procurar_controle(janela, titulo: str | None = None, classe: str | None = None):
    """Primeiro controle que atende a TODOS os critérios dados (ou None).

    Procura primeiro nos filhos diretos (barato) e só depois em todos os descendentes.
    """
    if titulo is None and classe is None:
        raise ValueError("Informe titulo e/ou classe.")
    for listar in (janela.children, janela.descendants):
        try:
            controles = listar()
        except EXC_AUTOMACAO as exc:
            log.debug("Falha ao listar controles: %s", exc)
            continue
        for controle in controles:
            props = propriedades_controle(controle)
            if props is None:
                continue
            if titulo is not None and props[0] != titulo:
                continue
            if classe is not None and props[1] != classe:
                continue
            return controle
    return None


def exigir_controle(janela, titulo: str | None = None, classe: str | None = None):
    controle = procurar_controle(janela, titulo, classe)
    if controle is None:
        try:
            onde = janela.window_text()
        except EXC_AUTOMACAO:
            onde = "?"
        raise ControleNaoEncontrado(
            f"Controle (título={titulo!r}, classe={classe!r}) não encontrado em '{onde}'."
        )
    return controle
