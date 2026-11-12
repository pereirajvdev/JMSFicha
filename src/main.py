import sys
import time
import win32gui
import win32print
import ctypes
import pyautogui
import subprocess
import re
import shutil
import fitz
import argparse
import psutil
import subprocess

import traceback
from contrato import adicionar_argumentos, pasta_existente, registros_para_dados, Retorno

from ctypes import wintypes
from pathlib import Path
from pywinauto import Desktop
from openpyxl import load_workbook

import subprocess
import pyautogui
import argparse
import ctypes
import re
import time
from contextlib import contextmanager
from ctypes import wintypes
from functools import wraps
from pathlib import Path


from pywinauto import Desktop, keyboard, mouse
from pywinauto.timings import Timings

import pyperclip

# ======================================================================
# WIN32 (ctypes): enumeração rápida de janelas
# ======================================================================
_user32 = ctypes.windll.user32

_WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

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


# Instância única de Desktop, usada só para criar o wrapper da janela escolhida.
_DESKTOP = Desktop(backend="win32")

_AVISOS_TRAVADA = set()

EXE = Path(r"C:\Folha\folha.exe")

RAIZ_PROJETO = Path(__file__).resolve().parents[1]   # se o arquivo está em src/
CRED = RAIZ_PROJETO / "data" / "credencial.txt"


# ============================================================
# CONFIGURAÇÕES
# ============================================================

TITULO_GERADOR = (
    "Gerador de Relatórios "
    "( Versão: 4.0 ) - Recursos Humanos e Folha de Pagamento"
)
TITULO_JANELA_EL = (
    "Recursos Humanos e Folha de Pagamento "
    "[ Versão 3.02 - 17/06/2026 ]"
)
TITULO_INICIAL = "Form1"
TITULO_JMS = "Folha - PREFEITURA MUNICIPAL DE TRES CORACOES - Abril de 2009"
TITULO_INTERNO = "Form1"
TITULO_INTERNO_IMPRESSAO = "Report Designer - fichafin.frx"
TITULO_EXTERNO_IMPRIMIR = "Imprimir"
TITULO_EXTERNO_PRINT = "Print"
TITULO_IMPRESSAO = "Salvar Saída de Impressão como"
TITULO_REPORT_VARIACAO = "Report Designer - fichafin.frx - Page 1 - Folha - PREFEITURA MUNICIPAL DE TRES CORACOES - Abril de 2009"
TITULO_REPORT = "Report Designer - fichafin.frx - Folha - PREFEITURA MUNICIPAL DE TRES CORACOES - Abril de 2009"

PASTA_DOWNLOADS = Path.home() / "Downloads"
PASTA_SERVIDORES = Path(
    r"C:\Users\Joao Castro\Desktop\JOAO\Servidores"
)

NOME_TEMPORARIO = "temporario.pdf"

TEMPO_LIMITE = 60
INTERVALO_VERIFICACAO = 1

# ============================================================
# LOGIN
# ============================================================
def fechar_jms():
    """Encerra o processo associado à janela principal, se estiver aberto."""

    print("[INFO] Verificando se a folha já está aberta...")

    janela = aguardar_janela(TITULO_INICIAL, TITULO_JMS, TITULO_REPORT, time_limit=.5)

    while janela is not None:
        try:
            pid = janela.process_id()
            processo = psutil.Process(pid)

            print(
                f"[INFO] Encerrando {processo.name()} "
                f"(PID {pid})..."
            )

            processo.terminate()

            try:
                processo.wait(timeout=1)
            except psutil.TimeoutExpired:
                processo.kill()
                processo.wait(timeout=1)

            print("[OK] Instância anterior encerrada.")

        except psutil.NoSuchProcess:
            print("[INFO] O processo já foi encerrado.")

        except psutil.AccessDenied:
            raise RuntimeError(
                "Acesso negado ao processo da folha. "
                "Verifique as permissões do Windows."
            )

        janela = aguardar_janela(TITULO_INICIAL, TITULO_JMS, TITULO_REPORT, time_limit=.5)


def abrir_jms():

    print(f"EXE: {EXE}")
    print(f"EXE existe: {EXE.exists()}")
    print(f"Diretório: {EXE.parent}")
    print(f"Diretório existe: {EXE.parent.exists()}")

    if not EXE.exists():
        print()
        print("ERRO: executável do Folha não encontrado.")
        print(f"Caminho procurado: {EXE}")
        return False

    fechar_jms()

    print("Abrindo Folha...")

    subprocess.Popen(
        [str(EXE)],
        cwd=str(EXE.parent)
    )
    
    return True


def abrir_janela_interna():
    janela = aguardar_janela(TITULO_JMS, time_limit=3)
    if janela == None: # passa pela janela de aviso com mesmo nome: Form1
        pyautogui.press("enter")
        janela = aguardar_janela(TITULO_JMS, time_limit=10)
    janela.set_focus()

    reposicionar(janela, "left")
    time.sleep(1)

    pyautogui.hotkey("alt")
    time.sleep(.02)

    pyautogui.hotkey("r")
    time.sleep(.02)

    pyautogui.press(
        "down",
        presses=6,
        interval=.001
    )

    time.sleep(.02)

    pyautogui.press("right")
    time.sleep(.02)

    pyautogui.press(
        "down",
        presses=9,
        interval=.001
    )

    time.sleep(.02)

    pyautogui.press("enter")
    time.sleep(.02)

    return True


def reposicionar(janela, lado="right"):
    """
    Reposiciona a janela ocupando metade da área útil da tela.

    lado:
        "left"  -> metade esquerda
        "right" -> metade direita
    """

    largura_tela, altura_tela = pyautogui.size()

    metade = largura_tela // 2

    if lado == "left":
        x = 0
    elif lado == "right":
        x = metade
    else:
        raise ValueError("lado deve ser 'left' ou 'right'.")

    # Restaura a janela caso esteja maximizada ou minimizada.
    if janela.is_maximized() or janela.is_minimized():
        janela.restore()

    # Reposiciona e redimensiona.
    janela.move_window(
        x=x,
        y=0,
        width=metade,
        height=altura_tela,
        repaint=True
    )

    return True


def login():

    print()
    print("=" * 80)
    print("INICIANDO LOGIN")
    print("=" * 80)

    usuario, senha = CRED.read_text(
        encoding="utf-8"
    ).splitlines()[:2]

    janela = aguardar_janela(TITULO_INICIAL)
    janela.set_focus()

    print("Preenchendo usuário...")

    pyautogui.write(
        usuario,
        interval=.001
    )

    pyautogui.press("tab")
    time.sleep(.005)

    pyautogui.write(
        senha,
        interval=.001
    )
    
    pyautogui.press("tab")
    time.sleep(.005)

    pyautogui.press("enter")

    print("Login concluído.")

    return True

    
# ============================================================
# LEITURA DA PLANILHA
# ============================================================

def ler_planilha(caminho):

    wb = load_workbook(
        caminho,
        data_only=True
    )

    ws = wb.active

    registros = {}

    for linha in ws.iter_rows(
        min_row=2,
        values_only=True
    ):

        matricula, data_inicial, data_final = linha[:3]

        if not matricula or not data_inicial or not data_final:
            continue

        matricula = str(matricula)

        anos = range(
            data_inicial.year,
            data_final.year + 1
        )

        if matricula not in registros:
            registros[matricula] = set()

        registros[matricula].update(anos)

    dados = []

    for matricula, anos in registros.items():

        dados.append(
            (
                matricula,
                sorted(anos)
            )
        )

    return dados


# ============================================================
# DEFINIR IMPRESSORA PDF
# ============================================================

def definir_impressora_pdf():
    """Executa o AutoFicha no modo padrão."""

    script = r"C:\Users\Joao Castro\Documents\AutoFicha\src\main.py"

    subprocess.Popen(
        [sys.executable, script, "--padrao"],
        cwd=r"C:\Users\Joao Castro\Documents\AutoFicha"
    )


def impressora_eh_pdf():
    impressora = win32print.GetDefaultPrinter()
    return impressora.strip().lower() == "microsoft print to pdf"


def impressora_pdf():
    
    if impressora_eh_pdf():
        print("Impressora padrão já é Microsoft Print to PDF.")
    else:
        definir_impressora_pdf()

        if not impressora_eh_pdf():
            print("ATENÇÃO: Definir impressora falhou. ")
            print("ATENÇÃO: A impressora padrão não é Microsoft Print to PDF.")

        return False
    
    return True

# ============================================================
# GERAR FICHA
# ============================================================

def gerarFicha(matricula, ano):
    # INSERE DADOS NA JANELA DESENHADA, NAO EH POSSIVEL DAR controle.set_focus()
    time.sleep(1)

    pyautogui.write(
        str(ano),
        interval=.001
    )

    pyautogui.press("tab")
    time.sleep(.02)

    pyautogui.write(
        matricula,
        interval=.001
    )

    pyautogui.press("tab")
    time.sleep(.02)

    pyautogui.write(
        matricula,
        interval=.001
    )

    time.sleep(.02)

    pyautogui.press(
        "tab",
        presses=9,
        interval=.001
    )

    time.sleep(.02)

    pyautogui.press("enter")
    time.sleep(.02)

    pyautogui.press("tab")
    time.sleep(.02)

    pyautogui.press("enter")


# ============================================================
# SALVAMENTO DE FICHA
# ============================================================
def procurar_controle(janela, titulo=None, classe=None):
    controles = janela.descendants()

    for i, controle in enumerate(controles):
        try:
            if titulo is not None and controle.window_text() == titulo:
                return controle
        
            if classe is not None and controle.class_name() == classe:
                return controle
            
        except Exception as e:
            print(f"[{i}] Erro: {e}")

    print(f"{titulo} não foi encontrado em {janela.window_text()}. ")

    return 
            

def tratar_externo_imprimir(janela):
    print("Iniciando salvamento pelo modo externo: IMPRIMIR")

    procurar_controle(janela, "&Sim").click()

    if not salvar_temporario():
        input("Erro no salvamento")

    tratar_interno_print()

    janela.set_focus()
    botao_ok = janela.child_window(
        title="&Sim",
        class_name="Button"
    )

    print(f"botao ok: {botao_ok}")
    botao_ok.set_focus()
    botao_ok.click()


def tratar_externo_print(janela):
    print("Iniciando salvamento pelo modo externo: PRINT")

    ok_btn = procurar_controle(janela, "OK")

    if ok_btn is None:
        input("Botao de ok nao identificado. ")

    janela.set_focus()
    ok_btn.click()

    if not salvar_temporario():
        input("Erro no salvamento")


def tratar_interno_print():
    print("Iniciando salvamento pelo modo interno: REPORT DESIGN")

    input("Continuar...")
    try:

        hwnd_report = None

        def procurar_report(hwnd, _):

            nonlocal hwnd_report

            if not win32gui.IsWindowVisible(hwnd):
                return

            titulo = (
                win32gui.GetWindowText(hwnd)
            )

            if "Report Designer" in titulo:

                hwnd_report = hwnd

        win32gui.EnumWindows(
            procurar_report,
            None
        )

        if hwnd_report:

            print()
            print(
                "REPORT DESIGNER ENCONTRADO."
            )

            janela = aguardar_janela(TITULO_REPORT)
            janela.set_focus()

            pyautogui.doubleClick(
                373,
                325,
                interval=0.1
            )

            time.sleep(0.5)

            if not salvar_temporario():
                input("Erro no salvamento")

    except Exception as e:

        print(
            "Erro procurando Report Designer: "
            f"{e}"
        )


def verificar_modo():
    janela = aguardar_janela(
        TITULO_EXTERNO_PRINT, 
        TITULO_EXTERNO_IMPRIMIR,
        time_limit=7
    )

    if janela == None: 
        return "interno_print", janela
    if janela.window_text() == TITULO_EXTERNO_IMPRIMIR: 
        return "externo_imprimir", janela
    if janela.window_text() == TITULO_EXTERNO_PRINT: 
        return "externo_print", janela

    return None, None


def tratar_saida_impressao():

    print()
    print("=" * 80)
    print("AGUARDANDO SAÍDA DE IMPRESSÃO")
    print("=" * 80)
    
    modo, janela = verificar_modo()
    
    if janela != None:
        janela.set_focus()
        titulo = janela.window_text()
        
    #input(f"\nVerifique se está certo\nModo: {modo}\nJanela: {titulo}\n")

    if modo == "externo_imprimir":
        tratar_externo_imprimir(janela)
    if modo == "externo_print":
        tratar_externo_print(janela)
    if modo == "interno_print":
        tratar_interno_print()

    if modo == None:
        input("Nenhum modo foi selecionado")

    return True


# ============================================================
# TRATAMENTO DO TEMPORARIO
# ============================================================
def deletar_temporario():
    caminho_temporario = PASTA_DOWNLOADS / NOME_TEMPORARIO

    print()
    print("=" * 80)
    print("VERIFICANDO PDF TEMPORÁRIO")
    print("=" * 80)

    if caminho_temporario.exists():
        try:
            caminho_temporario.unlink()

            print(
                f"PDF temporário anterior deletado: "
                f"{caminho_temporario}"
            )

        except Exception as e:
            print(
                f"ERRO ao deletar PDF temporário: {e}"
            )
            return False

    else:
        print("Nenhum PDF temporário anterior encontrado.")

    return True


def colar_texto(controle, valor):
    """
    Foca o controle UMA vez, copia o valor e cola substituindo o conteúdo.
    set_foreground=False evita que type_keys refaça set_focus a cada chamada
    (por padrão o pywinauto chama set_focus dentro de cada type_keys).
    """
    controle.set_focus()

    pyperclip.copy(str(valor))

    controle.type_keys("^a", set_foreground=False)
    controle.type_keys("^v", set_foreground=False)


def salvar_temporario():

    if not deletar_temporario():
        return False

    janela = aguardar_janela(TITULO_IMPRESSAO)
    janela.set_focus()

    caminho_arquivo = PASTA_DOWNLOADS / NOME_TEMPORARIO

    print("Arquivo:", NOME_TEMPORARIO)
    print("Caminho:", caminho_arquivo)

    # OBTÉM CAMPO DO NOME DO ARQUIVO

    campo_nome = None

    for controle in janela.descendants():
        try:
            if controle.class_name() != "Edit":
                continue

            rect = controle.rectangle()

            if rect.width() > 500:
                campo_nome = controle
                break
        except Exception:
            pass
        
    if campo_nome is None:
        print("Campo do nome do arquivo não encontrado.")
        return False

    # ESCREVE temporario.pdf

    colar_texto(campo_nome, caminho_arquivo)

    print("Nome do arquivo preenchido.")

    # BOTÃO SALVAR

    botao_salvar = procurar_controle(janela, "Sa&lvar")
    botao_salvar.set_focus()
    botao_salvar.click()
    
    # ============================================================
    # AGUARDAR PDF FICAR REALMENTE PRONTO
    # ============================================================

    print()
    print("Aguardando gravação completa do PDF...")

    limite = time.time() + 15

    ultimo_tamanho = -1
    tamanho_estavel = 0

    while time.time() < limite:

        try:

            if caminho_arquivo.exists():

                tamanho = caminho_arquivo.stat().st_size

                print(
                    f"PDF encontrado. "
                    f"Tamanho: {tamanho:,} bytes"
                )

                if tamanho > 0:

                    if tamanho == ultimo_tamanho:
                        tamanho_estavel += 1
                    else:
                        tamanho_estavel = 0

                    ultimo_tamanho = tamanho

                    if tamanho_estavel >= 3:

                        try:
                            documento = fitz.open(
                                caminho_arquivo
                            )

                            paginas = documento.page_count

                            documento.close()

                            print()
                            print(
                                "PDF temporário gravado "
                                "e validado com sucesso."
                            )

                            print(
                                f"Páginas: {paginas}"
                            )

                            print(
                                f"Arquivo: {caminho_arquivo}"
                            )

                            return True

                        except Exception as e:
                            print(
                                "PDF existe, mas ainda está "
                                f"sendo finalizado: {e}"
                            )

        except Exception as e:
            print(
                f"Erro verificando PDF: {e}"
            )

        time.sleep(0.3)

    print()
    print(
        "ERRO: PDF temporário não ficou disponível "
        "para leitura."
    )

    if caminho_arquivo.exists():
        try:
            tamanho = caminho_arquivo.stat().st_size

            print(
                f"Tamanho final encontrado: "
                f"{tamanho:,} bytes"
            )

        except Exception:
            pass

    return False


# ============================================================
# EXTRAIR DADOS DO SERVIDOR
# ============================================================

def extrair_dados_servidor_modelo_antigo(texto: str) -> dict:
    """
    Extrai os dados do servidor no modelo antigo
    de ficha financeira.
    """

    dados = {}

    # --------------------------------------------------------
    # MATRÍCULA + NOME
    # --------------------------------------------------------

    padrao = re.search(
        r"(?m)^Matricula:\s*(\d+)\s+(.+?)\s*$",
        texto,
        flags=re.IGNORECASE
    )

    if padrao:
        dados["Matrícula"] = padrao.group(1)
        dados["Nome"] = padrao.group(2).strip()

    # --------------------------------------------------------
    # CARGO
    # --------------------------------------------------------

    padrao = re.search(
        r"(?m)^Cargo:\s*(.+?)\s*$",
        texto,
        flags=re.IGNORECASE
    )

    if padrao:
        dados["Cargo"] = padrao.group(1).strip()

    # --------------------------------------------------------
    # ANO
    # --------------------------------------------------------

    padrao = re.search(
        r"Relatório\s+da\s+Ficha\s+Financeira\s+(\d{4})",
        texto,
        flags=re.IGNORECASE
    )

    if padrao:
        dados["Ano"] = int(padrao.group(1))

    return dados


# ============================================================
# EXTRAI TEXTO DO PDF
# ============================================================

def extrair_texto_pdf(caminho_pdf):

    print()
    print("=" * 80)
    print("EXTRAINDO DADOS DO PDF")
    print("=" * 80)

    print(f"PDF: {caminho_pdf}")

    try:

        documento = fitz.open(caminho_pdf)

        textos = []

        for pagina in documento:
            textos.append(
                pagina.get_text()
            )

        documento.close()

        texto = "\n".join(textos)

        print(
            f"Caracteres extraídos: {len(texto)}"
        )

        return texto

    except Exception as e:

        print(
            f"ERRO ao extrair texto do PDF: {e}"
        )

        return None


# ============================================================
# SANITIZA NOME DA PASTA
# ============================================================

def sanitizar_nome(nome):

    # Remove caracteres inválidos para nome de pasta
    nome = re.sub(
        r'[<>:"/\\|?*]',
        '',
        nome
    )

    # Remove espaços duplicados
    nome = re.sub(
        r'\s+',
        ' ',
        nome
    )

    return nome.strip()


# ============================================================
# MOVER PDF PARA PASTA DO SERVIDOR
# ============================================================

def organizar_pdf_temporario():

    caminho_temporario = (
        PASTA_DOWNLOADS /
        NOME_TEMPORARIO
    )

    print()
    print("=" * 80)
    print("ORGANIZANDO FICHA FINANCEIRA")
    print("=" * 80)

    print(
        f"Temporário: {caminho_temporario}"
    )

    if not caminho_temporario.exists():

        print(
            "ERRO: PDF temporário não encontrado."
        )

        return False

    # --------------------------------------------------------
    # EXTRAI TEXTO
    # --------------------------------------------------------

    texto = extrair_texto_pdf(
        caminho_temporario
    )

    if not texto:

        print(
            "ERRO: não foi possível extrair "
            "texto do PDF."
        )

        return False

    # --------------------------------------------------------
    # EXTRAI DADOS
    # --------------------------------------------------------

    dados = extrair_dados_servidor_modelo_antigo(
        texto
    )

    print()
    print("DADOS EXTRAÍDOS:")

    for chave, valor in dados.items():

        print(
            f"{chave}: {valor}"
        )

    # --------------------------------------------------------
    # VALIDA NOME
    # --------------------------------------------------------

    nome_servidor = dados.get("Nome")

    if not nome_servidor:

        print(
            "ERRO: nome do servidor não encontrado."
        )

        return False

    # --------------------------------------------------------
    # VALIDA ANO
    # --------------------------------------------------------

    ano = dados.get("Ano")

    if not ano:

        print(
            "ERRO: ano da ficha financeira não encontrado."
        )

        return False

    # --------------------------------------------------------
    # SANITIZA NOME
    # --------------------------------------------------------

    nome_servidor = sanitizar_nome(
        nome_servidor
    )

    # --------------------------------------------------------
    # CRIA PASTA DO SERVIDOR
    # --------------------------------------------------------

    pasta_servidor = (
        PASTA_SERVIDORES /
        nome_servidor
    )

    pasta_servidor.mkdir(
        parents=True,
        exist_ok=True
    )

    print()
    print(
        f"Pasta do servidor: {pasta_servidor}"
    )

    # --------------------------------------------------------
    # DEFINE DESTINO
    # --------------------------------------------------------

    destino = (
        pasta_servidor /
        f"FICHA FINANCEIRA {ano}.pdf"
    )

    print(
        f"Destino: {destino}"
    )

    # --------------------------------------------------------
    # MOVE
    # --------------------------------------------------------

    try:

        shutil.move(
            str(caminho_temporario),
            str(destino)
        )

        print()
        print("PDF ORGANIZADO COM SUCESSO.")
        print(
            f"Arquivo: {destino}"
        )

        return True

    except Exception as e:

        print()
        print(
            f"ERRO ao mover PDF: {e}"
        )

        return False


def avancar_para_proximo_ano():
    print()
    print("=" * 80)
    print("AVANÇANDO PARA O PRÓXIMO ANO")
    print("=" * 80)

    for i in range(3):
        pyautogui.press("tab")
        print(f"TAB {i + 1}/3")
        time.sleep(0.05)

    print("Pronto para o próximo ano.")

    return True


def preparar_proxima_matricula():
    print()
    print("=" * 80)
    print("PREPARANDO PRÓXIMA MATRÍCULA")
    print("=" * 80)

    time.sleep(0.5)

    # sequência para retornar ao campo inicial
    pyautogui.press("tab", presses=3, interval=0.02)

    print("Pronto para a próxima matrícula.")

    return True

# ======================================================================
# CRONOMETRAR
# ======================================================================
PERF_ATIVO = False
CACHE_ATIVO = True
_TEMPOS = {}

def _registrar_tempo(nome, duracao):
    registro = _TEMPOS.setdefault(nome, [0, 0.0])
    registro[0] += 1
    registro[1] += duracao


@contextmanager
def medir(nome):
    """Mede o tempo de um bloco. Sem efeito se --perf estiver desligado."""
    if not PERF_ATIVO:
        yield
        return

    inicio = time.perf_counter()
    try:
        yield
    finally:
        _registrar_tempo(nome, time.perf_counter() - inicio)


def cronometrar(nome=None):
    """Decorator que mede o tempo total de uma função (só com --perf)."""
    def decorador(func):
        rotulo = nome or func.__name__

        @wraps(func)
        def wrapper(*args, **kwargs):
            if not PERF_ATIVO:
                return func(*args, **kwargs)

            inicio = time.perf_counter()
            try:
                return func(*args, **kwargs)
            finally:
                _registrar_tempo(rotulo, time.perf_counter() - inicio)

        return wrapper
    return decorador


def imprimir_relatorio_tempos():
    if not PERF_ATIVO or not _TEMPOS:
        return

    print("\n" + "=" * 66)
    print(" RELATÓRIO DE DESEMPENHO (--perf)")
    print(" Obs.: tempos aninhados se sobrepõem; não some as linhas.")
    print("=" * 66)
    print(f"{'Etapa':<42}{'Chamadas':>9}{'Total (s)':>11}{'Média (s)':>11}")

    for nome, (chamadas, total) in sorted(
        _TEMPOS.items(), key=lambda item: item[1][1], reverse=True
    ):
        media = total / chamadas if chamadas else 0.0
        print(f"{nome:<42}{chamadas:>9}{total:>11.2f}{media:>11.3f}")

    print("=" * 66)

# ======================================================================
# JANELAS
# ======================================================================
def _listar_janelas_visiveis():
    """Enumera janelas de nível superior visíveis: [(hwnd, titulo), ...]."""
    resultado = []

    def callback(hwnd, _lparam):
        if _user32.IsWindowVisible(hwnd):
            tamanho = _user32.GetWindowTextLengthW(hwnd)
            if tamanho:
                buffer = ctypes.create_unicode_buffer(tamanho + 1)
                _user32.GetWindowTextW(hwnd, buffer, tamanho + 1)
                resultado.append((hwnd, buffer.value))
        return True

    proc = _WNDENUMPROC(callback)
    _user32.EnumWindows(proc, 0)
    return resultado


def _janela_visivel(hwnd):
    return bool(_user32.IsWindow(hwnd)) and bool(_user32.IsWindowVisible(hwnd))


def _buscar_janela(titulos):
    """
    Faz UMA enumeração e retorna (hwnd, titulo_real) da primeira janela
    válida, respeitando a ordem dos títulos informados.
    Comparação: texto idêntico (sem diferenciar maiúsculas/minúsculas),
    como na versão original. Janelas travadas são ignoradas.
    """
    candidatas = {}

    for hwnd, texto in _listar_janelas_visiveis():
        texto = texto.strip()
        candidatas.setdefault(texto.casefold(), []).append((hwnd, texto))

    for titulo in titulos:
        for hwnd, texto in candidatas.get(titulo.strip().casefold(), []):
            if _user32.IsHungAppWindow(hwnd):
                if hwnd not in _AVISOS_TRAVADA:
                    _AVISOS_TRAVADA.add(hwnd)
                    print(f"[AGUARDANDO] Janela sem resposta: {texto}")
                continue

            return hwnd, texto

    return None


def _criar_wrapper(hwnd):
    return _DESKTOP.window(handle=hwnd).wrapper_object()


def verificar_janela_interna(titulo_interno, titulo_janela):

    print(f"Aguardando: {titulo_interno} from {titulo_janela}")
    
    janela = aguardar_janela(titulo_janela)

    controles = janela.descendants()

    for i, controle in enumerate(controles):
        try:
            if titulo_interno == controle.window_text():
                print(f"Encontrado: {titulo_interno} from {titulo_janela}")
                return True

        except Exception as e:
            print(f"[{i}] Erro: {e}")

    return False


@cronometrar()
def localizar_janela(titulo_esperado):
    """Verificação única (sem espera). Retorna wrapper ou None."""
    achado = _buscar_janela([titulo_esperado])

    if achado is None:
        return None

    hwnd, titulo = achado
    print(f"[OK] Janela identificada: {titulo}")
    return _criar_wrapper(hwnd)


@cronometrar()
def aguardar_janela(*titulos_esperados, time_limit=TEMPO_LIMITE):
    """
    Aguarda até que uma janela com um dos títulos apareça.
    Retorna o wrapper da primeira janela válida (seguindo a ordem dos
    títulos) ou None se o tempo limite estourar.
    """
    print(
        f"[INFO] Aguardando uma das janelas: "
        f"{', '.join(titulos_esperados)}"
    )

    limite = time.monotonic() + time_limit

    while True:
        with medir("enumerar janelas"):
            achado = _buscar_janela(titulos_esperados)

        if achado is not None:
            hwnd, titulo = achado
            print(f"[OK] Janela identificada: {titulo}")
            return _criar_wrapper(hwnd)

        if time.monotonic() >= limite:
            break

        time.sleep(INTERVALO_VERIFICACAO)

    print(
        f"[INFO] Nenhuma janela válida apareceu em "
        f"{time_limit} segundos: "
        f"{', '.join(titulos_esperados)}"
    )

    return None


def exigir_janela(*titulos_esperados, time_limit=TEMPO_LIMITE):
    janela = aguardar_janela(*titulos_esperados, time_limit=time_limit)

    if janela is None:
        raise RuntimeError(
            "Janela não encontrada: " + ", ".join(titulos_esperados)
        )

    return janela


# ============================================================
# MAIN
# ============================================================
# Substitui TODO o trecho a partir do cabeçalho "# MAIN" até o fim do arquivo.

def obter_dados(args, parser):
    """Devolve [(matricula, [anos...]), ...] a partir de --registros, --mat ou da planilha."""

    if args.registros is not None:
        if args.mat or args.planilha:
            parser.error("Use --registros, --mat ou a planilha, apenas um deles.")
        return registros_para_dados(args.registros, parser)

    if args.mat:
        if args.planilha:
            parser.error("Use --mat ou o caminho da planilha, não os dois.")

        matricula, periodo = args.mat
        resultado = re.fullmatch(r"(\d{4})-(\d{4})", periodo)

        if not resultado:
            parser.error(
                "Período inválido. Use o formato ANO_INICIAL-ANO_FINAL, "
                "por exemplo: 2000-2005."
            )

        ano_inicial, ano_final = int(resultado.group(1)), int(resultado.group(2))

        if ano_inicial > ano_final:
            parser.error("O ano inicial não pode ser maior que o ano final.")

        return [(matricula, list(range(ano_inicial, ano_final + 1)))]

    if args.planilha:
        planilha = Path(args.planilha)

        if not planilha.is_file():
            parser.error(f"Planilha não encontrada: {planilha}")

        return ler_planilha(planilha)

    parser.error(
        "Informe --registros, uma planilha ou --mat MATRICULA ANO_INICIAL-ANO_FINAL."
    )


def executar(dados):
    """Processa todos os registros. Devolve (sucesso, mensagem_de_erro)."""

    print()
    print("=" * 80)
    print("REGISTROS ENCONTRADOS")
    print("=" * 80)

    for matricula, anos in dados:
        print(f"Matrícula: {matricula} | Anos: {anos[0]} até {anos[-1]}")

    # ------------------------------ ambiente inicial -------------------------------- 
    if not abrir_jms():
        return False, "Falhou ao abrir o JMS."

    if not login():
        return False, "Login no Folha não concluído."

    if not abrir_janela_interna():
        return False, "Abrir janela interna resultou em erro."

    # ------------------------ ciclo de matrículas ------------------------ 

    # identifica a janela interna
    if not verificar_janela_interna(TITULO_INTERNO, TITULO_JMS):
        return False, "Janela interna não identificada"
    
    for indice_matricula, (matricula, anos) in enumerate(dados, start=1):

        print()
        print("=" * 80)
        print(f"MATRÍCULA {indice_matricula}/{len(dados)}")
        print("=" * 80)
        print(f"Matrícula: {matricula}")
        print(f"Anos: {anos[0]} até {anos[-1]}")

        for indice_ano, ano in enumerate(anos, start=1):

            print()
            print("=" * 80)
            print(f"PROCESSANDO ANO {indice_ano}/{len(anos)}")
            print("=" * 80)
            print(f"Matrícula: {matricula}")
            print(f"Ano: {ano}")

            gerarFicha(matricula, ano)

            if not tratar_saida_impressao():
                return False, (
                    f"Erro ao gerar a saída de impressão "
                    f"(matrícula {matricula}, ano {ano})."
                )

            print()
            print("SAÍDA DE IMPRESSÃO CONCLUÍDA")

            if not organizar_pdf_temporario():
                return False, (
                    f"Erro ao organizar o PDF "
                    f"(matrícula {matricula}, ano {ano})."
                )

            print()
            print("PDF ORGANIZADO COM SUCESSO.")

            if indice_ano < len(anos):
                if not avancar_para_proximo_ano():
                    input("Falhou no sistema de avançar ano. ")
            else:
                print()
                print("TODOS OS ANOS DA MATRÍCULA FORAM PROCESSADOS.")

                if indice_matricula < len(dados):
                    preparar_proxima_matricula()

        if indice_matricula < len(dados):
            time.sleep(1)

    print()
    print("=" * 80)
    print("PROCESSAMENTO CONCLUÍDO")
    print("=" * 80)

    print()
    print("Fechando JMS")
    print()

    fechar_jms()

    return True, ""


def main():
    global PASTA_SERVIDORES, PASTA_DESTINO_FIXA, RETORNO

    parser = argparse.ArgumentParser(
        description="Gera fichas financeiras (2008 e anteriores) por registros, planilha ou matrícula."
    )

    parser.add_argument(
        "planilha",
        nargs="?",
        help="Caminho da planilha Excel com as matrículas e os períodos."
    )

    parser.add_argument(
        "--mat",
        nargs=2,
        metavar=("MATRICULA", "PERIODO"),
        help="Matrícula e intervalo de anos. Ex.: --mat 1234 2000-2005"
    )

    adicionar_argumentos(parser)

    args = parser.parse_args()  # entrada inválida: sai com código 2

    servidores = pasta_existente(parser, args.servidores, "--servidores")
    pasta = pasta_existente(parser, args.pasta, "--pasta")
    dados = obter_dados(args, parser)

    if servidores is not None:
        PASTA_SERVIDORES = servidores

    PASTA_DESTINO_FIXA = pasta
    RETORNO = Retorno(args.retorno)

    if not dados:
        print("Nenhum registro encontrado.")
        RETORNO.gravar(False, "Nenhum registro encontrado.")
        return 1
    
    # --------------------------- impressora ------------------------------- 
    if not impressora_pdf():
        input("Impressora não é a pdf")
    
    try:
        sucesso, erro = executar(dados)
    except Exception as exc:
        traceback.print_exc()
        sucesso, erro = False, f"Erro inesperado: {exc}"

    if not sucesso:
        print()
        print(f"ERRO: {erro}", file=sys.stderr)
        RETORNO.falha(erro)

    RETORNO.gravar(sucesso, erro)

    return 0 if sucesso else 1


if __name__ == "__main__":

    sys.exit(main())