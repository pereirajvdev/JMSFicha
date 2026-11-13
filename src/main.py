"""
Gera fichas financeiras (2008 e anteriores) a partir de registros, planilha
ou matrícula, automatizando o sistema Folha e organizando os PDFs gerados.

Versão refatorada: a lógica, a sequência de automação, os tempos de espera,
os argumentos de CLI e os destinos dos arquivos foram preservados.
"""

import argparse
import ctypes
import re
import shutil
import subprocess
import sys
import time
import traceback
from ctypes import wintypes
from functools import wraps
from pathlib import Path

# A ordem destes imports foi mantida de propósito (pyautogui e pywinauto
# configuram o DPI do processo ao serem importados).
import pyautogui
import fitz
import psutil
import win32print

from contrato import adicionar_argumentos, pasta_existente, registros_para_dados, Retorno

from openpyxl import load_workbook
from pywinauto import Desktop
import pyperclip


# ======================================================================
# CONFIGURAÇÕES
# ======================================================================
EXE = Path(r"C:\Folha\folha.exe")

RAIZ_PROJETO = Path(__file__).resolve().parents[1]   # se o arquivo está em src/
CRED = RAIZ_PROJETO / "data" / "credencial.txt"

TITULO_INICIAL = "Form1"
TITULO_JMS = "Folha - PREFEITURA MUNICIPAL DE TRES CORACOES - Abril de 2009"
TITULO_INTERNO = "Form1"
TITULO_EXTERNO_IMPRIMIR = "Imprimir"
TITULO_EXTERNO_PRINT = "Print"
TITULO_IMPRESSAO = "Salvar Saída de Impressão como"
TITULO_REPORT = "Report Designer - fichafin.frx - Folha - PREFEITURA MUNICIPAL DE TRES CORACOES - Abril de 2009"

PASTA_DOWNLOADS = Path.home() / "Downloads"
PASTA_SERVIDORES = Path(r"C:\Users\Joao Castro\Desktop\JOAO\Servidores")

NOME_TEMPORARIO = "temporario.pdf"
CAMINHO_TEMPORARIO = PASTA_DOWNLOADS / NOME_TEMPORARIO

TEMPO_LIMITE = 60
INTERVALO_VERIFICACAO = 1

# Teto de tentativas em fechar_jms (evita laço infinito se a janela persistir).
MAX_TENTATIVAS_FECHAR = 20

# Preenchidos em main()
PASTA_DESTINO_FIXA = None
RETORNO = None


# ======================================================================
# DESEMPENHO (--perf)
# ======================================================================
PERF_ATIVO = False
_TEMPOS = {}


def _registrar_tempo(nome, duracao):
    registro = _TEMPOS.setdefault(nome, [0, 0.0])
    registro[0] += 1
    registro[1] += duracao


class _SemMedicao:
    """Context manager nulo, compartilhado: custo praticamente zero com --perf desligado."""
    __slots__ = ()

    def __enter__(self):
        return None

    def __exit__(self, *_exc):
        return False


class _Medicao:
    __slots__ = ("nome", "inicio")

    def __init__(self, nome):
        self.nome = nome
        self.inicio = 0.0

    def __enter__(self):
        self.inicio = time.perf_counter()
        return self

    def __exit__(self, *_exc):
        _registrar_tempo(self.nome, time.perf_counter() - self.inicio)
        return False


_SEM_MEDICAO = _SemMedicao()


def medir(nome):
    """Mede o tempo de um bloco (`with medir(...)`). Sem efeito se --perf estiver desligado."""
    return _Medicao(nome) if PERF_ATIVO else _SEM_MEDICAO


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


def _banner(texto):
    """Imprime o cabeçalho padrão de seção (linha em branco + moldura)."""
    print()
    print("=" * 80)
    print(texto)
    print("=" * 80)


# ======================================================================
# WIN32 (ctypes): enumeração rápida de janelas
# ======================================================================
_user32 = ctypes.windll.user32

_WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

_user32.EnumWindows.argtypes = [_WNDENUMPROC, wintypes.LPARAM]
_user32.EnumWindows.restype = wintypes.BOOL
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


def _buscar_janela(titulos):
    """
    Faz UMA enumeração e retorna (hwnd, titulo_real) da primeira janela
    válida, respeitando a ordem dos títulos informados.
    Comparação: texto idêntico (sem diferenciar maiúsculas/minúsculas).
    Janelas travadas são ignoradas.
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
    with medir("criar wrapper pywinauto"):
        return _DESKTOP.window(handle=hwnd).wrapper_object()


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


def verificar_janela_interna(titulo_interno, titulo_janela):
    """Procura um controle com o texto `titulo_interno` dentro da janela."""
    print(f"Aguardando: {titulo_interno} from {titulo_janela}")

    janela = exigir_janela(titulo_janela)

    for i, controle in enumerate(janela.descendants()):
        try:
            if titulo_interno == controle.window_text():
                print(f"Encontrado: {titulo_interno} from {titulo_janela}")
                return True

        except Exception as e:
            print(f"[{i}] Erro: {e}")

    return False


def _exigir_controle(janela, titulo, controles=None):
    """Como procurar_controle, mas levanta RuntimeError (mensagem clara) se não achar."""
    controle = procurar_controle(janela, titulo, controles=controles)

    if controle is None:
        raise RuntimeError(f"Controle '{titulo}' não encontrado na janela.")

    return controle


def procurar_controle(janela, titulo=None, classe=None, controles=None):
    """
    Devolve o primeiro descendente com o texto `titulo` ou a classe `classe`.
    `controles` permite reutilizar uma lista de descendants() já obtida.
    """
    if controles is None:
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

    return None


# ======================================================================
# LOGIN / ABERTURA DO SISTEMA
# ======================================================================
def _janela_folha_aberta():
    return aguardar_janela(TITULO_INICIAL, TITULO_JMS, TITULO_REPORT, time_limit=.5)


def fechar_jms():
    """Encerra o processo associado à janela principal, se estiver aberto."""

    print("[INFO] Verificando se a folha já está aberta...")

    janela = _janela_folha_aberta()
    tentativas = 0

    while janela is not None:
        tentativas += 1
        if tentativas > MAX_TENTATIVAS_FECHAR:
            raise RuntimeError(
                "Não foi possível encerrar a instância anterior da folha "
                f"após {MAX_TENTATIVAS_FECHAR} tentativas."
            )

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

        janela = _janela_folha_aberta()


@cronometrar("abrir folha")
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

    janela.move_window(
        x=x,
        y=0,
        width=metade,
        height=altura_tela,
        repaint=True
    )

    return True


@cronometrar("abrir janela interna")
def abrir_janela_interna():
    janela = aguardar_janela(TITULO_JMS, time_limit=3)
    if janela is None:  # passa pela janela de aviso com mesmo nome: Form1
        pyautogui.press("enter")
        janela = exigir_janela(TITULO_JMS, time_limit=10)
    janela.set_focus()

    reposicionar(janela, "left")
    time.sleep(1)

    pyautogui.hotkey("alt")
    time.sleep(.02)

    pyautogui.hotkey("r")
    time.sleep(.02)

    pyautogui.press("down", presses=6, interval=.001)
    time.sleep(.02)

    pyautogui.press("right")
    time.sleep(.02)

    pyautogui.press("down", presses=9, interval=.001)
    time.sleep(.02)

    pyautogui.press("enter")
    time.sleep(.02)

    return True


@cronometrar("login")
def login():
    _banner("INICIANDO LOGIN")

    usuario, senha = CRED.read_text(encoding="utf-8").splitlines()[:2]

    janela = exigir_janela(TITULO_INICIAL)
    janela.set_focus()

    print("Preenchendo usuário...")

    pyautogui.write(usuario, interval=.001)
    pyautogui.press("tab")
    time.sleep(.005)

    pyautogui.write(senha, interval=.001)
    pyautogui.press("tab")
    time.sleep(.005)

    pyautogui.press("enter")

    print("Login concluído.")

    return True


# ======================================================================
# LEITURA DA PLANILHA
# ======================================================================
def ler_planilha(caminho):
    """Devolve [(matricula, [anos ordenados]), ...] a partir da planilha."""
    wb = load_workbook(caminho, data_only=True)

    try:
        ws = wb.active
        registros = {}

        for linha in ws.iter_rows(min_row=2, values_only=True):
            matricula, data_inicial, data_final = linha[:3]

            if not matricula or not data_inicial or not data_final:
                continue

            registros.setdefault(str(matricula), set()).update(
                range(data_inicial.year, data_final.year + 1)
            )
    finally:
        wb.close()

    return [(matricula, sorted(anos)) for matricula, anos in registros.items()]


# ======================================================================
# IMPRESSORA PDF
# ======================================================================
def definir_impressora_pdf():
    """Executa o AutoFicha no modo padrão."""
    script = r"C:\Users\Joao Castro\Documents\AutoFicha\src\main.py"

    subprocess.Popen(
        [sys.executable, script, "--padrao"],
        cwd=r"C:\Users\Joao Castro\Documents\AutoFicha"
    )


def impressora_eh_pdf():
    return win32print.GetDefaultPrinter().strip().lower() == "microsoft print to pdf"


def impressora_pdf():
    if impressora_eh_pdf():
        print("Impressora padrão já é Microsoft Print to PDF.")
    else:
        definir_impressora_pdf()

        if not impressora_eh_pdf():
            print("ATENÇÃO: Definir impressora falhou. ")
            print("ATENÇÃO: A impressora padrão não é Microsoft Print to PDF.")

        # Comportamento original preservado: quando a impressora não era PDF
        # na primeira verificação, o retorno é sempre False.
        return False

    return True


# ======================================================================
# GERAR FICHA
# ======================================================================
@cronometrar("gerar ficha (digitação)")
def gerarFicha(matricula, ano):
    # INSERE DADOS NA JANELA DESENHADA, NAO EH POSSIVEL DAR controle.set_focus()
    time.sleep(1)

    pyautogui.write(str(ano), interval=.001)
    pyautogui.press("tab")
    time.sleep(.02)

    pyautogui.write(matricula, interval=.001)
    pyautogui.press("tab")
    time.sleep(.02)

    pyautogui.write(matricula, interval=.001)
    time.sleep(.02)

    pyautogui.press("tab", presses=9, interval=.001)
    time.sleep(.02)

    pyautogui.press("enter")
    time.sleep(.02)

    pyautogui.press("tab")
    time.sleep(.02)

    pyautogui.press("enter")


# ======================================================================
# SALVAMENTO DE FICHA
# ======================================================================
def tratar_externo_imprimir(janela):
    print("Iniciando salvamento pelo modo externo: IMPRIMIR")

    _exigir_controle(janela, "&Sim").click()

    salvou = salvar_temporario()
    if not salvou:
        input("Erro no salvamento")

    salvou_report = tratar_interno_print()

    janela.set_focus()
    botao_ok = janela.child_window(
        title="&Sim",
        class_name="Button"
    )

    print(f"botao ok: {botao_ok}")
    botao_ok.set_focus()
    botao_ok.click()

    # Retorna se o salvamento foi confirmado (usado só para log).
    return salvou and salvou_report is not False


def tratar_externo_print(janela):
    print("Iniciando salvamento pelo modo externo: PRINT")

    ok_btn = procurar_controle(janela, "OK")

    if ok_btn is None:
        input("Botao de ok nao identificado. ")
        raise RuntimeError("Botão 'OK' não identificado na janela 'Print'.")

    janela.set_focus()
    ok_btn.click()

    salvou = salvar_temporario()
    if not salvou:
        input("Erro no salvamento")

    return salvou


def _report_designer_aberto():
    """True se existir janela visível cujo título contém 'Report Designer'."""
    return any(
        "Report Designer" in titulo
        for _hwnd, titulo in _listar_janelas_visiveis()
    )


def tratar_interno_print():
    print("Iniciando salvamento pelo modo interno: REPORT DESIGN")

    try:
        if _report_designer_aberto():
            print()
            print("REPORT DESIGNER ENCONTRADO.")

            janela = exigir_janela(TITULO_REPORT)
            janela.set_focus()

            pyautogui.doubleClick(373, 325, interval=0.1)

            time.sleep(0.5)

            salvou = salvar_temporario()
            if not salvou:
                input("Erro no salvamento")

            return salvou

    except Exception as e:
        print(f"Erro procurando Report Designer: {e}")
        return False

    return None  # Report Designer não estava aberto: nada foi tentado


@cronometrar("espera pela impressão (detectar modo)")
def verificar_modo():
    """Identifica o modo de impressão. Retorna (modo, janela)."""
    janela = aguardar_janela(
        TITULO_EXTERNO_PRINT,
        TITULO_EXTERNO_IMPRIMIR,
        time_limit=7
    )

    if janela is None:
        return "interno_print", janela

    texto = janela.window_text()

    if texto == TITULO_EXTERNO_IMPRIMIR:
        return "externo_imprimir", janela
    if texto == TITULO_EXTERNO_PRINT:
        return "externo_print", janela

    return None, None


@cronometrar("impressão (saída de impressão)")
def tratar_saida_impressao():
    _banner("AGUARDANDO SAÍDA DE IMPRESSÃO")

    modo, janela = verificar_modo()

    if janela is not None:
        janela.set_focus()

    confirmado = False

    if modo == "externo_imprimir":
        confirmado = tratar_externo_imprimir(janela)
    elif modo == "externo_print":
        confirmado = tratar_externo_print(janela)
    elif modo == "interno_print":
        confirmado = tratar_interno_print() is True
    else:
        input("Nenhum modo foi selecionado")

    if not confirmado:
        print(
            "[AVISO] Salvamento do PDF não confirmado nesta etapa; "
            "a validação do PDF temporário decidirá o resultado."
        )

    # Contrato original preservado: sempre True; a falha é detectada
    # depois, em organizar_pdf_temporario().
    return True


# ======================================================================
# PDF TEMPORÁRIO
# ======================================================================
def deletar_temporario():
    _banner("VERIFICANDO PDF TEMPORÁRIO")

    try:
        CAMINHO_TEMPORARIO.unlink()
        print(f"PDF temporário anterior deletado: {CAMINHO_TEMPORARIO}")

    except FileNotFoundError:
        print("Nenhum PDF temporário anterior encontrado.")

    except Exception as e:
        print(f"ERRO ao deletar PDF temporário: {e}")
        return False

    return True


def _remover_temporario_obsoleto():
    """
    Remove (sem banner) um temporario.pdf deixado por execução anterior, para
    que ele nunca seja organizado como se fosse a ficha atual caso o salvamento
    seja pulado.
    """
    try:
        CAMINHO_TEMPORARIO.unlink()
        print(f"[INFO] PDF temporário obsoleto removido: {CAMINHO_TEMPORARIO}")

    except FileNotFoundError:
        pass

    except Exception as e:
        print(f"[AVISO] Não foi possível remover o PDF temporário obsoleto: {e}")


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


def _localizar_campo_nome(controles):
    """Campo 'Edit' do nome do arquivo (o único com largura > 500)."""
    for controle in controles:
        try:
            if controle.class_name() != "Edit":
                continue

            if controle.rectangle().width() > 500:
                return controle
        except Exception as e:
            print(f"[AVISO] Controle ignorado ao procurar o campo do nome: {e}")

    return None


@cronometrar("validar PDF (aguardar gravação)")
def _aguardar_pdf_pronto(caminho, limite_s=15, intervalo=0.3, leituras_estaveis=3):
    """
    Aguarda o PDF ficar com tamanho estável e abrir no PyMuPDF.
    Mesmas regras do original: 15 s no total, verificação a cada 0,3 s,
    3 leituras consecutivas com o mesmo tamanho.
    """
    print()
    print("Aguardando gravação completa do PDF...")

    limite = time.monotonic() + limite_s

    ultimo_tamanho = -1
    ultimo_impresso = None
    tamanho_estavel = 0

    while time.monotonic() < limite:
        try:
            try:
                tamanho = caminho.stat().st_size
            except FileNotFoundError:
                tamanho = None

            if tamanho is not None:
                if tamanho != ultimo_impresso:  # evita repetir a mesma linha a cada 0,3 s
                    print(f"PDF encontrado. Tamanho: {tamanho:,} bytes")
                    ultimo_impresso = tamanho

                if tamanho > 0:
                    if tamanho == ultimo_tamanho:
                        tamanho_estavel += 1
                    else:
                        tamanho_estavel = 0

                    ultimo_tamanho = tamanho

                    if tamanho_estavel >= leituras_estaveis:
                        try:
                            with medir("abrir PDF no PyMuPDF (validação)"):
                                with fitz.open(caminho) as documento:
                                    paginas = documento.page_count

                            print()
                            print("PDF temporário gravado e validado com sucesso.")
                            print(f"Páginas: {paginas}")
                            print(f"Arquivo: {caminho}")
                            return True

                        except Exception as e:
                            print(
                                "PDF existe, mas ainda está "
                                f"sendo finalizado: {e}"
                            )

        except Exception as e:
            print(f"Erro verificando PDF: {e}")

        time.sleep(intervalo)

    print()
    print("ERRO: PDF temporário não ficou disponível para leitura.")

    try:
        print(f"Tamanho final encontrado: {caminho.stat().st_size:,} bytes")
    except FileNotFoundError:
        pass
    except Exception as e:
        print(f"Não foi possível ler o tamanho final: {e}")

    return False


@cronometrar("salvar temporário")
def salvar_temporario():
    if not deletar_temporario():
        return False

    janela = exigir_janela(TITULO_IMPRESSAO)
    janela.set_focus()

    print("Arquivo:", NOME_TEMPORARIO)
    print("Caminho:", CAMINHO_TEMPORARIO)

    # Uma única enumeração de descendants() serve ao campo e ao botão.
    with medir("localizar controles do diálogo"):
        controles = janela.descendants()
        campo_nome = _localizar_campo_nome(controles)

    if campo_nome is None:
        print("Campo do nome do arquivo não encontrado.")
        return False

    colar_texto(campo_nome, CAMINHO_TEMPORARIO)

    print("Nome do arquivo preenchido.")

    botao_salvar = procurar_controle(janela, "Sa&lvar", controles=controles)
    if botao_salvar is None:
        # Rede de segurança: nova enumeração (comportamento original).
        botao_salvar = _exigir_controle(janela, "Sa&lvar")

    botao_salvar.set_focus()
    botao_salvar.click()

    return _aguardar_pdf_pronto(CAMINHO_TEMPORARIO)


# ======================================================================
# EXTRAÇÃO DE DADOS DO SERVIDOR
# ======================================================================
_RE_MATRICULA_NOME = re.compile(r"(?m)^Matricula:\s*(\d+)\s+(.+?)\s*$", re.IGNORECASE)
_RE_CARGO = re.compile(r"(?m)^Cargo:\s*(.+?)\s*$", re.IGNORECASE)
_RE_ANO = re.compile(r"Relatório\s+da\s+Ficha\s+Financeira\s+(\d{4})", re.IGNORECASE)
_RE_CHARS_INVALIDOS = re.compile(r'[<>:"/\\|?*]')
_RE_ESPACOS = re.compile(r"\s+")


def extrair_dados_servidor_modelo_antigo(texto: str) -> dict:
    """Extrai matrícula, nome, cargo e ano da ficha financeira (modelo antigo)."""
    dados = {}

    padrao = _RE_MATRICULA_NOME.search(texto)
    if padrao:
        dados["Matrícula"] = padrao.group(1)
        dados["Nome"] = padrao.group(2).strip()

    padrao = _RE_CARGO.search(texto)
    if padrao:
        dados["Cargo"] = padrao.group(1).strip()

    padrao = _RE_ANO.search(texto)
    if padrao:
        dados["Ano"] = int(padrao.group(1))

    return dados


@cronometrar("extrair texto do PDF")
def extrair_texto_pdf(caminho_pdf):
    _banner("EXTRAINDO DADOS DO PDF")

    print(f"PDF: {caminho_pdf}")

    try:
        with fitz.open(caminho_pdf) as documento:
            texto = "\n".join(pagina.get_text() for pagina in documento)

        print(f"Caracteres extraídos: {len(texto)}")

        return texto

    except Exception as e:
        print(f"ERRO ao extrair texto do PDF: {e}")

        return None


def sanitizar_nome(nome):
    """Remove caracteres inválidos para nome de pasta e espaços duplicados."""
    nome = _RE_CHARS_INVALIDOS.sub("", nome)
    nome = _RE_ESPACOS.sub(" ", nome)

    return nome.strip()


# ======================================================================
# ORGANIZAÇÃO DO PDF
# ======================================================================
@cronometrar("organizar arquivo (total, inclui extração)")
def organizar_pdf_temporario():
    _banner("ORGANIZANDO FICHA FINANCEIRA")

    print(f"Temporário: {CAMINHO_TEMPORARIO}")

    if not CAMINHO_TEMPORARIO.exists():
        print("ERRO: PDF temporário não encontrado.")
        return False

    texto = extrair_texto_pdf(CAMINHO_TEMPORARIO)

    if not texto:
        print("ERRO: não foi possível extrair texto do PDF.")
        return False

    dados = extrair_dados_servidor_modelo_antigo(texto)

    print()
    print("DADOS EXTRAÍDOS:")

    for chave, valor in dados.items():
        print(f"{chave}: {valor}")

    nome_servidor = dados.get("Nome")

    if not nome_servidor:
        print("ERRO: nome do servidor não encontrado.")
        return False

    ano = dados.get("Ano")

    if not ano:
        print("ERRO: ano da ficha financeira não encontrado.")
        return False

    nome_servidor = sanitizar_nome(nome_servidor)

    if not nome_servidor:
        # Sem esta guarda, o PDF iria para a raiz de PASTA_SERVIDORES.
        print("ERRO: nome do servidor vazio após a sanitização.")
        return False

    pasta_servidor = PASTA_SERVIDORES / nome_servidor

    try:
        pasta_servidor.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        print(f"ERRO ao criar a pasta do servidor: {e}")
        return False

    print()
    print(f"Pasta do servidor: {pasta_servidor}")

    destino = pasta_servidor / f"FICHA FINANCEIRA {ano}.pdf"

    print(f"Destino: {destino}")

    try:
        with medir("mover PDF"):
            shutil.move(str(CAMINHO_TEMPORARIO), str(destino))

        if not destino.is_file():
            print()
            print("ERRO: o PDF não foi encontrado no destino após a movimentação.")
            return False

        print()
        print("PDF ORGANIZADO COM SUCESSO.")
        print(f"Arquivo: {destino}")

        return True

    except Exception as e:
        print()
        print(f"ERRO ao mover PDF: {e}")

        return False


# ======================================================================
# NAVEGAÇÃO ENTRE ANOS / MATRÍCULAS
# ======================================================================
def avancar_para_proximo_ano():
    _banner("AVANÇANDO PARA O PRÓXIMO ANO")

    for i in range(3):
        pyautogui.press("tab")
        print(f"TAB {i + 1}/3")
        time.sleep(0.05)

    print("Pronto para o próximo ano.")

    return True


def preparar_proxima_matricula():
    _banner("PREPARANDO PRÓXIMA MATRÍCULA")

    time.sleep(0.5)

    # sequência para retornar ao campo inicial
    pyautogui.press("tab", presses=3, interval=0.02)

    print("Pronto para a próxima matrícula.")

    return True


# ======================================================================
# MAIN
# ======================================================================
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
        if ano_inicial < 2001:
            parser.error("O ano inicial não pode ser menor que 2001.")

        return [(matricula, list(range(ano_inicial, ano_final + 1)))]

    if args.planilha:
        planilha = Path(args.planilha)

        if not planilha.is_file():
            parser.error(f"Planilha não encontrada: {planilha}")

        return ler_planilha(planilha)

    parser.error(
        "Informe --registros, uma planilha ou --mat MATRICULA ANO_INICIAL-ANO_FINAL."
    )


def _processar_ano(matricula, ano):
    """Gera, salva e organiza a ficha de um ano. Devolve mensagem de erro ou None."""
    _remover_temporario_obsoleto()

    gerarFicha(matricula, ano)

    if not tratar_saida_impressao():
        return (
            f"Erro ao gerar a saída de impressão "
            f"(matrícula {matricula}, ano {ano})."
        )

    print()
    print("SAÍDA DE IMPRESSÃO ENCERRADA (PDF será validado a seguir)")

    if not organizar_pdf_temporario():
        return (
            f"Erro ao organizar o PDF "
            f"(matrícula {matricula}, ano {ano})."
        )

    print()
    print("PDF ORGANIZADO COM SUCESSO.")

    return None


def executar(dados):
    """Processa todos os registros. Devolve (sucesso, mensagem_de_erro)."""

    _banner("REGISTROS ENCONTRADOS")

    for matricula, anos in dados:
        print(f"Matrícula: {matricula} | Anos: {anos[0]} até {anos[-1]}")

    # ------------------------------ ambiente inicial --------------------------------
    with medir("inicialização (abrir, login, janela interna)"):
        if not abrir_jms():
            return False, "Falhou ao abrir o JMS."

        if not login():
            return False, "Login no Folha não concluído."

        if not abrir_janela_interna():
            return False, "Abrir janela interna resultou em erro."

        # identifica a janela interna
        with medir("verificar janela interna"):
            janela_interna_ok = verificar_janela_interna(TITULO_INTERNO, TITULO_JMS)

        if not janela_interna_ok:
            return False, "Janela interna não identificada"

    # ------------------------ ciclo de matrículas ------------------------
    total_matriculas = len(dados)

    for indice_matricula, (matricula, anos) in enumerate(dados, start=1):

        _banner(f"MATRÍCULA {indice_matricula}/{total_matriculas}")
        print(f"Matrícula: {matricula}")
        print(f"Anos: {anos[0]} até {anos[-1]}")

        with medir("matrícula (total)"):

            for indice_ano, ano in enumerate(anos, start=1):

                _banner(f"PROCESSANDO ANO {indice_ano}/{len(anos)}")
                print(f"Matrícula: {matricula}")
                print(f"Ano: {ano}")

                with medir("ano (total)"):
                    erro = _processar_ano(matricula, ano)

                    if erro:
                        return False, erro

                    if indice_ano < len(anos):
                        if not avancar_para_proximo_ano():
                            input("Falhou no sistema de avançar ano. ")
                    else:
                        print()
                        print("TODOS OS ANOS DA MATRÍCULA FORAM PROCESSADOS.")

                        if indice_matricula < total_matriculas:
                            preparar_proxima_matricula()

        if indice_matricula < total_matriculas:
            time.sleep(1)

    _banner("PROCESSAMENTO CONCLUÍDO")

    print()
    print("Fechando JMS")
    print()

    fechar_jms()

    return True, ""


def main():
    global PASTA_SERVIDORES, PASTA_DESTINO_FIXA, RETORNO, PERF_ATIVO

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

    try:
        parser.add_argument(
            "--perf",
            action="store_true",
            help="Exibe o relatório de desempenho ao final."
        )
    except argparse.ArgumentError:
        pass  # já definido em contrato.adicionar_argumentos

    args = parser.parse_args()  # entrada inválida: sai com código 2

    PERF_ATIVO = bool(getattr(args, "perf", False))

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
        with medir("execução total"):
            sucesso, erro = executar(dados)
    except Exception as exc:
        traceback.print_exc()
        sucesso, erro = False, f"Erro inesperado: {exc}"

    if not sucesso:
        print()
        print(f"ERRO: {erro}", file=sys.stderr)
        RETORNO.falha(erro)

    RETORNO.gravar(sucesso, erro)

    imprimir_relatorio_tempos()

    return 0 if sucesso else 1


if __name__ == "__main__":
    sys.exit(main())