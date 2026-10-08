import sys
import time
import win32gui
import pyautogui
import win32print
import subprocess
import re
import shutil
import fitz


from pathlib import Path
from pywinauto import Desktop
from openpyxl import load_workbook


EXE = Path(r"C:\Folha\folha.exe")

CRED = Path("data/credencial.txt")


# ============================================================
# CONFIGURAÇÕES
# ============================================================

TITULO_GERADOR = (
    "Gerador de Relatórios "
    "( Versão: 4.0 ) - Recursos Humanos e Folha de Pagamento"
)

PASTA_DOWNLOADS = Path.home() / "Downloads"

PASTA_SERVIDORES = Path(
    r"C:\Users\Joao Castro\Desktop\JOAO\Servidores"
)

NOME_TEMPORARIO = "temporario.pdf"

# ============================================================
# LOGIN
# ============================================================

def login():

    print()
    print("=" * 80)
    print("INICIANDO LOGIN")
    print("=" * 80)

    print(f"EXE: {EXE}")
    print(f"EXE existe: {EXE.exists()}")
    print(f"Diretório: {EXE.parent}")
    print(f"Diretório existe: {EXE.parent.exists()}")

    if not EXE.exists():
        print()
        print("ERRO: executável do Folha não encontrado.")
        print(f"Caminho procurado: {EXE}")
        return False

    usuario, senha = CRED.read_text(
        encoding="utf-8"
    ).splitlines()[:2]

    print("Abrindo Folha...")

    subprocess.Popen(
        [str(EXE)],
        cwd=str(EXE.parent)
    )

    time.sleep(.5)

    print("Preenchendo usuário...")

    pyautogui.write(
        usuario,
        interval=.05
    )

    pyautogui.press("tab")
    time.sleep(.02)

    pyautogui.write(
        senha,
        interval=.05
    )

    pyautogui.press("tab")
    time.sleep(.02)

    pyautogui.press("enter")
    time.sleep(.4)

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

    pyautogui.hotkey("win", "left")
    time.sleep(.02)

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

    dados = []

    for linha in ws.iter_rows(
        min_row=2,
        values_only=True
    ):

        matricula, data_inicial, data_final = linha[:3]

        if not matricula or not data_inicial or not data_final:
            continue

        anos = list(
            range(
                data_inicial.year,
                data_final.year + 1
            )
        )

        dados.append(
            (
                str(matricula),
                anos
            )
        )

    return dados


# ============================================================
# FOLHA INÍCIO
# ============================================================

def folhaInicio():

    SCRIPT = Path(
        r"C:\Users\Joao Castro\Documents\FolhaInicio\src\main.py"
    )

    print()
    print("=" * 80)
    print("EXECUTANDO FOLHA INÍCIO")
    print("=" * 80)

    print(f"Script: {SCRIPT}")
    print(f"Existe: {SCRIPT.exists()}")
    print(f"Diretório: {SCRIPT.parent}")
    print(f"Diretório existe: {SCRIPT.parent.exists()}")

    if not SCRIPT.exists():

        print(
            "ERRO: main.py do FolhaInicio não encontrado."
        )

        return None

    if not SCRIPT.parent.is_dir():

        print(
            "ERRO: diretório do FolhaInicio inválido."
        )

        return None

    processo = subprocess.Popen(
        ["python", str(SCRIPT)],
        cwd=str(SCRIPT.parent)
    )

    return processo


# ============================================================
# VERIFICA GERADOR
# ============================================================

def verificarGerador():

    print()
    print("=" * 80)
    print("VERIFICANDO GERADOR DE RELATÓRIOS")
    print("=" * 80)

    limite = time.time() + 10

    while time.time() < limite:

        try:

            hwnd_encontrado = None

            def procurar_janela(hwnd, _):

                nonlocal hwnd_encontrado

                if not win32gui.IsWindowVisible(hwnd):
                    return

                titulo = win32gui.GetWindowText(hwnd).strip()

                if titulo == TITULO_GERADOR:
                    hwnd_encontrado = hwnd

            win32gui.EnumWindows(
                procurar_janela,
                None
            )

            if hwnd_encontrado:

                print()
                print("GERADOR DE RELATÓRIOS ENCONTRADO")
                print(f"Handle : {hwnd_encontrado}")
                print(
                    f"Título : "
                    f"{win32gui.GetWindowText(hwnd_encontrado)}"
                )

                return True

        except Exception as e:

            print(
                f"Erro ao verificar Gerador: {e}"
            )

        time.sleep(.2)

    print()
    print(
        "ERRO: Gerador de Relatórios não foi encontrado."
    )

    return False


# ============================================================
# DEFINIR IMPRESSORA PDF
# ============================================================

def definir_impressora_pdf():

    MAX_TENTATIVAS = 3

    # ========================================================
    # TENTATIVAS DE ABRIR O GERADOR
    # ========================================================

    for tentativa in range(1, MAX_TENTATIVAS + 1):

        print()
        print("=" * 80)
        print(
            f"TENTATIVA {tentativa}/{MAX_TENTATIVAS}"
        )
        print("=" * 80)

        # ====================================================
        # INICIA FOLHA INÍCIO
        # ====================================================

        processo = folhaInicio()

        if processo is None:

            print(
                "FolhaInicio não pôde ser iniciado."
            )

            return

        codigo = processo.wait()

        print(
            f"FolhaInicio finalizado. Código: {codigo}"
        )

        if codigo != 0:

            print(
                "ATENÇÃO: FolhaInicio terminou com erro."
            )

            return

        # ====================================================
        # ABRE GERADOR
        # ====================================================

        time.sleep(.3)

        print(
            "Enviando sequência: ALT → E → 0 → S"
        )

        pyautogui.press("alt")
        time.sleep(.05)

        pyautogui.press("e")
        time.sleep(.05)

        pyautogui.press("0")
        time.sleep(.05)

        pyautogui.press("s")
        time.sleep(.5)

        print(
            "Sequência concluída."
        )

        # ====================================================
        # VERIFICA SE O GERADOR ABRIU
        # ====================================================

        if verificarGerador():

            print()
            print(
                "GERADOR CONFIRMADO!"
            )

            print(
                "Continuando para o AutoFicha..."
            )

            break

        # ====================================================
        # NÃO ABRIU
        # ====================================================

        print()
        print(
            "Gerador não foi confirmado."
        )

        if tentativa < MAX_TENTATIVAS:

            print(
                "Retornando para a tela inicial..."
            )

            # -----------------------------------------------
            # FolhaInicio novamente
            # -----------------------------------------------

            processo = folhaInicio()

            if processo is not None:

                codigo = processo.wait()

                print(
                    f"FolhaInicio finalizado. Código: {codigo}"
                )

                if codigo != 0:

                    print(
                        "ATENÇÃO: erro ao retornar "
                        "para a tela inicial."
                    )

                    return

            time.sleep(.5)

            print(
                "Nova tentativa será realizada..."
            )

        else:

            print()
            print(
                "=" * 80
            )
            print(
                "FALHA: não foi possível abrir "
                "o Gerador de Relatórios."
            )
            print(
                "=" * 80
            )

            return

    # ========================================================
    # EXECUTA AUTOFICHA
    # ========================================================

    DIRETORIO_AUTOFICHA = Path(
        r"C:\Users\Joao Castro\Documents\AutoFicha"
    )

    SCRIPT_AUTOFICHA = (
        DIRETORIO_AUTOFICHA
        / "src"
        / "main.py"
    )

    print()
    print("=" * 80)
    print(
        "EXECUTANDO AUTOFICHA - MODO PADRÃO"
    )
    print("=" * 80)

    print(
        f"Script: {SCRIPT_AUTOFICHA}"
    )

    print(
        f"Existe: {SCRIPT_AUTOFICHA.exists()}"
    )

    if not SCRIPT_AUTOFICHA.exists():

        print(
            "ERRO: main.py do AutoFicha não encontrado."
        )

        return

    processo = subprocess.Popen(
        [
            "python",
            str(SCRIPT_AUTOFICHA),
            "--padrao"
        ],
        cwd=str(DIRETORIO_AUTOFICHA)
    )

    codigo = processo.wait()

    print(
        f"AutoFicha finalizado. Código: {codigo}"
    )

    if codigo != 0:

        print(
            "ATENÇÃO: AutoFicha terminou com erro."
        )

        return

    print(
        "AutoFicha concluído."
    )

    # ========================================================
    # RETORNA PARA FOLHA INÍCIO
    # ========================================================

    print()
    print(
        "Retornando para FolhaInicio..."
    )

    folhaInicio()


# ============================================================
# GERAR FICHA
# ============================================================

def gerarFicha(matricula, ano):

    time.sleep(1)

    pyautogui.write(
        str(ano),
        interval=.05
    )

    pyautogui.press("tab")
    time.sleep(.02)

    pyautogui.write(
        matricula,
        interval=.05
    )

    pyautogui.press("tab")
    time.sleep(.02)

    pyautogui.write(
        matricula,
        interval=.05
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
# VERIFICA JANELA PRINT
# ============================================================

def verificar_print():

    print()
    print("=" * 80)
    print("VERIFICANDO JANELA PRINT")
    print("=" * 80)

    limite = time.time() + 10

    while time.time() < limite:

        try:

            hwnd_encontrado = None

            def procurar_janela(hwnd, _):

                nonlocal hwnd_encontrado

                if not win32gui.IsWindowVisible(hwnd):
                    return

                titulo = win32gui.GetWindowText(hwnd).strip()

                if titulo == "Print":
                    hwnd_encontrado = hwnd

            win32gui.EnumWindows(
                procurar_janela,
                None
            )

            if hwnd_encontrado:

                print()
                print("JANELA PRINT ENCONTRADA")
                print(
                    f"Handle : {hwnd_encontrado}"
                )
                print(
                    f"Título : "
                    f"{win32gui.GetWindowText(hwnd_encontrado)}"
                )

                return True

        except Exception as e:

            print(
                f"Erro ao verificar janela Print: {e}"
            )

        time.sleep(.2)

    print()
    print(
        "ERRO: janela Print não foi encontrada."
    )

    return False


# ============================================================
# AVANÇA NA JANELA PRINT
# ============================================================

def avancar_print():

    if not verificar_print():
        return False

    print()
    print("Janela Print confirmada.")
    print("Procurando botão OK...")

    try:
        desktop = Desktop(backend="win32")

        janela_print = desktop.window(
            class_name="#32770",
            title="Print"
        )

        if not janela_print.exists(timeout=3):
            print("ERRO: Janela Print não encontrada.")
            return False

        print(f"Janela Print encontrada. Handle: {janela_print.handle}")

        botao_ok = janela_print.child_window(
            title="OK",
            class_name="Button"
        )

        if not botao_ok.exists(timeout=3):
            print("ERRO: Botão OK não encontrado.")
            return False

        rect = botao_ok.rectangle()

        x = (rect.left + rect.right) // 2
        y = (rect.top + rect.bottom) // 2

        print("BOTÃO OK ENCONTRADO")
        print(f"Handle : {botao_ok.handle}")
        print(f"Rect   : {rect}")
        print(f"Centro : X={x}, Y={y}")

        # ----------------------------------------------------
        # TRAZ A JANELA PRINT PARA FRENTE
        # ----------------------------------------------------

        print("Ativando janela Print...")

        try:
            janela_print.set_focus()
        except Exception as e:
            print(f"Aviso ao ativar janela: {e}")

        time.sleep(0.5)

        # ----------------------------------------------------
        # MOVE O MOUSE ATÉ O BOTÃO
        # ----------------------------------------------------

        print("Movendo mouse até o botão OK...")

        pyautogui.moveTo(
            x,
            y,
            duration=0.02
        )

        time.sleep(0.05)

        # ----------------------------------------------------
        # DUPLO CLIQUE
        # ----------------------------------------------------

        print("Executando DUPLO CLIQUE no OK...")

        pyautogui.doubleClick(
            x,
            y,
            interval=0.15
        )

        print("Duplo clique enviado.")

        # ----------------------------------------------------
        # VERIFICA SE A JANELA FECHOU
        # ----------------------------------------------------

        print("Aguardando janela Salvar impressão...")

        if verificar_salvar_impressao():
            print("Salvar impressão confirmada.")
            return salvar_temporario()

        print(
            "ERRO: Salvar impressão não apareceu "
            "após clicar em OK."
        )

        return False

    except Exception as e:

        print(
            f"ERRO ao clicar no botão OK: {e}"
        )

        return False


def verificar_salvar_impressao():

    print()
    print("=" * 80)
    print("VERIFICANDO JANELA SALVAR IMPRESSÃO")
    print("=" * 80)

    limite = time.time() + 10

    while time.time() < limite:

        try:
            hwnd_encontrado = None

            def procurar_janela(hwnd, _):
                nonlocal hwnd_encontrado

                if not win32gui.IsWindowVisible(hwnd):
                    return

                titulo = win32gui.GetWindowText(hwnd).strip()

                if titulo == "Salvar Saída de Impressão como":
                    hwnd_encontrado = hwnd

            win32gui.EnumWindows(
                procurar_janela,
                None
            )

            if hwnd_encontrado:

                print()
                print("JANELA SALVAR IMPRESSÃO ENCONTRADA.")
                print(f"Handle : {hwnd_encontrado}")
                print(
                    f"Título : "
                    f"{win32gui.GetWindowText(hwnd_encontrado)}"
                )

                return True

        except Exception as e:

            print(
                f"Erro ao verificar Salvar impressão: {e}"
            )

        time.sleep(.2)

    print()
    print(
        "ERRO: janela Salvar impressão "
        "não foi encontrada."
    )

    return False


# ============================================================
# PROCURA PRINT OU PRINT PREVIEW
# ============================================================

def tratar_saida_impressao():
    print()
    print("=" * 80)
    print("AGUARDANDO SAÍDA DE IMPRESSÃO")
    print("=" * 80)

    limite = time.time() + 15

    while time.time() < limite:
        try:
            janelas = Desktop(backend="win32").windows()

            for janela in janelas:
                try:
                    if not janela.is_visible():
                        continue

                    titulo = janela.window_text()

                    if (
                        titulo.startswith("Report Designer - fichafin.frx")
                        and janela.class_name() == "folha7c000000"
                    ):
                        print()
                        print("REPORT DESIGNER ENCONTRADO.")
                        print("Executando clique em Print Preview...")

                        # ==================================================
                        # CLIQUE DIRETO NO PRINT PREVIEW
                        # ==================================================

                        pyautogui.doubleClick(
                            373,
                            325,
                            interval=0.15
                        )

                        print("Print Preview duplo clicado.")

                        if verificar_salvar_impressao():
                            print("Salvar impressão confirmada.")
                            return salvar_temporario()
                        
                        return aguardar_print_e_tab()

                except Exception:
                    pass

        except Exception as e:
            print(f"Erro procurando Report Designer: {e}")

        time.sleep(0.2)

    print()
    print("ERRO: Report Designer não foi encontrado.")
    return False


# ============================================================
# AGUARDA PRINT APÓS PRINT PREVIEW
# ============================================================

def aguardar_print_e_tab():

    print()
    print("Aguardando janela Print...")

    limite = time.time() + 10

    while time.time() < limite:

        try:

            encontrado = False

            def procurar(hwnd, _):

                nonlocal encontrado

                if not win32gui.IsWindowVisible(hwnd):
                    return

                titulo = win32gui.GetWindowText(hwnd).strip()

                if titulo == "Print":
                    encontrado = True

            win32gui.EnumWindows(
                procurar,
                None
            )

            if encontrado:
                print("JANELA PRINT ENCONTRADA.")

                return avancar_print()

        except Exception as e:

            print(
                f"Erro aguardando Print: {e}"
            )

        time.sleep(.2)

    print(
        "ERRO: Print não apareceu após o Print Preview."
    )

    return False


def clicar_print_preview():
    print()
    print("Procurando controle 'Print Preview'...")

    try:
        controle = Desktop(backend="win32").window(
            class_name="folha7c000000",
            title="Print Preview"
        )

        if not controle.exists(timeout=3):
            print("ERRO: Controle 'Print Preview' não encontrado.")
            return False

        rect = controle.rectangle()

        x = (rect.left + rect.right) // 2
        y = (rect.top + rect.bottom) // 2

        print("CONTROLE PRINT PREVIEW ENCONTRADO.")
        print(f"Classe : {controle.class_name()}")
        print(f"Rect   : {rect}")
        print(f"Clicando em X={x}, Y={y}...")

        # Garante que a janela esteja ativa
        controle.set_focus()

        time.sleep(0.2)

        # Clique físico do mouse
        pyautogui.click(x, y)

        print("Clique enviado.")

        return True

    except Exception as e:
        print(f"ERRO ao clicar no Print Preview: {e}")
        return False


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


def salvar_temporario():
    desktop = Desktop(backend="win32")

    if not deletar_temporario():
        return False

    print("Aguardando janela para salvar PDF...")

    janela = None

    while janela is None:
        for item in desktop.windows():
            try:
                if (
                    item.is_visible()
                    and item.class_name() == "#32770"
                    and item.window_text() == "Salvar Saída de Impressão como"
                ):
                    janela = item
                    break
            except Exception:
                pass

        if janela is None:
            time.sleep(0.2)

    print("Janela de salvamento encontrada.")
    print("Título:", janela.window_text())

    caminho_arquivo = PASTA_DOWNLOADS / NOME_TEMPORARIO

    print("Arquivo:", NOME_TEMPORARIO)
    print("Caminho:", caminho_arquivo)

    # ============================================================
    # CAMPO DO NOME DO ARQUIVO
    # ============================================================

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

    # ------------------------------------------------------------
    # LIMPA O CAMPO ANTES DE ESCREVER
    # ------------------------------------------------------------

    campo_nome.set_focus()
    time.sleep(0.3)

    campo_nome.click_input()
    time.sleep(0.2)

    campo_nome.select()
    time.sleep(0.2)

    # ------------------------------------------------------------
    # ESCREVE COM INTERVALO ENTRE OS CARACTERES
    # ------------------------------------------------------------

    print("Preenchendo caminho do arquivo...")

    campo_nome.type_keys(
        str(caminho_arquivo),
        with_spaces=True,
        pause=0.05
    )

    print("Nome do arquivo preenchido.")

    time.sleep(0.5)

    # ============================================================
    # CONFIRMA CAMPO
    # ============================================================

    try:
        texto_digitado = campo_nome.window_text()

        print(
            f"Texto atualmente no campo: {texto_digitado}"
        )

    except Exception:
        pass

    # ============================================================
    # BOTÃO SALVAR
    # ============================================================

    botao_salvar = None

    for controle in janela.descendants():
        try:
            if (
                controle.class_name() == "Button"
                and controle.window_text() == "Sa&lvar"
            ):
                botao_salvar = controle
                break

        except Exception:
            pass

    if botao_salvar is None:
        print("Botão Salvar não encontrado.")
        return False

    print()
    print("BOTÃO SALVAR ENCONTRADO.")
    print(f"Handle: {botao_salvar.handle}")

    # ============================================================
    # CLICA EM SALVAR
    # ============================================================

    print("Clicando em Salvar...")

    try:
        botao_salvar.click_input()
    except Exception:
        botao_salvar.click()

    print("Clique em Salvar realizado.")

    # ============================================================
    # CONFIRMA QUE O CLIQUE FOI PROCESSADO
    # ============================================================

    print("Confirmando fechamento da janela de salvamento...")

    limite_confirmacao = time.time() + 5

    while time.time() < limite_confirmacao:

        try:
            if not janela.exists():
                print("Janela de salvamento fechada.")
                print("Salvar confirmado.")
                break

        except Exception:
            # Se o controle deixou de existir,
            # também consideramos que a janela fechou.
            print("Janela de salvamento fechada.")
            print("Salvar confirmado.")
            break

        time.sleep(0.2)

    else:
        print()
        print(
            "ERRO: a janela de salvamento "
            "não fechou após o clique."
        )

        return False

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

    time.sleep(0.5)

    for i in range(3):
        pyautogui.press("tab")
        print(f"TAB {i + 1}/3")
        time.sleep(0.3)

    print("Pronto para o próximo ano.")

    return True


# ============================================================
# MAIN
# ============================================================

def main():

    if len(sys.argv) != 2:

        print(
            'Uso: python src/main.py "planilha.xlsx"'
        )

        return

    planilha = Path(
        sys.argv[1]
    )

    dados = ler_planilha(
        planilha
    )

    if not dados:

        print(
            "Nenhum registro encontrado."
        )

        return

    matricula, anos = dados[0]

    ano = anos[0]

    print(
        f"Matrícula: {matricula}"
    )

    print(
        f"Ano: {ano}"
    )

    impressora = win32print.GetDefaultPrinter()

    if impressora.strip().lower() == "microsoft print to pdf":
        print("Impressora padrão já é Microsoft Print to PDF.")
    else:
        definir_impressora_pdf()

        print("IMPRESSORA PADRÃO:")
        print(impressora)

        time.sleep(1)

    login()

    gerarFicha(
        matricula,
        ano
    )

    if tratar_saida_impressao():

        print("=" * 80)
        print("SAÍDA DE IMPRESSÃO CONCLUÍDA")
        print("=" * 80)

        if organizar_pdf_temporario():

            print()
            print("PDF ORGANIZADO COM SUCESSO.")

            # Volta para o campo correto para continuar o ciclo
            avancar_para_proximo_ano()

        else:

            print()
            print("ERRO AO ORGANIZAR O PDF.")
            return

    else:

        print()
        print("=" * 80)
        print("ERRO AO GERAR A SAÍDA DE IMPRESSÃO")
        print("=" * 80)
    

if __name__ == "__main__":

    main()