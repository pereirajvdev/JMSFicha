"""Configuração central, constantes e exceções do JMSFicha.

Tudo o que é caminho, título de janela ou tempo de espera fica aqui.
Os caminhos podem ser sobrescritos por variáveis de ambiente; sem elas,
valem os mesmos caminhos funcionais da versão anterior.
"""

from __future__ import annotations

import os
from pathlib import Path


# ======================================================================
# EXCEÇÕES
# ======================================================================
class ErroJMSFicha(RuntimeError):
    """Base dos erros previstos do projeto."""


class ErroConfiguracao(ErroJMSFicha):
    """Executável, credencial ou pasta ausente/inválida."""


class ErroAutomacao(ErroJMSFicha):
    """Falha ao conduzir a interface do sistema Folha."""


class JanelaNaoEncontrada(ErroAutomacao):
    """Janela esperada não apareceu (ou não responde)."""


class ControleNaoEncontrado(ErroAutomacao):
    """Controle esperado não existe na janela."""


class ErroPDF(ErroJMSFicha):
    """PDF ausente, incompleto, corrompido ou com dados inconsistentes."""


# ======================================================================
# CAMINHOS
# ======================================================================
def _caminho_env(variavel: str, padrao: Path) -> Path:
    valor = os.environ.get(variavel, "").strip()
    return Path(valor) if valor else padrao


RAIZ_PROJETO = Path(__file__).resolve().parents[1]  # este arquivo está em src/

EXE = _caminho_env("JMS_FOLHA_EXE", Path(r"C:\Folha\folha.exe"))
CRED = _caminho_env("JMS_CREDENCIAL", RAIZ_PROJETO / "data" / "credencial.txt")

PASTA_DOWNLOADS = Path.home() / "Downloads"
PASTA_SERVIDORES_PADRAO = _caminho_env(
    "JMS_SERVIDORES", Path.home() / "Desktop" / "JOAO" / "Servidores"
)

AUTOFICHA_DIR = _caminho_env(
    "JMS_AUTOFICHA_DIR", Path.home() / "Documents" / "AutoFicha"
)
AUTOFICHA_SCRIPT = AUTOFICHA_DIR / "src" / "main.py"

# Nome próprio (diferente do "temporario.pdf" do AutoFicha) para não haver
# colisão em Downloads caso os dois projetos rodem na mesma máquina.
NOME_TEMPORARIO = "temporario_jms.pdf"
CAMINHO_TEMPORARIO = PASTA_DOWNLOADS / NOME_TEMPORARIO

IMPRESSORA_PDF = "microsoft print to pdf"

# ======================================================================
# TÍTULOS DE JANELAS DO SISTEMA LEGADO
# ======================================================================
TITULO_INICIAL = "Form1"
TITULO_INTERNO = "Form1"
TITULO_JMS = "Folha - PREFEITURA MUNICIPAL DE TRES CORACOES - Abril de 2009"
TITULO_REPORT = (
    "Report Designer - fichafin.frx - "
    "Folha - PREFEITURA MUNICIPAL DE TRES CORACOES - Abril de 2009"
)
TITULO_REPORT_VARIACAO = (
    "Report Designer - fichafin.frx - Page 1 - "
    "Folha - PREFEITURA MUNICIPAL DE TRES CORACOES - Abril de 2009"
)
TRECHO_REPORT = "Report Designer"
TITULO_EXTERNO_IMPRIMIR = "Imprimir"
TITULO_EXTERNO_PRINT = "Print"
TITULO_IMPRESSAO = "Salvar Saída de Impressão como"

# ======================================================================
# TEMPOS (segundos)
# ======================================================================
TEMPO_LIMITE = 60                # espera padrão por uma janela
INTERVALO_VERIFICACAO = 0.2      # polling; a enumeração de janelas é barata
TEMPO_JMS_INICIAL = 3            # espera pela janela principal antes do aviso "Form1"
TEMPO_JMS = 10                   # espera pela janela principal depois do aviso
TEMPO_JANELA_INTERNA = 15
TEMPO_REPORT = 15                # Report Designer obrigatório (modo interno)
TEMPO_REPORT_OPCIONAL = 2        # Report Designer opcional (modo externo imprimir)
TEMPO_GRACE_MODO = 2.0           # após surgir um Report Designer novo, espera Imprimir/Print
TEMPO_PDF = 15.0
PDF_INTERVALO = 0.2
PDF_LEITURAS_ESTAVEIS = 2

# Pausas do legado. Mantidas iguais às originais: só reduza com medições (--perf).
ESPERA_FORMULARIO = 1.0
ESPERA_REPOSICIONAR = 0.5
INTERVALO_TAB_ANO = 0.3
INTERVALO_TAB_MATRICULA = 0.02

# pyautogui.PAUSE (0.1 s) já era o valor implícito da versão anterior e é ele
# que separa as teclas; os time.sleep(.02) extras eram redundantes.
PYAUTOGUI_PAUSE = 0.1

MAX_TENTATIVAS = 2               # por ficha; entre tentativas a sessão é reiniciada

# True: um Report Designer NOVO (hwnd inédito) indica modo interno sem esperar
# os 60 s do timeout. Desligue se o seu fluxo externo abrir o Report antes do diálogo.
MODO_INTERNO_RAPIDO = True

# Clique do modo interno (coordenadas absolutas de tela, como no original).
REPORT_CLIQUE = (373, 325)


# ======================================================================
# CREDENCIAIS
# ======================================================================
def ler_credenciais() -> tuple[str, str]:
    """Lê usuário e senha (duas primeiras linhas). Nunca registra a senha em log."""
    if not CRED.is_file():
        raise ErroConfiguracao(f"Arquivo de credenciais não encontrado: {CRED}")
    try:
        linhas = CRED.read_text(encoding="utf-8-sig").splitlines()
    except (OSError, UnicodeDecodeError) as exc:
        raise ErroConfiguracao(f"Não foi possível ler {CRED}: {exc}") from exc
    if len(linhas) < 2 or not linhas[0].strip() or not linhas[1]:
        raise ErroConfiguracao(
            f"{CRED} deve ter o usuário na linha 1 e a senha na linha 2."
        )
    return linhas[0].strip(), linhas[1]
