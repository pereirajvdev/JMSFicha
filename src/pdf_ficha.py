"""Validação, leitura e arquivamento dos PDFs das fichas.

Sem dependência do Windows nem do sistema Folha: pode ser testado em
qualquer máquina com PyMuPDF.
"""

from __future__ import annotations

import errno
import hashlib
import logging
import os
import re
import shutil
import time
from pathlib import Path

import fitz  # PyMuPDF

from config import (
    PDF_INTERVALO,
    PDF_LEITURAS_ESTAVEIS,
    TEMPO_PDF,
    ErroPDF,
)
from perf import cronometrar

log = logging.getLogger(__name__)

_RESERVADOS_WINDOWS = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}
_TAMANHO_MAX_NOME = 120
_ERROS_FITZ = (RuntimeError, ValueError, OSError)  # FileDataError herda de RuntimeError


# ======================================================================
# TEMPORÁRIO
# ======================================================================
def deletar_temporario(caminho: Path) -> None:
    """Remove o PDF temporário anterior. Levanta ErroPDF se não for possível."""
    try:
        caminho.unlink(missing_ok=True)
    except OSError as exc:
        raise ErroPDF(f"Não foi possível remover o PDF temporário {caminho}: {exc}") from exc


def validar_pdf(caminho: Path) -> int:
    """Confirma que o PDF está completo e legível. Devolve o nº de páginas.

    Existir e ter tamanho > 0 não basta: confere cabeçalho, marcador %%EOF,
    abertura pelo PyMuPDF, ausência de criptografia e leitura da 1ª página.
    """
    try:
        with caminho.open("rb") as arq:
            cabecalho = arq.read(5)
            arq.seek(0, os.SEEK_END)
            arq.seek(max(arq.tell() - 2048, 0))
            final = arq.read()
    except OSError as exc:
        raise ErroPDF(f"PDF inacessível: {exc}") from exc

    if cabecalho != b"%PDF-":
        raise ErroPDF("arquivo sem cabeçalho PDF")
    if b"%%EOF" not in final:
        raise ErroPDF("PDF incompleto (sem marcador %%EOF)")

    try:
        with fitz.open(caminho) as doc:
            if doc.is_encrypted:
                raise ErroPDF("PDF criptografado")
            paginas = doc.page_count
            if paginas < 1:
                raise ErroPDF("PDF sem páginas")
            doc.load_page(0)
            return paginas
    except _ERROS_FITZ as exc:
        raise ErroPDF(f"PDF corrompido ou ilegível: {exc}") from exc


@cronometrar("pdf:aguardar_pronto")
def aguardar_pdf_pronto(caminho: Path, timeout: float = TEMPO_PDF) -> int:
    """Espera o PDF existir, parar de crescer e passar em validar_pdf()."""
    limite = time.monotonic() + timeout
    ultimo = -1
    estavel = 0
    motivo = "arquivo não apareceu"

    while True:
        try:
            tamanho = caminho.stat().st_size
        except OSError:
            tamanho = 0
            ultimo, estavel = -1, 0
        if tamanho > 0:
            estavel = estavel + 1 if tamanho == ultimo else 0
            ultimo = tamanho
            if estavel >= PDF_LEITURAS_ESTAVEIS - 1:
                try:
                    paginas = validar_pdf(caminho)
                    log.info("PDF temporário válido: %d página(s), %s bytes", paginas, f"{tamanho:,}")
                    return paginas
                except ErroPDF as exc:
                    motivo = str(exc)
        if time.monotonic() >= limite:
            break
        time.sleep(PDF_INTERVALO)

    raise ErroPDF(f"PDF temporário não ficou pronto em {timeout:.0f}s ({motivo}): {caminho}")


# ======================================================================
# EXTRAÇÃO DE DADOS
# ======================================================================
def extrair_dados_servidor_modelo_antigo(texto: str) -> dict:
    """Extrai Matrícula, Nome, Cargo e Ano do texto da ficha (modelo antigo)."""
    dados: dict = {}

    achado = re.search(r"(?m)^Matr[ií]cula:\s*(\d+)\s+(.+?)\s*$", texto, flags=re.IGNORECASE)
    if achado:
        dados["Matrícula"] = achado.group(1)
        dados["Nome"] = achado.group(2).strip()

    achado = re.search(r"(?m)^Cargo:\s*(.+?)\s*$", texto, flags=re.IGNORECASE)
    if achado:
        dados["Cargo"] = achado.group(1).strip()

    achado = re.search(r"Relat[óo]rio\s+da\s+Ficha\s+Financeira\s+(\d{4})", texto, flags=re.IGNORECASE)
    if achado:
        dados["Ano"] = int(achado.group(1))

    return dados


@cronometrar("pdf:ler_dados")
def ler_dados_ficha(caminho: Path) -> dict:
    """Lê os dados da ficha; usa só a 1ª página e recorre ao documento todo se faltar algo."""
    try:
        with fitz.open(caminho) as doc:
            if doc.page_count < 1:
                raise ErroPDF("PDF sem páginas")
            dados = extrair_dados_servidor_modelo_antigo(doc.load_page(0).get_text())
            if "Ano" not in dados or "Nome" not in dados:
                texto = "\n".join(pagina.get_text() for pagina in doc)
                dados = {**extrair_dados_servidor_modelo_antigo(texto), **dados}
    except _ERROS_FITZ as exc:
        raise ErroPDF(f"Falha ao extrair texto do PDF: {exc}") from exc
    return dados


def sanitizar_nome(nome: str) -> str:
    """Torna o texto um nome de pasta válido no Windows (pode devolver vazio)."""
    nome = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "", nome)
    nome = re.sub(r"\s+", " ", nome).strip(" .")
    if nome.split(".")[0].upper() in _RESERVADOS_WINDOWS:
        nome = f"_{nome}"
    return nome[:_TAMANHO_MAX_NOME].strip(" .")


# ======================================================================
# ARQUIVAMENTO
# ======================================================================
def _sha256(caminho: Path) -> str:
    h = hashlib.sha256()
    with caminho.open("rb") as arq:
        for bloco in iter(lambda: arq.read(1 << 20), b""):
            h.update(bloco)
    return h.hexdigest()


def _numero(matricula: str) -> str:
    return str(matricula).strip().lstrip("0") or "0"


def _mover(origem: Path, destino: Path) -> None:
    """Move com troca atômica; entre volumes, copia, confere o tamanho e só então troca."""
    try:
        os.replace(origem, destino)
        return
    except OSError as exc:
        if exc.errno != errno.EXDEV and getattr(exc, "winerror", None) != 17:
            raise
    tmp = destino.with_name(destino.name + ".tmp")
    try:
        shutil.copy2(origem, tmp)
        if tmp.stat().st_size != origem.stat().st_size:
            raise ErroPDF("cópia incompleta ao mover o PDF entre volumes")
        os.replace(tmp, destino)
    finally:
        tmp.unlink(missing_ok=True)
    origem.unlink(missing_ok=True)


class Arquivador:
    """Move o PDF temporário para <Servidores>/<Nome>/FICHA FINANCEIRA <ano>.pdf
    (ou direto em ``pasta_fixa``), validando identidade e evitando sobrescritas."""

    def __init__(self, pasta_servidores: Path, pasta_fixa: Path | None = None) -> None:
        self.pasta_servidores = pasta_servidores
        self.pasta_fixa = pasta_fixa
        self._salvos: set[str] = set()  # destinos gravados nesta execução

    @cronometrar("pdf:arquivar")
    def arquivar(self, temporario: Path, matricula: str, ano: int) -> Path:
        if not temporario.is_file():
            raise ErroPDF(f"PDF temporário não encontrado: {temporario}")

        dados = ler_dados_ficha(temporario)
        log.info("Dados extraídos: %s", dados)

        ano_pdf = dados.get("Ano")
        if ano_pdf is None:
            raise ErroPDF("ano da ficha financeira não encontrado no PDF")
        if ano_pdf != ano:
            raise ErroPDF(f"PDF é do ano {ano_pdf}, mas o esperado era {ano}")
        mat_pdf = dados.get("Matrícula")
        if mat_pdf is not None and _numero(mat_pdf) != _numero(matricula):
            raise ErroPDF(f"PDF é da matrícula {mat_pdf}, mas o esperado era {matricula}")

        pasta = self._pasta_destino(dados)
        destino = pasta / f"FICHA FINANCEIRA {ano}.pdf"
        if self._chave(destino) in self._salvos:
            # outra matrícula do mesmo servidor, mesmo ano, nesta execução
            destino = pasta / f"FICHA FINANCEIRA {ano} - matricula {_numero(matricula)}.pdf"

        self._gravar(temporario, destino)
        self._salvos.add(self._chave(destino))
        log.info("Ficha arquivada: %s", destino)
        return destino

    @staticmethod
    def _chave(caminho: Path) -> str:
        return os.path.normcase(str(caminho))

    def _pasta_destino(self, dados: dict) -> Path:
        if self.pasta_fixa is not None:
            return self.pasta_fixa
        nome = sanitizar_nome(dados.get("Nome", ""))
        if not nome:
            raise ErroPDF("nome do servidor não encontrado (ou inválido) no PDF")
        pasta = self.pasta_servidores / nome
        try:
            pasta.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            raise ErroPDF(f"Não foi possível criar a pasta {pasta}: {exc}") from exc
        return pasta

    def _gravar(self, temporario: Path, destino: Path) -> None:
        try:
            if destino.exists():
                if _sha256(destino) == _sha256(temporario):
                    temporario.unlink(missing_ok=True)  # idêntico: nada a trocar
                    log.info("Ficha já existia idêntica: %s", destino)
                    return
                backup = destino.with_name(destino.name + ".bak")
                os.replace(destino, backup)
                log.warning("Ficha existente substituída (cópia anterior em %s)", backup.name)
            _mover(temporario, destino)
        except OSError as exc:
            raise ErroPDF(f"Falha ao mover o PDF para {destino}: {exc}") from exc
