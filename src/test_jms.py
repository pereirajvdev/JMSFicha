import json, os, sys, tempfile, threading, time, types, unittest
from pathlib import Path
from unittest import mock

SRC = Path(__file__).resolve().parents[1] / "src"
sys.path.insert(0, str(SRC))

# ---- fitz falso: lê "%PDF-...", texto entre <T>..</T> por página separadas por <P> ----
class _Pag:
    def __init__(s, t): s.t = t
    def get_text(s): return s.t
class _Doc:
    def __init__(s, caminho):
        raw = Path(caminho).read_bytes()
        if not raw.startswith(b"%PDF-"): raise RuntimeError("not a pdf")
        corpo = raw.decode("latin-1").split("<T>", 1)[-1].split("</T>")[0]
        s.pags = [_Pag(p) for p in corpo.split("<P>")]
        s.is_encrypted = False
    page_count = property(lambda s: len(s.pags))
    def load_page(s, i): return s.pags[i]
    def __iter__(s): return iter(s.pags)
    def __enter__(s): return s
    def __exit__(s, *a): pass
fitz = types.ModuleType("fitz"); fitz.open = _Doc
sys.modules["fitz"] = fitz

import config, perf, pdf_ficha, contrato
from config import ErroPDF

def pdf(path, paginas, completo=True):
    t = "<P>".join(paginas)
    Path(path).write_bytes(f"%PDF-1.4\n<T>{t}</T>\n".encode("latin-1") + (b"%%EOF\n" if completo else b""))

FICHA = ("Relatório da Ficha Financeira {ano}\nMatricula: {mat} {nome}\nCargo: Professor\n")

class TestPDF(unittest.TestCase):
    def setUp(self): self.d = Path(tempfile.mkdtemp())

    def test_sanitizar(self):
        S = pdf_ficha.sanitizar_nome
        self.assertEqual(S('  JOSE  da  Silva: "X"/\\? '), "JOSE da Silva X")
        self.assertEqual(S("CON"), "_CON"); self.assertEqual(S("..."), ""); self.assertEqual(S("a."), "a")
        self.assertLessEqual(len(S("A" * 500)), 120)

    def test_extrair(self):
        d = pdf_ficha.extrair_dados_servidor_modelo_antigo(FICHA.format(ano=2005, mat="0123", nome="MARIA SILVA"))
        self.assertEqual(d, {"Matrícula": "0123", "Nome": "MARIA SILVA", "Cargo": "Professor", "Ano": 2005})
        self.assertEqual(pdf_ficha.extrair_dados_servidor_modelo_antigo("lixo"), {})

    def test_validar(self):
        a = self.d / "a.pdf"; pdf(a, ["x"]); self.assertEqual(pdf_ficha.validar_pdf(a), 1)
        pdf(a, ["x"], completo=False)
        with self.assertRaises(ErroPDF): pdf_ficha.validar_pdf(a)
        a.write_bytes(b"nao eh pdf")
        with self.assertRaises(ErroPDF): pdf_ficha.validar_pdf(a)

    def test_aguardar_incompleto_e_completo(self):
        a = self.d / "t.pdf"
        with self.assertRaises(ErroPDF): pdf_ficha.aguardar_pdf_pronto(a, timeout=0.6)  # nunca aparece
        pdf(a, ["x"], completo=False)
        with self.assertRaises(ErroPDF): pdf_ficha.aguardar_pdf_pronto(a, timeout=0.8)   # truncado
        def escreve():
            time.sleep(0.5); pdf(a, ["x", "y"])
        threading.Thread(target=escreve).start()
        self.assertEqual(pdf_ficha.aguardar_pdf_pronto(a, timeout=5), 2)

    def arq(self, fixa=None):
        return pdf_ficha.Arquivador(self.d / "Serv", fixa)

    def test_arquivar_fluxo(self):
        A = self.arq(); t = self.d / "tmp.pdf"
        pdf(t, [FICHA.format(ano=2005, mat="12", nome="ANA / LIMA")])
        dest = A.arquivar(t, "12", 2005)
        self.assertEqual(dest, self.d / "Serv" / "ANA LIMA" / "FICHA FINANCEIRA 2005.pdf")
        self.assertTrue(dest.exists()); self.assertFalse(t.exists())
        # mesma ficha de novo (nova execução): idêntica -> sem .bak
        A2 = self.arq(); pdf(t, [FICHA.format(ano=2005, mat="12", nome="ANA / LIMA")])
        self.assertEqual(A2.arquivar(t, "12", 2005), dest); self.assertFalse(dest.with_name(dest.name + ".bak").exists())
        # conteúdo diferente de execução anterior: substitui e guarda .bak
        A3 = self.arq(); pdf(t, [FICHA.format(ano=2005, mat="12", nome="ANA / LIMA") + "extra"])
        A3.arquivar(t, "12", 2005); self.assertTrue(dest.with_name(dest.name + ".bak").exists())
        # outra matrícula, mesmo servidor/ano, NA MESMA execução: não sobrescreve
        pdf(t, [FICHA.format(ano=2005, mat="99", nome="ANA / LIMA")])
        d2 = A3.arquivar(t, "99", 2005)
        self.assertEqual(d2.name, "FICHA FINANCEIRA 2005 - matricula 99.pdf"); self.assertTrue(dest.exists())

    def test_arquivar_rejeita(self):
        A = self.arq(); t = self.d / "tmp.pdf"
        pdf(t, [FICHA.format(ano=2004, mat="12", nome="ANA")])
        with self.assertRaises(ErroPDF): A.arquivar(t, "12", 2005)           # ano errado
        pdf(t, [FICHA.format(ano=2005, mat="13", nome="ANA")])
        with self.assertRaises(ErroPDF): A.arquivar(t, "12", 2005)           # matrícula errada
        pdf(t, ["Relatório da Ficha Financeira 2005\n"])
        with self.assertRaises(ErroPDF): A.arquivar(t, "12", 2005)           # sem nome
        with self.assertRaises(ErroPDF): A.arquivar(self.d / "nao.pdf", "12", 2005)
        self.assertTrue(t.exists())                                          # inválido não é movido

    def test_pasta_fixa(self):
        fixa = self.d / "fixa"; fixa.mkdir(); A = self.arq(fixa); t = self.d / "tmp.pdf"
        pdf(t, [FICHA.format(ano=2001, mat="5", nome="BETO")])
        self.assertEqual(A.arquivar(t, "5", 2001), fixa / "FICHA FINANCEIRA 2001.pdf")

class TestPerf(unittest.TestCase):
    def test_proprio_vs_total(self):
        perf.PERFIL.limpar(); perf.ativar(True)
        with perf.medir("pai"):
            time.sleep(0.05)
            with perf.medir("filho"): time.sleep(0.1)
        d = perf.PERFIL._dados
        self.assertAlmostEqual(d["pai"][1], 0.15, delta=0.03); self.assertAlmostEqual(d["pai"][2], 0.05, delta=0.03)
        self.assertIn("filho", perf.PERFIL.relatorio())
        perf.ativar(False); perf.PERFIL.limpar()
        @perf.cronometrar()
        def f(): return 7
        self.assertEqual(f(), 7); self.assertEqual(perf.PERFIL._dados, {})

class TestContrato(unittest.TestCase):
    def test_retorno_atomico(self):
        d = Path(tempfile.mkdtemp()); r = contrato.Retorno(d / "sub" / "r.json")
        r.ficha_salva(d); r.ficha_salva(d); r.falha("x"); r.gravar(False, "e")
        j = json.loads((d / "sub" / "r.json").read_text("utf-8"))
        self.assertEqual((j["total_pastas"], j["total_fichas_salvas"], j["sucesso"]), (1, 2, False))
        self.assertFalse((d / "sub" / "r.json.tmp").exists())

# ---- main com automação/janelas falsas ----
def instala_falsos(script):
    """script: lista de resultados por chamada de gerar_e_salvar_pdf ('ok' ou exceção)."""
    auto = types.ModuleType("automacao"); log = []
    auto.configurar_pyautogui = lambda: log.append("cfg")
    auto.garantir_impressora_pdf = lambda: log.append("imp")
    auto.iniciar_sessao = lambda: log.append("sessao")
    auto.encerrar_folha = lambda: log.append("encerrar")
    auto.reposicionar_formulario = lambda mudou_matricula: log.append(("repos", mudou_matricula))
    def gerar(mat, ano):
        log.append(("gerar", mat, ano)); r = script.pop(0)
        if isinstance(r, Exception): raise r
        pdf(config.CAMINHO_TEMPORARIO, [FICHA.format(ano=ano, mat=mat, nome=f"SERV {mat}")])
    auto.gerar_e_salvar_pdf = gerar
    jan = types.ModuleType("janelas"); jan.EXC_AUTOMACAO = (OSError,)
    sys.modules.update(automacao=auto, janelas=jan)
    return log

class TestMain(unittest.TestCase):
    def setUp(self):
        self.d = Path(tempfile.mkdtemp())
        (self.d / "dl").mkdir(); (self.d / "exe").write_text("x"); (self.d / "cred").write_text("u\np\n")
        for k, v in dict(CAMINHO_TEMPORARIO=self.d / "dl" / "t.pdf", PASTA_DOWNLOADS=self.d / "dl",
                         EXE=self.d / "exe", CRED=self.d / "cred").items():
            setattr(config, k, v)
        import main as m; self.m = m
        for k in ("CAMINHO_TEMPORARIO", "PASTA_DOWNLOADS", "EXE", "CRED"): setattr(m, k, getattr(config, k))
        self.serv = self.d / "Serv"; self.serv.mkdir(); self.ret = self.d / "ret.json"

    def run_main(self, *extra):
        return self.m.main(["--servidores", str(self.serv), "--retorno", str(self.ret), *extra])

    def regs(self, *itens): return json.dumps([{"matricula": m, "inicio": a, "fim": b} for m, a, b in itens])

    def test_sucesso_e_reposicionamento(self):
        log = instala_falsos(["ok"] * 3)
        rc = self.run_main("--registros", self.regs(("10", 2000, 2001), ("20", 2005, 2005)))
        j = json.loads(self.ret.read_text("utf-8"))
        self.assertEqual((rc, j["sucesso"], j["total_fichas_salvas"], j["total_pastas"]), (0, True, 3, 2))
        self.assertEqual([l for l in log if isinstance(l, tuple) and l[0] == "repos"], [("repos", False), ("repos", True)])
        self.assertEqual(log[-1], "encerrar"); self.assertTrue((self.serv / "SERV 10" / "FICHA FINANCEIRA 2001.pdf").exists())
        self.assertFalse(config.CAMINHO_TEMPORARIO.exists())

    def test_falha_nao_derruba_os_demais_e_tem_retry(self):
        # ano 2000: falha nas 2 tentativas; 2001: falha 1x e passa; 2002 ok
        log = instala_falsos([OSError("boom"), OSError("boom"), OSError("x"), "ok", "ok"])
        rc = self.run_main("--registros", self.regs(("10", 2000, 2002)))
        j = json.loads(self.ret.read_text("utf-8"))
        self.assertEqual((rc, j["sucesso"], j["total_fichas_salvas"], len(j["falhas"])), (1, False, 2, 1))
        self.assertIn("ano 2000", j["falhas"][0]["mensagem"])
        self.assertEqual(log.count("sessao"), 1 + 3)   # 1 inicial + 1 reinício por falha

    def test_falha_de_login_grava_retorno(self):
        instala_falsos([])
        sys.modules["automacao"].iniciar_sessao = lambda: (_ for _ in ()).throw(config.ErroAutomacao("sem login"))
        rc = self.run_main("--registros", self.regs(("10", 2000, 2000)))
        j = json.loads(self.ret.read_text("utf-8")); self.assertEqual((rc, j["erro"]), (1, "sem login"))

    def test_pre_requisito_ausente(self):
        instala_falsos([]); self.m.EXE = self.d / "nao_existe.exe"
        rc = self.run_main("--registros", self.regs(("10", 2000, 2000)))
        self.assertEqual(rc, 1); self.assertIn("Executável", json.loads(self.ret.read_text("utf-8"))["erro"])

    def test_cli_invalida(self):
        for argv in (["--mat", "12", "2005-2000"], ["--mat", "abc", "2000-2001"], ["--mat", "1", "2000"],
                     ["--registros", "{}"], ["--registros", "[]"], [], ["--servidores", "/nao/existe", "--mat", "1", "2000-2000"],
                     ["--registros", self.regs(("1", 1, 2)), "--mat", "1", "2000-2000"]):
            with self.assertRaises(SystemExit) as cm, mock.patch("sys.stderr"):
                self.m.main(argv)
            self.assertEqual(cm.exception.code, 2, argv)

    def test_planilha(self):
        from openpyxl import Workbook
        from datetime import datetime
        wb = Workbook(); ws = wb.active; ws.append(["mat", "ini", "fim"])
        ws.append([1234.0, datetime(2000, 5, 1), datetime(2002, 1, 1)])
        ws.append(["1234", "2001-01-01", "05/03/2004"]); ws.append([None, None, None])
        ws.append(["abc", datetime(2000, 1, 1), datetime(2000, 1, 1)]); ws.append([77, datetime(2005, 1, 1), datetime(2003, 1, 1)])
        ws.append([88, 2007, 2008]); f = self.d / "p.xlsx"; wb.save(f)
        dados, avisos = self.m.ler_planilha(f)
        self.assertEqual(dados, [("1234", [2000, 2001, 2002, 2003, 2004]), ("88", [2007, 2008])]); self.assertEqual(len(avisos), 2)
        (self.d / "ruim.xlsx").write_text("nao zip")
        with self.assertRaises(config.ErroConfiguracao): self.m.ler_planilha(self.d / "ruim.xlsx")

    def test_credenciais(self):
        self.assertEqual(config.ler_credenciais(), ("u", "p"))
        (self.d / "cred").write_text("so_usuario\n")
        with self.assertRaises(config.ErroConfiguracao): config.ler_credenciais()

if __name__ == "__main__":
    unittest.main(verbosity=2)
