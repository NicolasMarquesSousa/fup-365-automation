import importlib.util
import json
import sys
import tempfile
import types
import unittest
from pathlib import Path


class _Attachments:
    def __init__(self):
        self.paths = []

    def Add(self, path):
        self.paths.append(path)


class _Message:
    def __init__(self):
        self.HTMLBody = "<p>Assinatura</p>"
        self.Attachments = _Attachments()
        self.saved = False

    def Display(self, _):
        pass

    def Save(self):
        self.saved = True

    def Close(self, _):
        pass


class _Outlook:
    def __init__(self):
        self.messages = []

    def CreateItem(self, _):
        message = _Message()
        self.messages.append(message)
        return message


OUTLOOK = _Outlook()
client = types.ModuleType("win32com.client")
client.Dispatch = lambda _: OUTLOOK
win32com = types.ModuleType("win32com")
win32com.client = client
sys.modules.setdefault("win32com", win32com)
sys.modules.setdefault("win32com.client", client)

MODULE_PATH = Path(__file__).parents[1] / "src" / "criar_rascunhos_outlook.py"
SPEC = importlib.util.spec_from_file_location("criar_rascunhos_outlook", MODULE_PATH)
MODULE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(MODULE)


class CriarRascunhosTest(unittest.TestCase):
    def setUp(self):
        OUTLOOK.messages.clear()

    def test_tabela_html_escapa_conteudo(self):
        row = ["<script>", 1, "P-1", "01/2026", "Fornecedor", 2, 10, 20]
        result = MODULE.tabela_html([row])
        self.assertIn("&lt;script&gt;", result)
        self.assertNotIn("<script>", result)
        self.assertIn("border-collapse:collapse", result)

    def test_corpo_preserva_assinatura(self):
        item = {"pedidos": [["Item", 1, "P-1", "01/2026", "F", 1, 1, 1]]}
        result = MODULE.corpo_html(item, "<footer>Assinatura</footer>")
        self.assertIn("Assinatura", result)
        self.assertIn("revisado antes do envio", result)

    def test_cria_rascunho_sem_enviar(self):
        with tempfile.TemporaryDirectory(dir=Path.cwd()) as temp:
            folder = Path(temp)
            attachment = folder / "pedidos.csv"
            attachment.write_text("pedido;status\n1;pendente", encoding="utf-8")
            manifest = folder / "manifesto.json"
            manifest.write_text(json.dumps([{
                "fornecedor": "Fornecedor Teste",
                "para": "contato@example.test",
                "assunto": "Follow-up",
                "anexo": attachment.name,
                "pedidos": [["Item", 1, "P-1", "01/2026", "F", 1, 1, 1]],
            }]), encoding="utf-8")

            self.assertEqual(MODULE.criar_rascunhos(manifest, "cc@example.test"), 1)
            message = OUTLOOK.messages[0]
            self.assertTrue(message.saved)
            self.assertEqual(message.To, "contato@example.test")
            self.assertEqual(message.CC, "cc@example.test")
            self.assertEqual(message.Attachments.paths, [str(attachment.resolve())])

    def test_falha_quando_anexo_nao_existe(self):
        with tempfile.TemporaryDirectory(dir=Path.cwd()) as temp:
            folder = Path(temp)
            manifest = folder / "manifesto.json"
            manifest.write_text(json.dumps([{
                "fornecedor": "Fornecedor Teste", "para": "a@example.test",
                "assunto": "Follow-up", "anexo": "ausente.csv", "pedidos": [],
            }]), encoding="utf-8")
            with self.assertRaises(FileNotFoundError):
                MODULE.criar_rascunhos(manifest)


if __name__ == "__main__":
    unittest.main()
