import io
import unittest
from unittest.mock import Mock, patch

import renomeador


class AtualizacoesTests(unittest.TestCase):
    def test_versoes_sao_comparadas_numericamente(self):
        self.assertGreater(renomeador.version_parts("1.10.0"), renomeador.version_parts("1.9.2"))

    def test_manifesto_invalido_e_rejeitado(self):
        with patch.object(renomeador.urllib.request, "urlopen", return_value=io.BytesIO(b"{}")):
            with self.assertRaisesRegex(ValueError, "versão inválida"):
                renomeador.fetch_update_info()

    def test_falha_na_consulta_mostra_erro_e_reativa_botao(self):
        app = object.__new__(renomeador.ToolApp)
        app.btn_update = Mock()
        with patch.object(renomeador.messagebox, "askyesno", return_value=False) as dialog:
            app._show_update_result(None, "Falha de conexão")

        self.assertFalse(app.checking_updates)
        app.btn_update.configure.assert_called_once_with(text="♻️ Atualizações", state="normal")
        self.assertIn("Falha de conexão", dialog.call_args.args[1])


if __name__ == "__main__":
    unittest.main()
