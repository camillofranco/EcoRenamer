import os
import tempfile
import unittest

from PIL import Image

from compressao_fotos import compress_photo, compressed_name


class CompressaoFotosTests(unittest.TestCase):
    def test_nomes_jpeg_sao_preservados(self):
        self.assertEqual(compressed_name("Minha Foto.JPEG"), "Minha Foto.JPEG")
        self.assertEqual(compressed_name("foto.jpg"), "foto.jpg")
        self.assertEqual(compressed_name("foto.png"), "foto.JPG")
        self.assertEqual(compressed_name("foto.heic"), "foto.JPG")

    def test_gera_jpeg_com_mesma_configuracao_sem_tocar_no_original(self):
        with tempfile.TemporaryDirectory() as folder:
            source = os.path.join(folder, "Minha Foto.jpg")
            output_dir = os.path.join(folder, "Fotos_Comprimidas")
            os.mkdir(output_dir)
            target = os.path.join(output_dir, "Minha Foto.jpg")
            Image.new("RGB", (1600, 1200), "red").save(source, "JPEG", quality=95)
            with open(source, "rb") as image_file:
                original = image_file.read()

            compress_photo(source, target)

            with Image.open(target) as compressed:
                self.assertEqual(compressed.format, "JPEG")
                self.assertEqual(compressed.size, (800, 600))
                self.assertEqual(compressed.quantization[0][0], 13)  # JPEG quality 60
            with open(source, "rb") as image_file:
                self.assertEqual(image_file.read(), original)

    def test_png_mantem_nome_base_e_vira_jpeg(self):
        with tempfile.TemporaryDirectory() as folder:
            source = os.path.join(folder, "foto.png")
            target = os.path.join(folder, compressed_name("foto.png"))
            Image.new("RGB", (1200, 900), "blue").save(source, "PNG")

            compress_photo(source, target)

            self.assertTrue(os.path.exists(source))
            with Image.open(target) as compressed:
                self.assertEqual(compressed.format, "JPEG")
                self.assertEqual(compressed.size, (800, 600))


if __name__ == "__main__":
    unittest.main()
