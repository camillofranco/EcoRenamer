from PIL import Image
import os

def create_icon_from_image(input_path, output_icon_path):
    img = Image.open(input_path)
    # Garantir que seja quadrada com fundo branco
    size = max(img.size)
    new_img = Image.new("RGBA", (size, size), (255, 255, 255, 255))
    # Centralizar logo
    offset = ((size - img.width) // 2, (size - img.height) // 2)
    new_img.paste(img, offset, img if img.mode == 'RGBA' else None)
    
    # Salvar em múltiplos tamanhos para .ico (Windows) e criar uma versão grande para Mac
    # Infelizmente PIL não cria .icns nativamente fácil, mas o PyInstaller no Mac
    # aceita um .png de alta resolução e converte se necessário, ou podemos salvar como .ico
    # que o Mac também costuma aceitar na compilação.
    icon_sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    new_img.save(output_icon_path, sizes=icon_sizes)

if __name__ == "__main__":
    create_icon_from_image("icon_source.png", "icon.ico")
