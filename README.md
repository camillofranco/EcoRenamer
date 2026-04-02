# 🌿 EcoRenamer Pro v1.4.5 (Enterprise Edition)

O **EcoRenamer Pro** é um assistente de produtividade empresarial desenvolvido para automação de gestão de arquivos e conversões avançadas. Esta versão **Clean Edition** foi otimizada para oferecer o máximo desempenho, privacidade e facilidade de uso, com processamento 100% local.

![Interface EcoRenamer Pro](https://raw.githubusercontent.com/camillofranco/EcoRenamer/main/logo_ecowave.png)

## ✨ Principais Funcionalidades

### 📸 Gestão Avançada de Imagens
- **Renomeação em Massa**: Renomeia centenas de fotos em segundos com base em planilhas Excel ou sequências numéricas customizadas.
- **Drag & Drop**: Reordene as fotos manualmente antes de processar.
- **Compressão Inteligente**: Opção de otimizar o tamanho das fotos para armazenamento em nuvem sem perda de qualidade visual perceptível.
- **Processamento Paralelo**: Motor multi-core que processa arquivos simultaneamente.

### 📄 Gestão de PDFs
- **Merge Profissional**: Una múltiplos arquivos PDF em um único documento unificado.
- **Reordenação**: Organize a ordem das páginas arrastando os arquivos na lista.
- **Divisão (Split)**: Separe cada página de um PDF em arquivos individuais com um único clique.

### 🛠️ Caixa de Ferramentas (Utilitários)
- **PDF para Word (OCR)**: Converte PDFs nativos e **escaneados** (imagens) em documentos Word (.docx) 100% editáveis usando tecnologia Tesseract OCR.
- **Excel para PDF**: Gera relatórios PDF profissionais a partir de tabelas Excel.
- **PDF para Excel**: Extrai tabelas integradas de PDFs para análise de dados.
- **Fotos para PDF**: Cria álbuns PDF de alta qualidade a partir de fotos (JPG/PNG/HEIF), respeitando a rotação correta da câmera.

## 🚀 Instalação e Execução

O EcoRenamer Pro é um aplicativo portátil disponível para as principais plataformas.

### 🍏 MacOS
1. Baixe o arquivo `EcoRenamer_Mac_v1.4.5.zip` nas [Releases](https://github.com/camillofranco/EcoRenamer/releases).
2. Extraia e mova o `RenomeadorApp.app` para sua pasta de Aplicativos.
3. Se o Mac bloquear a abertura por ser de um "Desenvolvedor não Identificado":
   - Botão direito no App > Abrir > Confirmar.

### 🪟 Windows
1. Baixe o `EcoRenamer_Win_v1.4.5.zip`.
2. Extraia e execute o `RenomeadorApp.exe`.

## 📦 Como Compilar (Para Desenvolvedores)

Se você desejar compilar o executável sozinho, o projeto já inclui scripts de automação:

- **MacOS**: `bash compilar_mac.command`
- **Windows**: Executar `compilar_win.bat`

**Requisitos**: 
- Python 3.10+
- Tesseract OCR instalado no sistema (para a função de OCR local).

### 🛠️ Instalando o Tesseract (Obrigatório para PDF -> Word)
- **Mac**: `brew install tesseract tesseract-lang`
- **Win**: Baixar instalador em [UB-Mannheim Tesseract](https://github.com/UB-Mannheim/tesseract/wiki)

## 🔄 Sistema de Atualização
O aplicativo possui um verificador de versão integrado. Sempre que houver uma nova release no GitHub, o botão **"Baixar Atualizações"** irá notificá-lo e permitir o download automático.

---
**Desenvolvido por EcoWave Tech**
*Versão v1.4.5 - Edição Estável & Privada*
