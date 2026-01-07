# 📸 AutoFoto Report Pro

![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)
![Python](https://img.shields.io/badge/Python-3.x-yellow.svg)
![Status](https://img.shields.io/badge/Status-Stable-green.svg)

**AutoFoto Report** é uma ferramenta de automação desktop desenvolvida para engenheiros, técnicos e profissionais que precisam gerar relatórios fotográficos complexos no Microsoft Word de forma rápida e padronizada.

O software elimina o trabalho manual de "copiar, colar e redimensionar" fotos, permitindo organizar centenas de imagens em lotes, editá-las visualmente e gerar um documento `.docx` (e `.pdf`) formatado em segundos.

---

## ✨ Funcionalidades Principais

* **📂 Organização Inteligente por Lotes:** Detecta subpastas automaticamente e cria seções com títulos no relatório (ex: "Lote A", "Lote B").
* **🎨 Editor Visual Integrado:**
    * **Corte (Crop):** Selecione a área de interesse na foto.
    * **Rotação:** Gire imagens individualmente ou em lote.
    * **Brilho:** Ajuste a luminosidade de fotos escuras.
    * **Legendas:** Adicione legendas que aparecem formatadas no documento.
* **📄 Layouts Flexíveis:**
    * **Normal:** Uma foto por linha.
    * **Lado a Lado (Tabela):** Duas fotos por linha, ideal para comparações.
* **⚙️ Configurações Avançadas:**
    * **Carimbo de Data/Hora:** Adiciona a data original da foto (EXIF) na imagem.
    * **Exportação PDF:** Gera automaticamente uma versão PDF usando o motor do MS Word.
    * **Controle de Tamanho:** Defina altura/largura máxima em centímetros.
* **🛡️ Segurança:** Impede sobrescrita acidental de arquivos e condições de corrida.
* **💎 Interface Moderna:** Tema escuro "Superhero" (via `ttkbootstrap`) para conforto visual.

---

## 🚀 Instalação e Requisitos

### Pré-requisitos
* Python 3.10 ou superior.
* Microsoft Word instalado (para conversão PDF).

### Instalação das Dependências

Abra o terminal na pasta do projeto e execute:

```bash
pip install opencv-python numpy python-docx Pillow ttkbootstrap docx2pdf
```

## 📖 Como Usar

1.  **Execute o programa:**
    ```bash
    python launcher.py
    ```
2.  **Selecione os Arquivos:**
    * **1. Relatório Word:** Escolha seu modelo `.docx` (pode ter cabeçalho, rodapé, textos prévios).
    * **2. Pasta de Fotos:** Selecione a pasta raiz contendo as imagens ou subpastas (lotes).
    * **3. Salvar Como:** Escolha onde salvar o relatório final (o nome deve ser diferente da entrada!).
3.  **Defina o Local (Passo 4):**
    * Clique em "Local de Inserção" e escolha após qual parágrafo do seu modelo as fotos devem começar.
4.  **Configurações (Opcional):**
    * Clique em "Configurações" para ajustar tamanho, layout (lado a lado), carimbos, etc.
5.  **Iniciar:**
    * Clique em **INICIAR PROCESSAMENTO**.
    * O Editor Visual abrirá. Faça seus cortes, ajustes e adicione legendas.
    * Clique em "Finalizar Edição" para gerar o relatório.

## 🛠️ Gerando Executável (.exe)

Para distribuir o software sem precisar instalar Python em outras máquinas, utilize o **PyInstaller**:

```bash
pyinstaller --onefile --windowed --name="AutoFotoReport" --hidden-import="editortkinter" --hidden-import="PIL.ImageEnhance" --hidden-import="tkinter.scrolledtext" --hidden-import="cv2" --hidden-import="numpy" --hidden-import="docx" --hidden-import="docx2pdf" launcher.py
```

O arquivo .exe será criado na pasta dist.

## ⚖️ Licença

Este projeto está licenciado sob a **GNU General Public License v3.0 (GPLv3)**.

Isso significa que você tem a liberdade de:
* Usar o software para fins comerciais ou privados.
* Modificar o código fonte.
* Distribuir cópias.

**Contudo**, se você distribuir o software (original ou modificado), você **deve** disponibilizar o código-fonte sob a mesma licença (GPLv3). Você não pode fechar o código e torná-lo proprietário.

Veja o arquivo [LICENSE](LICENSE) para mais detalhes.

---

## 👨‍💻 Autor

Desenvolvido por **GiraGirardelli** (Pedro H.G.C Vidal).
