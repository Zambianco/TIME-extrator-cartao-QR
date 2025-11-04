"""
TIME-extrator-cartao-QR.py - Script para extração automatizada de times de cartões manuais com QR codes

Este script processa arquivos PDF contendo cartões manuais com QR codes, realiza:
1. Detecção de QR codes nas páginas
2. Alinhamento das páginas baseado na posição dos QR codes
3. Recorte das áreas contendo os times
4. Organização dos arquivos resultantes em pastas

Funcionalidades principais:
- Processamento em lote de múltiplos PDFs
- Conversão de nomes em base36 para decimal
- Organização automática em subpastas
- Tratamento de páginas sem QR codes

Autor: Rodrigo Zambianco
Data: [DATA DE CRIAÇÃO]
"""

import os
import re
import fitz
from PIL import Image
from pyzbar.pyzbar import decode
import math
import uuid
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
from tkinter import ttk

# Variável global para contar páginas sem QR codes
contador_paginas_sem_qr = 0


def sanitize_filename(text):
    """
    Remove caracteres inválidos para nomes de arquivos.
    
    Args:
        text (str): Texto original a ser sanitizado
        
    Returns:
        str: Texto sem caracteres inválidos para nomes de arquivos
    """
    return re.sub(r'[\\/*?:"<>|]', "_", text)


def pdf_page_to_image(page, dpi=300):
    """
    Renderiza uma página do PDF em alta resolução como imagem PIL.
    
    Args:
        page (fitz.Page): Página do PDF a ser renderizada
        dpi (int): Resolução desejada (default 300)
        
    Returns:
        PIL.Image: Imagem renderizada da página
    """
    zoom = dpi / 96  # 72 is the default PDF DPI
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return img


def read_qrcodes_from_image(img, max_qrcodes=2):
    """
    Detecta QR codes em uma imagem e retorna seus conteúdos e coordenadas.
    
    Args:
        img (PIL.Image): Imagem a ser analisada
        max_qrcodes (int): Número máximo de QR codes a serem detectados (default 2)
        
    Returns:
        list: Lista de dicionários com informações dos QR codes detectados
              Cada dicionário contém:
              - 'conteudo': texto decodificado do QR code
              - 'cantos': coordenadas dos cantos do QR code
    """
    detected = decode(img)
    qrcodes = []
    for qr in detected[:max_qrcodes]:
        data = qr.data.decode('utf-8')
        corners = [(p.x, p.y) for p in qr.polygon]
        qrcodes.append({
            "conteudo": data,
            "cantos": corners
        })
    return qrcodes


def calcular_angulo_entre_pontos(p1, p2):
    """
    Calcula o ângulo entre dois pontos para alinhamento vertical.
    
    Args:
        p1 (tuple): Coordenadas (x,y) do primeiro ponto
        p2 (tuple): Coordenadas (x,y) do segundo ponto
        
    Returns:
        float: Ângulo de rotação necessário para alinhamento vertical
    """
    if p1[1] > p2[1]:
        p1, p2 = p2, p1

    dx = p2[0] - p1[0]
    dy = p2[1] - p1[1]
    ang_calc = math.degrees(math.atan2(dy, dx))
    ang_rot = -(90 - ang_calc)

    # Normaliza para [-180, 180)
    while ang_rot >= 180:
        ang_rot -= 360
    while ang_rot < -180:
        ang_rot += 360

    return ang_rot


def calcular_centro_qr(cantos):
    """
    Calcula o centro geométrico de um QR code a partir de seus cantos.
    
    Args:
        cantos (list): Lista de coordenadas (x,y) dos cantos do QR code
        
    Returns:
        tuple: Coordenadas (x,y) do centro do QR code
    """
    xs = [p[0] for p in cantos]
    ys = [p[1] for p in cantos]
    return (sum(xs) / len(xs), sum(ys) / len(ys))


def mm_to_pixels(mm, dpi=300):
    """
    Converte milímetros para pixels baseado na resolução DPI.
    
    Args:
        mm (float): Medida em milímetros
        dpi (int): Resolução em dots per inch (default 300)
        
    Returns:
        int: Valor convertido em pixels
    """
    inches = mm / 25.4
    return int(inches * dpi)


def processar_qr_code(img_rotacionada, qr_info, dpi=300):
    """
    Processa um QR code individualmente e retorna o recorte da área do time.
    
    Args:
        img_rotacionada (PIL.Image): Imagem já rotacionada
        qr_info (dict): Informações do QR code (conteúdo e cantos)
        dpi (int): Resolução em DPI (default 300)
        
    Returns:
        PIL.Image: Imagem recortada contendo o time ou None em caso de erro
    """
    # Tamanhos em mm
    qr_size_mm = 17
    crop_width_mm = 70
    crop_height_mm = 75
    crop_center_x_mm = -20  # Relativo ao centro do QR
    crop_center_y_mm = 28  # Relativo ao centro do QR

    # Converter para pixels
    qr_size_px = mm_to_pixels(qr_size_mm, dpi)
    crop_width_px = mm_to_pixels(crop_width_mm, dpi)
    crop_height_px = mm_to_pixels(crop_height_mm, dpi)
    crop_center_x_px = mm_to_pixels(crop_center_x_mm, dpi)
    crop_center_y_px = mm_to_pixels(crop_center_y_mm, dpi)

    # Calcular centro do QR code
    qr_center_x, qr_center_y = calcular_centro_qr(qr_info['cantos'])

    # Calcular posição do recorte
    crop_x = qr_center_x + crop_center_x_px - (crop_width_px // 2)
    crop_y = qr_center_y + crop_center_y_px - (crop_height_px // 2)

    # Fazer o recorte
    box = (
        int(crop_x),
        int(crop_y),
        int(crop_x + crop_width_px),
        int(crop_y + crop_height_px))
    
    try:
        recorte = img_rotacionada.crop(box)
        return recorte
    except Exception as e:
        print(f"Erro ao recortar imagem: {e}")
        return None


def processar_pagina_e_alinhar(pdf_path, pagina, pasta_saida, dpi=300):
    """
    Processa uma página individual do PDF, detecta QR codes e realiza os recortes.
    
    Args:
        pdf_path (str): Caminho do arquivo PDF
        pagina (int): Índice da página a ser processada
        pasta_saida (str): Pasta de destino para os recortes
        dpi (int): Resolução em DPI (default 300)
    """
    global contador_paginas_sem_qr
    
    doc = fitz.open(pdf_path)
    page = doc.load_page(pagina)
    img = pdf_page_to_image(page, dpi)
    qrcodes = read_qrcodes_from_image(img, max_qrcodes=2)
    
    if len(qrcodes) < 1:
        print("Nenhum QR code detectado na página.")
        nome_arquivo = f"sem_qr_{uuid.uuid4()}"
        criar_pasta_e_salvar(pasta_saida, nome_arquivo, img)
        return
    
    # Se houver 2 QR codes, usamos para alinhar a página
    if len(qrcodes) >= 2:
        p1 = calcular_centro_qr(qrcodes[0]['cantos'])
        p2 = calcular_centro_qr(qrcodes[1]['cantos'])
        angulo = calcular_angulo_entre_pontos(p1, p2)
    else:
        # Se só tiver 1 QR code, salva com nome sequencial
        nome_arquivo = f"um_qr_{uuid.uuid4()}"
        criar_pasta_e_salvar(pasta_saida, nome_arquivo, img)
        return
    
    img_rotacionada = extrair_e_rotacionar_pagina_pdf(pdf_path, pagina, angulo, dpi)
    
    # Processar cada QR code individualmente
    for idx, qr in enumerate(qrcodes):
        recorte = processar_qr_code(img_rotacionada, qr, dpi)
        if recorte:
            nome_qr = sanitize_filename(qr['conteudo'])
            novo_nome = converter_nome_qr(nome_qr)
            criar_pasta_e_salvar(pasta_saida, novo_nome, recorte)


def extrair_e_rotacionar_pagina_pdf(pdf_path, pagina, angulo, dpi=300):
    """
    Extrai uma página do PDF e aplica rotação conforme ângulo calculado.
    
    Args:
        pdf_path (str): Caminho do arquivo PDF
        pagina (int): Índice da página
        angulo (float): Ângulo de rotação a ser aplicado
        dpi (int): Resolução em DPI (default 300)
        
    Returns:
        PIL.Image: Imagem da página extraída e rotacionada
    """
    doc = fitz.open(pdf_path)
    page = doc.load_page(pagina)
    img = pdf_page_to_image(page, dpi)
    img_rotacionada = img.rotate(angulo, expand=True)
    return img_rotacionada


def base36_to_base10_padded(s, length=8):
    """
    Converte string base36 para decimal com padding de zeros à esquerda.
    
    Args:
        s (str): String em base36 a ser convertida
        length (int): Comprimento desejado para o resultado (default 8)
        
    Returns:
        str: Número em base10 com zeros à esquerda
    """
    num = int(s, 36)
    return str(num).zfill(length)


def base36_to_base10(s):
    """
    Converte string base36 para decimal sem padding.
    
    Args:
        s (str): String em base36 a ser convertida
        
    Returns:
        str: Número em base10
    """
    return str(int(s, 36))


def converter_nome_qr(nome_qr):
    """
    Converte o formato do nome do QR code (XXXX-YYYY) para decimal (XXXXXXXX_YYYYY).
    
    Args:
        nome_qr (str): String no formato 'XXXX-YYYY'
        
    Returns:
        str: String convertida para 'XXXXXXXX_YYYYY' em decimal
             Retorna o nome original se o formato não for reconhecido
    """
    partes = nome_qr.split('-')
    if len(partes) != 2:
        return nome_qr  # Se formato não for como esperado, retorna igual
    parte1, parte2 = partes
    parte1_decimal = base36_to_base10_padded(parte1)
    parte2_decimal = base36_to_base10(parte2)
    return f"{parte1_decimal}_{parte2_decimal}"


def processar_pdf_completo(pdf_path, pasta_saida, dpi=300):
    """
    Processa todas as páginas de um arquivo PDF.
    
    Args:
        pdf_path (str): Caminho do arquivo PDF
        pasta_saida (str): Pasta de destino para os recortes
        dpi (int): Resolução em DPI (default 300)
    """
    global contador_paginas_sem_qr
    contador_paginas_sem_qr = 0  # Reseta o contador para cada novo PDF
    
    os.makedirs(pasta_saida, exist_ok=True)
    doc = fitz.open(pdf_path)
    n_paginas = len(doc)
    for i in range(n_paginas):
        print(f"Processando página {i+1}/{n_paginas}")
        processar_pagina_e_alinhar(pdf_path, i, pasta_saida, dpi)


def processar_todos_pdfs(pasta_entrada, pasta_saida, dpi=300):
    """
    Processa todos os arquivos PDF encontrados em uma pasta.
    
    Args:
        pasta_entrada (str): Pasta contendo os PDFs a serem processados
        pasta_saida (str): Pasta de destino para os recortes
        dpi (int): Resolução em DPI (default 300)
    """
    os.makedirs(pasta_saida, exist_ok=True)
    
    for arquivo in os.listdir(pasta_entrada):
        if arquivo.lower().endswith('.pdf'):
            pdf_path = os.path.join(pasta_entrada, arquivo)
            print(f"\nProcessando arquivo: {arquivo}")
            try:
                processar_pdf_completo(pdf_path, pasta_saida, dpi)
            except Exception as e:
                print(f"Erro ao processar {arquivo}: {str(e)}")
    
    print("\nProcessamento de todos os PDFs concluído!")


def criar_pasta_e_salvar(pasta_saida, nome_arquivo, imagem, quality=95):
    """
    Cria subpastas organizadas e salva a imagem com qualidade especificada.
    
    Args:
        pasta_saida (str): Pasta base de destino
        nome_arquivo (str): Nome base do arquivo (sem extensão)
        imagem (PIL.Image): Imagem a ser salva
        quality (int): Qualidade JPEG (default 95)
    """
    # Extrair o prefixo do nome do arquivo (antes do "_")
    prefixo = nome_arquivo.split('_')[0]
    # Criar o caminho da subpasta
    subpasta = os.path.join(pasta_saida, prefixo)
    os.makedirs(subpasta, exist_ok=True)

    # Montar caminho e evitar overwrite
    base_path = os.path.join(subpasta, f"{nome_arquivo}.jpg")
    arquivo_completo = base_path
    if os.path.exists(arquivo_completo):
        # adiciona sufixo incremental se já existir
        i = 1
        while True:
            tentativa = os.path.join(subpasta, f"{nome_arquivo}_{i}.jpg")
            if not os.path.exists(tentativa):
                arquivo_completo = tentativa
                break
            i += 1

    imagem.save(arquivo_completo, "JPEG", quality=quality)
    print(f"Salvando em: {arquivo_completo}")


def processar_pdf_completo_gui(pdf_path, pasta_saida, dpi, barra, label_status, idx_atual, total_arquivos):
    import fitz
    doc = fitz.open(pdf_path)
    n_paginas = len(doc)

    for i in range(n_paginas):
        label_status.config(
            text=f"Arquivo {idx_atual}/{total_arquivos} | Página {i+1}/{n_paginas}: {os.path.basename(pdf_path)}"
        )
        label_status.update_idletasks()

        try:
            processar_pagina_e_alinhar(pdf_path, i, pasta_saida, dpi)
        except Exception as e:
            print(f"Erro ao processar página {i+1} de {pdf_path}: {e}")

        # Atualiza a barra a cada página
        barra["value"] += 1
        barra.update_idletasks()

    doc.close()


def iniciar_processamento(arquivos_pdf, pasta_saida, janela, barra, label_status):
    total_paginas = 0
    import fitz

    # Conta páginas totais antes de começar
    for pdf in arquivos_pdf:
        try:
            total_paginas += len(fitz.open(pdf))
        except Exception as e:
            print(f"Erro ao contar páginas de {pdf}: {e}")

    barra["maximum"] = total_paginas
    barra["value"] = 0

    total_arquivos = len(arquivos_pdf)

    for idx, pdf in enumerate(arquivos_pdf, start=1):
        try:
            processar_pdf_completo_gui(pdf, pasta_saida, 300, barra, label_status, idx, total_arquivos)
        except Exception as e:
            print(f"Erro ao processar {pdf}: {e}")

    messagebox.showinfo("Concluído", "Processamento finalizado com sucesso.")
    janela.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()

    arquivos_pdf = filedialog.askopenfilenames(
        title="Selecione os arquivos PDF para processar",
        filetypes=[("Arquivos PDF", "*.pdf")],
    )
    if not arquivos_pdf:
        messagebox.showinfo("Aviso", "Nenhum PDF selecionado. Encerrando.")
        exit()

    pasta_saida = filedialog.askdirectory(
        title="Selecione a pasta de destino para os arquivos extraídos"
    )
    if not pasta_saida:
        messagebox.showinfo("Aviso", "Nenhuma pasta selecionada. Encerrando.")
        exit()

    janela = tk.Toplevel()
    janela.title("Extração de QR Codes")
    janela.geometry("460x140")
    janela.resizable(False, False)

    label_status = tk.Label(janela, text="Pronto para iniciar...", anchor="w")
    label_status.pack(fill="x", padx=10, pady=10)

    barra = ttk.Progressbar(janela, orient="horizontal", length=400, mode="determinate")
    barra.pack(padx=20, pady=10)

    # Thread para manter GUI responsiva
    t = threading.Thread(
        target=iniciar_processamento,
        args=(arquivos_pdf, pasta_saida, janela, barra, label_status),
        daemon=True,
    )
    t.start()

    janela.mainloop()
