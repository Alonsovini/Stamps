# Organizar Life 

import tkinter as tk  # Importa a blibioeca tkinter para criar a interface gráfica
from tkinter import filedialog  # Importa a função filedialog para abrir uma janela de seleção de arquivo
import fitz  # Importa o módulo fitz do PyMuPDF para trabalhar com arquivos PDF
import re  # Importa o módulo re para expressões regulares
from openpyxl import Workbook, load_workbook  # IMporta a classe Workbook do openpycl para criar arquivos Excel
import os  # Importa o módulo os para lidar com operações relacionadas ao sistema operacional

def extrair_numeros_notas(pdf_path):  # Função para extrair os números das notas de um arquivo PDF
    notas = []  # Inicializa uma lista vazia para armazenar os números das notas

    # Abre o arquivo PDF e itera sobre suas páginas
    with fitz.open(pdf_path) as pdf:
        for page_num in range(pdf.page_count):
            page = pdf.load_page(page_num)  # Carrega a página atual
            text = page.get_text()  # Extrai o texto da página
            # Usa expressão regular para encontrar os números das notas no texto
            # matches = re.findall(r'Número da Nota\s*[\r\n]+\s*(\d{6})', text)
            matches = re.findall(r'Nº: \s*(\d{6})', text)

            if matches:
                for nota in matches:
                    notas.append(nota.strip())  # Adiciona o número da nota à lista de notas

    return notas  # Retorna a lista de notas

def salvar_em_excel(notas,excel_path):
    # Função par salvar os números das notas em um arquivo Excel
    wb = Workbook()  # Cria um novo Workbook do openpyxl
    ws = wb.active  # Seleciona a planilha ativa
    ws.append(['Número das Notas'])  # Adiciona um cabeçlho a planilha
    
    # for nota in notas:
    #    nota_sem_quatro_primeiros = nota[4:] # Exclui os 4 primeiros digitos do número da nota
    #    ws.append([nota_sem_quatro_primeiros])
    # Itera sobre os números das notas e os adiciona a planilha
    for nota in notas:
        ws.append([nota])

    wb.save(excel_path)  # Salva o arquivo Excel
    return excel_path  # Retorna o caminho do arquivo Excel salvo

def selecionar_arquivo():
    # Função para selecionar uma pasta contendo arquivos PDF
    root = tk.Tk()  # Cria uma instância da janela principal do tkinter 
    root.withdraw()  # Esconde a janela principal

    # Abre uma janela de seleção de pasta e retorna o caminho da pasta selecionada
    arquivo = filedialog.askopenfilename(title="Selecione o arquivo PDF", filetypes=(("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*")))
    return arquivo

def main():
    # Função principal
    pdf_path = selecionar_arquivo()  # Seleciona o arquivo PDF usando uma janela de seleção de arquivos
    if not pdf_path:
        print("Nenhum arquivo selecionado.")  #Exibe uma mensagem se nenhum arquivo foi selecionado 
        return
    
    # Pede ao usúario para selecionar o local e o nome do arquivo Excel de Saída
    excel_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=(("Arquivo Excel", "*.xlsx"), ("Todos os arquivos", "*.*")))
    if not excel_path:
        print("Nenhum local de salvamento selecionado.")  # Exibe uma mensagem se nenhum local de salvamento foi selecionado

    # Extrai os números das notas do arquivo PDF selecionado 
    notas = extrair_numeros_notas(pdf_path)
    #  Salva os números das notas em um arquivo Excel
    excel_salvo = salvar_em_excel(notas, excel_path)

    print("Notas extraídas e salvas em Excel com sucesso!")  # Exibe uma mensagem de sucesso
    os.startfile(excel_salvo)  # Abre o arquivo Excel salvo

if __name__ == "__main__":
    main()  # Chama a função principal se o script for executado diretamente