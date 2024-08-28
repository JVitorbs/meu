#vc vai ter que instalar as bibliotecas, tá aqui o pip: pip install pandas openpyxl tk

import pandas as pd
import re
from datetime import timedelta
import tkinter as tk
from tkinter import filedialog
from typing import List, Tuple

# Classe para representar um par de horários, com tempo de início (start) e término (stop)
class HorarioPar:
    def __init__(self, start: timedelta, stop: timedelta):
        self.start = start
        self.stop = stop

    # Método para calcular a duração entre start e stop
    def duracao(self) -> timedelta:
        return self.stop - self.start

# Função para processar cada linha do DataFrame e calcular os pares de horário
def calcular_horas(row) -> Tuple[List[HorarioPar], str]:
    pares: List[HorarioPar] = []  # Lista que armazenará os pares de horários válidos
    start_time = None  # Variável para armazenar o horário de início

    # Iterando sobre cada célula da linha
    for celula in row:
        # Regex para identificar o padrão "número-hora:minuto:segundo"
        padrao = r'(\d+)-(\d{2}):(\d{2}):(\d{2})'
        match = re.match(padrao, str(celula))  # Tenta casar a célula com o padrão
        
        if match:
            # Extrai o número e o tempo da célula
            numero = int(match.group(1))
            hora = int(match.group(2))
            minuto = int(match.group(3))
            segundo = int(match.group(4))
            # Converte o horário para um objeto timedelta
            tempo_atual = timedelta(hours=hora, minutes=minuto, seconds=segundo)
            
            # Se o número for 1 ou 6, trata-se de um horário de início (start)
            if numero in [1, 6]:
                start_time = tempo_atual  # Armazena o horário de início
                
            # Se o número for 2, 3, 4, ou 5 e houver um horário de início armazenado,
            # trata-se de um horário de término (stop)
            elif numero in [2, 3, 4, 5] and start_time is not None:
                # Cria um par de horários e o adiciona à lista de pares
                pares.append(HorarioPar(start=start_time, stop=tempo_atual))
                start_time = None  # Reseta o start_time para preparar para o próximo par

    if not pares:  # Se não houver pares de horários válidos, retorna None
        return None, None
    
    # Calcula o total de horas somando todas as durações dos pares de horários
    total_horas = sum([par.duracao() for par in pares], timedelta())

    # Formata o total de horas em "HH:MM:SS" para exibição
    total_horas_formatado = str(total_horas)
    return pares, total_horas_formatado  # Retorna a lista de pares e o total de horas formatado

# Função para abrir o seletor de arquivos e processar o arquivo escolhido
def abrir_e_processar_arquivo():
    # Abre uma janela para o usuário selecionar o arquivo Excel
    file_path = filedialog.askopenfilename(title="Selecione o arquivo Excel", 
                                           filetypes=[("Excel files", "*.xlsx *.xls")])
    
    # Se o usuário selecionou um arquivo
    if file_path:
        # Lê o arquivo Excel em um DataFrame
        df = pd.read_excel(file_path)

        # Itera sobre cada linha do DataFrame
        for index, row in df.iterrows():
            # Calcula os pares de horários e o total de horas para a linha atual
            pares, total_horas = calcular_horas(row)
            
            # Verifica se há pares de horários válidos; caso contrário, ignora a linha
            if pares is None:
                continue
            
            # Exibe os resultados no terminal
            print(f"Linha {index + 1}:")
            for par in pares:
                print(f"  Par de Horários: Start = {str(par.start)}, Stop = {str(par.stop)}")
            print(f"  Total de Horas: {total_horas}\n")

# Configura a interface gráfica
root = tk.Tk()
root.title("Processador de Horas")  # Título da janela
root.geometry("300x150")  # Tamanho da janela

# Botão para o usuário selecionar o arquivo Excel
btn_selecionar = tk.Button(root, text="Selecionar Arquivo Excel", command=abrir_e_processar_arquivo)
btn_selecionar.pack(pady=20)  # Adiciona o botão à janela e define um espaçamento vertical

# Inicia o loop principal da interface gráfica
root.mainloop()
