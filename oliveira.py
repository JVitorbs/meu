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

    # Iterando sobre cada célula da linha (exceto a primeira coluna que é a data)
    for celula in row[1:]:
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
                # Só armazena o horário de início se não houver um horário de início já armazenado
                if start_time is None:
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

        # Converte a primeira coluna para datetime, interpretando-a como 'dia/mês/ano'
        df.iloc[:, 0] = pd.to_datetime(df.iloc[:, 0], dayfirst=True)

        # Lista para armazenar os resultados
        resultados = []

        # Itera sobre cada linha do DataFrame
        for index, row in df.iterrows():
            # Calcula os pares de horários e o total de horas para a linha atual
            pares, total_horas = calcular_horas(row)
            
            # Verifica se há pares de horários válidos; caso contrário, ignora a linha
            if pares is None:
                continue
            
            # Cria um dicionário para armazenar a data e os pares de horários
            resultado_linha = {"Data": row[0].strftime('%d/%m/%Y')}  # Armazena a data formatada
            
            # Adiciona cada par de horários a colunas 'Start_n' e 'Stop_n'
            for i, par in enumerate(pares):
                resultado_linha[f"Start_{i+1}"] = str(par.start)
                resultado_linha[f"Stop_{i+1}"] = str(par.stop)
            
            # Adiciona o total de horas após os pares de horários
            resultado_linha["Total de Horas"] = total_horas
            
            # Adiciona o resultado processado à lista de resultados
            resultados.append(resultado_linha)

        # Cria um DataFrame para os resultados
        df_resultados = pd.DataFrame(resultados)

        # Reorganiza as colunas para garantir que "Total de Horas" fique após todos os 'Start' e 'Stop'
        colunas_ordenadas = ['Data']  # Inicia com a coluna de data
        for i in range(len(resultados[0]) // 2):  # Calcula a quantidade de pares
            colunas_ordenadas.append(f"Start_{i+1}")
            colunas_ordenadas.append(f"Stop_{i+1}")
        colunas_ordenadas.append('Total de Horas')  # Adiciona a coluna "Total de Horas"

        # Reorganiza o DataFrame para as colunas ordenadas
        df_resultados = df_resultados.reindex(columns=colunas_ordenadas)

        # Abre uma janela para salvar o arquivo Excel com os resultados
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                 filetypes=[("Excel files", "*.xlsx *.xls")],
                                                 title="Salvar arquivo Excel de resultados")
        
        # Se o usuário escolheu um caminho para salvar
        if save_path:
            # Salva o DataFrame de resultados em um novo arquivo Excel
            df_resultados.to_excel(save_path, index=False)
            print(f"Resultados salvos em: {save_path}")

# Configura a interface gráfica
root = tk.Tk()
root.title("Processador de Horas")  # Título da janela
root.geometry("300x150")  # Tamanho da janela

# Botão para o usuário selecionar o arquivo Excel
btn_selecionar = tk.Button(root, text="Selecionar Arquivo Excel", command=abrir_e_processar_arquivo)
btn_selecionar.pack(pady=20)  # Adiciona o botão à janela e define um espaçamento vertical

# Inicia o loop principal da interface gráfica
root.mainloop()
