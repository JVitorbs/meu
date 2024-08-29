import pandas as pd
import re
from datetime import timedelta
import tkinter as tk
from tkinter import filedialog
import webbrowser
from typing import List, Tuple

class HorarioPar:
    def __init__(self, start: timedelta, stop: timedelta):
        self.start = start
        self.stop = stop

    def duracao(self) -> timedelta:
        return self.stop - self.start

    def duracao_valida(self) -> bool:
        return self.duracao() > timedelta(seconds=60)

def calcular_horas(row) -> Tuple[List[HorarioPar], str]:
    pares: List[HorarioPar] = []
    start_time = None
    horarios: List[Tuple[int, timedelta]] = []

    # Extrai os horários e números das células da linha
    for celula in row[1:]:
        padrao = r'(\d+)-(\d{2}):(\d{2}):(\d{2})'
        match = re.match(padrao, str(celula))
        
        if match:
            numero = int(match.group(1))
            hora = int(match.group(2))
            minuto = int(match.group(3))
            segundo = int(match.group(4))
            tempo_atual = timedelta(hours=hora, minutes=minuto, seconds=segundo)
            horarios.append((numero, tempo_atual))
    
    # Ordena a lista de horários pelo tempo
    horarios.sort(key=lambda x: x[1])
    
    i = 0
    while i < len(horarios):
        if horarios[i][0] in [1, 6]:  # Se for um horário de início
            if start_time is None:
                start_time = horarios[i][1]
        
        elif start_time is not None and horarios[i][0] not in [1, 6]:  # Se for um horário de término
            j = i
            while j < len(horarios) and horarios[j][0] not in [1, 6]:  # Ignora stops consecutivos, mantendo o último
                j += 1
            
            if j <= len(horarios):  # Verifica se encontrou um stop
                stop_time = horarios[j-1][1]  # Considera o último stop consecutivo válido
                if (stop_time - start_time) > timedelta(seconds=60):  # Verifica se a diferença é válida
                    pares.append(HorarioPar(start=start_time, stop=stop_time))
                    start_time = None
                i = j - 1  # Move o índice para o próximo potencial par
                
        i += 1

    if not pares:
        return None, None
    
    total_horas = sum([par.duracao() for par in pares], timedelta())
    total_horas_formatado = str(total_horas)
    return pares, total_horas_formatado

def abrir_e_processar_arquivo():
    file_path = filedialog.askopenfilename(title="Selecione o arquivo Excel", 
                                           filetypes=[("Excel files", "*.xlsx *.xls")])
    
    if file_path:
        df = pd.read_excel(file_path)
        df.iloc[:, 0] = pd.to_datetime(df.iloc[:, 0], dayfirst=True)

        resultados = []

        for index, row in df.iterrows():
            pares, total_horas = calcular_horas(row)
            resultado_linha = {"Data": row[0].strftime('%d/%m/%Y')}
            
            if pares:
                for i, par in enumerate(pares):
                    resultado_linha[f"Start_{i+1}"] = str(par.start)
                    resultado_linha[f"Stop_{i+1}"] = str(par.stop)
                resultado_linha["Total de Horas"] = total_horas
            else:
                max_pares = (len(row) - 1) // 2
                for i in range(max_pares):
                    resultado_linha[f"Start_{i+1}"] = ""
                    resultado_linha[f"Stop_{i+1}"] = ""
                resultado_linha["Total de Horas"] = ""

            resultados.append(resultado_linha)

        df_resultados = pd.DataFrame(resultados)
        df_resultados['Data'] = pd.to_datetime(df_resultados['Data'], dayfirst=True)
        df_resultados = df_resultados.sort_values(by='Data')
        df_resultados['Data'] = df_resultados['Data'].dt.strftime('%d/%m/%Y')
        
        colunas_ordenadas = ['Data']
        max_pares = max(len(r) for r in resultados) // 2
        for i in range(max_pares):
            colunas_ordenadas.append(f"Start_{i+1}")
            colunas_ordenadas.append(f"Stop_{i+1}")
        colunas_ordenadas.append('Total de Horas')

        df_resultados = df_resultados.reindex(columns=colunas_ordenadas)

        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                 filetypes=[("Excel files", "*.xlsx *.xls")],
                                                 title="Salvar arquivo Excel de resultados")
        
        if save_path:
            df_resultados.to_excel(save_path, index=False)
            print(f"Resultados salvos em: {save_path}")

def abrir_link(event):
    webbrowser.open("https://www.linkedin.com/in/joão-vitor-batista-silva-50b280279")

root = tk.Tk()
root.title("Processador de Horas")
root.geometry("400x200")

info_label = tk.Label(root, text="Autor: João Vitor Batista Silva\nwww.linkedin.com/in/joão-vitor-batista-silva-50b280279",
                      fg="blue", cursor="hand2")
info_label.pack(pady=10)

info_label.bind("<Button-1>", abrir_link)

btn_selecionar = tk.Button(root, text="Selecionar Arquivo Excel", command=abrir_e_processar_arquivo)
btn_selecionar.pack(pady=20)

root.mainloop()
