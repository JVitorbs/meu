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

def calcular_horas(row, tipo: str) -> Tuple[List[HorarioPar], str]:
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
        if tipo == "Dirigido" and horarios[i][0] in [1, 6]:  # Se for um horário de início para "Dirigido"
            if start_time is None:
                start_time = horarios[i][1]
        elif tipo == "Descanso" and horarios[i][0] in [2, 7]:  # Se for um horário de início para "Descanso"
            if start_time is None:
                start_time = horarios[i][1]
        elif tipo == "Espera" and horarios[i][0] == 3:  # Se for um horário de início para "Espera"
            if start_time is None:
                start_time = horarios[i][1]
        elif start_time is not None and ((tipo == "Dirigido" and horarios[i][0] not in [1, 6]) or \
                                         (tipo == "Descanso" and horarios[i][0] not in [2, 7]) or \
                                         (tipo == "Espera" and horarios[i][0] != 3)):  # Se for um horário de término
            # Encontrar o próximo horário que pode ser um início
            j = i
            while j < len(horarios) and ((tipo == "Dirigido" and horarios[j][0] not in [1, 6]) or \
                                         (tipo == "Descanso" and horarios[j][0] not in [2, 7]) or \
                                         (tipo == "Espera" and horarios[j][0] != 3)):
                j += 1
            
            # Se j-1 está dentro do intervalo e start_time não é None
            if j > 0 and start_time is not None:
                stop_time = horarios[j-1][1] if j - 1 < len(horarios) else None
                if stop_time and (stop_time - start_time) > timedelta(seconds=60):  # Verifica se a diferença é válida
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

        resultados_dirigido = []
        resultados_descanso = []
        resultados_espera = []

        for index, row in df.iterrows():
            # Processa a aba "Horario Dirigido"
            pares_dirigido, total_horas_dirigido = calcular_horas(row, "Dirigido")
            resultado_linha_dirigido = {"Data": row[0].strftime('%d/%m/%Y')}
            
            if pares_dirigido:
                for i, par in enumerate(pares_dirigido):
                    resultado_linha_dirigido[f"Start_{i+1}"] = str(par.start)
                    resultado_linha_dirigido[f"Stop_{i+1}"] = str(par.stop)
                resultado_linha_dirigido["Total de Horas"] = total_horas_dirigido
            else:
                max_pares = (len(row) - 1) // 2
                for i in range(max_pares):
                    resultado_linha_dirigido[f"Start_{i+1}"] = ""
                    resultado_linha_dirigido[f"Stop_{i+1}"] = ""
                resultado_linha_dirigido["Total de Horas"] = ""

            resultados_dirigido.append(resultado_linha_dirigido)

            # Processa a aba "Horario de Descanso"
            pares_descanso, total_horas_descanso = calcular_horas(row, "Descanso")
            resultado_linha_descanso = {"Data": row[0].strftime('%d/%m/%Y')}
            
            if pares_descanso:
                for i, par in enumerate(pares_descanso):
                    resultado_linha_descanso[f"Start_{i+1}"] = str(par.start)
                    resultado_linha_descanso[f"Stop_{i+1}"] = str(par.stop)
                resultado_linha_descanso["Total de Horas"] = total_horas_descanso
            else:
                max_pares = (len(row) - 1) // 2
                for i in range(max_pares):
                    resultado_linha_descanso[f"Start_{i+1}"] = ""
                    resultado_linha_descanso[f"Stop_{i+1}"] = ""
                resultado_linha_descanso["Total de Horas"] = ""

            resultados_descanso.append(resultado_linha_descanso)

            # Processa a aba "Horario de Espera"
            pares_espera, total_horas_espera = calcular_horas(row, "Espera")
            resultado_linha_espera = {"Data": row[0].strftime('%d/%m/%Y')}
            
            if pares_espera:
                for i, par in enumerate(pares_espera):
                    resultado_linha_espera[f"Start_{i+1}"] = str(par.start)
                    resultado_linha_espera[f"Stop_{i+1}"] = str(par.stop)
                resultado_linha_espera["Total de Horas"] = total_horas_espera
            else:
                max_pares = (len(row) - 1) // 2
                for i in range(max_pares):
                    resultado_linha_espera[f"Start_{i+1}"] = ""
                    resultado_linha_espera[f"Stop_{i+1}"] = ""
                resultado_linha_espera["Total de Horas"] = ""

            resultados_espera.append(resultado_linha_espera)

        df_resultados_dirigido = pd.DataFrame(resultados_dirigido)
        df_resultados_descanso = pd.DataFrame(resultados_descanso)
        df_resultados_espera = pd.DataFrame(resultados_espera)

        df_resultados_dirigido['Data'] = pd.to_datetime(df_resultados_dirigido['Data'], dayfirst=True)
        df_resultados_dirigido = df_resultados_dirigido.sort_values(by='Data')
        df_resultados_dirigido['Data'] = df_resultados_dirigido['Data'].dt.strftime('%d/%m/%Y')

        df_resultados_descanso['Data'] = pd.to_datetime(df_resultados_descanso['Data'], dayfirst=True)
        df_resultados_descanso = df_resultados_descanso.sort_values(by='Data')
        df_resultados_descanso['Data'] = df_resultados_descanso['Data'].dt.strftime('%d/%m/%Y')

        df_resultados_espera['Data'] = pd.to_datetime(df_resultados_espera['Data'], dayfirst=True)
        df_resultados_espera = df_resultados_espera.sort_values(by='Data')
        df_resultados_espera['Data'] = df_resultados_espera['Data'].dt.strftime('%d/%m/%Y')

        colunas_ordenadas_dirigido = ['Data']
        max_pares = max(len(r) for r in resultados_dirigido) // 2
        for i in range(max_pares):
            colunas_ordenadas_dirigido.append(f"Start_{i+1}")
            colunas_ordenadas_dirigido.append(f"Stop_{i+1}")
        colunas_ordenadas_dirigido.append('Total de Horas')

        df_resultados_dirigido = df_resultados_dirigido.reindex(columns=colunas_ordenadas_dirigido)

        colunas_ordenadas_descanso = ['Data']
        max_pares = max(len(r) for r in resultados_descanso) // 2
        for i in range(max_pares):
            colunas_ordenadas_descanso.append(f"Start_{i+1}")
            colunas_ordenadas_descanso.append(f"Stop_{i+1}")
        colunas_ordenadas_descanso.append('Total de Horas')

        df_resultados_descanso = df_resultados_descanso.reindex(columns=colunas_ordenadas_descanso)

        colunas_ordenadas_espera = ['Data']
        max_pares = max(len(r) for r in resultados_espera) // 2
        for i in range(max_pares):
            colunas_ordenadas_espera.append(f"Start_{i+1}")
            colunas_ordenadas_espera.append(f"Stop_{i+1}")
        colunas_ordenadas_espera.append('Total de Horas')

        df_resultados_espera = df_resultados_espera.reindex(columns=colunas_ordenadas_espera)

        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                 filetypes=[("Excel files", "*.xlsx *.xls")],
                                                 title="Salvar arquivo Excel de resultados")
        
        if save_path:
            with pd.ExcelWriter(save_path) as writer:
                df_resultados_dirigido.to_excel(writer, sheet_name='Horaras Trabalhadas', index=False)
                df_resultados_descanso.to_excel(writer, sheet_name='Horario de Descanso', index=False)
                df_resultados_espera.to_excel(writer, sheet_name='Horario de Espera', index=False)
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
