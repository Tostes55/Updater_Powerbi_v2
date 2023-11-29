import tkinter as tk
import customtkinter
import requests
import logging
import locale
import pandas as pd
import os
import shutil
from token_movidesk import token_api
from tkinter import *
from tkinter import messagebox
from tkcalendar import *
from datetime import datetime




locale.setlocale(locale.LC_TIME, 'pt_BR.utf8')

# Configura o logging
logging.basicConfig(filename='log_app.txt', filemode='w', format='%(asctime)s - %(levelname)s - %(message)s', level=logging.INFO)


class Ambiente:
    def __init__(self):
        self.data_inicio = None
        self.data_final = None
        self.token = token_api

    def get_data(self):
        url = 'https://api.movidesk.com/public/v1/tickets'

        query_params = [
            '$select=id,ownerTeam,serviceFirstLevel,serviceSecondLevel,serviceThirdLevel,status,origin,resolvedIn',
            '$orderby=resolvedIn',
            '$expand=Clients($expand=Organization),owner',
            f'token={self.token}',
            f"$filter=resolvedIn ge {self.data_inicio} and resolvedIn le {self.data_final}"
        ]

        query_string = '&'.join(query_params)
        full_url = f"{url}?{query_string}"

        logging.info(f'Making GET request to URL: {full_url}')

        try:
            progressbar.start()
            response = requests.get(full_url)
            response.raise_for_status()  # Lança uma exceção para códigos de status de erro
            data = response.json()
            print(f"Realizado Request")
            return data
        except requests.exceptions.RequestException as e:
            logging.error(f'Erro na solicitação HTTP: {e}')
            return None

    def save_as_xlsx(self, data):
        if data is not None and isinstance(data, list):
            try:
                filtered_data = []
                for item in data:
                    filtered_item = {
                        "resolvedIn": item.get("resolvedIn", ""),
                        "origin": item.get("origin", ""),
                        "status": item.get("status", ""),
                        "serviceThirdLevel": item.get("serviceThirdLevel", ""),
                        "serviceSecondLevel": item.get("serviceSecondLevel", ""),
                        "serviceFirstLevel": item.get("serviceFirstLevel", ""),
                        "ownerTeam": item.get("ownerTeam", ""),
                        "id": item.get("id", ""),
                        "businessNameOwner": item.get("owner", {}).get("businessName", ""),
                        "business_company": item.get("clients")[0].get("organization", {}).get("businessName", ""),
                        "businessNameClient": item.get("clients", [{}])[0].get("businessName", "")
                    }
                    filtered_data.append(filtered_item)

                if filtered_data:
                    output_file = 'output.xlsx'
                    backup_folder = 'backup'
                    today_folder = os.path.join(backup_folder, datetime.now().strftime('%Y-%m-%d'))

                    if not os.path.exists(today_folder):
                        os.makedirs(today_folder)
                        logging.info(f"Pasta '{today_folder}' criada com sucesso para o backup de hoje.")

                    backup_path = os.path.join(today_folder, f"{output_file.split('.')[0]}_backup.xlsx")
                    shutil.copy2(output_file, backup_path)
                    logging.info(f"Backup do arquivo '{output_file}' criado em '{backup_path}'.")

                    df_existing = pd.read_excel(output_file)

                    df_to_append = []
                    for item in filtered_data:
                        if item['id'] not in df_existing['id'].tolist():
                            df_to_append.append(item)

                    if df_to_append:
                        df_updated = pd.concat([df_existing, pd.DataFrame(df_to_append)], ignore_index=True)
                        df_updated.to_excel(output_file, index=False)
                        logging.info(f'Dados salvos em {output_file}')
                        return True
                    else:
                        logging.warning('Nenhum novo dado para adicionar.')
                        return False
                else:
                    logging.warning('Nenhum dado filtrado encontrado.')
                    return False
            except Exception as ex:
                logging.error(f'Erro ao processar e salvar dados: {ex}')
                return False
        else:
            logging.error('Nenhum dado válido recebido.')
            return False

    def update_data(self):
        progressbar.start()
        selected_date_entry = start_date_entry.get_date()
        formatted_date_entry = selected_date_entry.strftime("%Y-%m-%dT%H:%M:%S.00z")

        selected_date_final = end_date_entry.get_date()
        formatted_date_final = selected_date_final.strftime("%Y-%m-%dT%H:%M:%S.00z")

        self.data_inicio = formatted_date_entry
        self.data_final = formatted_date_final

        logging.info('Comando get pronto para ser executado')

        data = self.get_data()
        if data:
            messagebox.showinfo("Sucesso", "Dados recebidos da API !")
            print(f"Os dados obtidos foram: {data}")

            # Chama o método save_as_xlsx para salvar os dados
            saved_successfully = self.save_as_xlsx(data)
            if saved_successfully:
                progressbar.stop()
                messagebox.showinfo("Sucesso", "Dados salvos com sucesso!")
            else:
                progressbar.stop()
                messagebox.showerror("Erro", "Falha ao salvar os dados. Verifique o log para mais informações.")
        else:
            progressbar.stop()
            messagebox.showerror("Erro", "Falha ao obter os dados. Verifique o log para mais informações.")

def update_excel(data):
    try:
        
        # Caminho do arquivo original e da pasta de backup
        output_file = 'output.xlsx'
        backup_folder = 'backup'

        # Verificar se a pasta de backup existe; se não, criar
        if not os.path.exists(backup_folder):
            os.makedirs(backup_folder)
            logging.info(f"Pasta de backup '{backup_folder}' criada com sucesso.")

        # Criar o caminho completo para o arquivo de backup
        backup_path = os.path.join(backup_folder, f"{output_file.split('.')[0]}_backup.xlsx")

        # Copiar o arquivo original para a pasta de backup
        shutil.copy2(output_file, backup_path)
        logging.info(f"Backup do arquivo '{output_file}' criado em '{backup_path}'.")

        # Carregar o arquivo XLSX existente em um DataFrame
        df_existing = pd.read_excel('output.xlsx')
        
        
        # Criar um DataFrame com os novos dados
        df_new = pd.DataFrame(data['result'])  # Supondo que 'result' seja a chave dos dados retornados

        # Verificar as colunas disponíveis nos DataFrames
        print("Colunas em df_existing:", df_existing.columns)
        print("Colunas em df_new:", df_new.columns)
        logging.info("Colunas em df_existing:" + str(df_existing.columns))
        logging.info("Colunas em df_new:" + str(df_new.columns))

        # Verificar os conjuntos de IDs
        existing_ids = set(df_existing['id']) if 'id' in df_existing.columns else set()
        new_ids = set(df_new['id']) if 'id' in df_new.columns else set()
        logging.info("Existing IDs:" + str(existing_ids))
        logging.info("New IDs:" + str(new_ids))

        print("Existing IDs:", existing_ids)
        print("New IDs:", new_ids)

        # Adicionar apenas os novos dados ao DataFrame existente
        df_to_append = df_new[~df_new['id'].isin(existing_ids)]

        # Verificar se existem novos dados para adicionar
        if not df_to_append.empty:
            # Adicionar os novos dados ao DataFrame existente
            df_updated = pd.concat([df_existing, df_to_append], ignore_index=True)

            # Salvar o DataFrame atualizado de volta no arquivo XLSX
            df_updated.to_excel('output.xlsx', index=False)
            print("Dados atualizados com sucesso no arquivo XLSX!")
        else:
            print("Não há novos dados para adicionar.")
    except Exception as e:
        print(f"Erro ao atualizar dados no arquivo XLSX: {e}")

# Chamada da função para atualizar o arquivo XLSX com os novos dados
        update_excel(data)
        progressbar.stop()

# Cria uma janela raiz
janela = customtkinter.CTk()
janela.title('Updater Power BI')
janela.geometry("450x450")
janela.maxsize(width=450, height=450)
janela.minsize(width=450, height=450)
janela.resizable(width=False, height=False)
janela._set_appearance_mode("Dark")


title_label = customtkinter.CTkLabel(janela, text='Updater Power BI', font=("Arial Black",22))
title_label.place(x=120, y=60)


start_date_label = customtkinter.CTkLabel(janela, text='Data Inicial:', font=("Arial",15))
start_date_label.place(x=120, y=160)

start_date_entry = DateEntry(janela,date_pattern='dd/mm/yyyy')
start_date_entry.place(x=220, y=165)

end_date_label = customtkinter.CTkLabel(janela, text='Data Final:', font=("Arial",15))
end_date_label.place(x=120, y=200)

end_date_entry = DateEntry(janela,date_pattern='dd/mm/yyyy') 
end_date_entry.place(x=220, y=205)

progressbar=customtkinter.CTkProgressBar(janela, orientation="horizontal")
progressbar.place(x=120, y=260)
progressbar.set(0)
mode="indeterminate",
determinate_speed=5,
indeterminate_speed=.5,

ambiente = Ambiente()

def update_data():
    selected_date_entry = start_date_entry.get_date()
    formatted_date_entry = selected_date_entry.strftime("%Y-%m-%dT%H:%M:%S.00z")

    selected_date_final = end_date_entry.get_date()
    formatted_date_final = selected_date_final.strftime("%Y-%m-%dT%H:%M:%S.00z")

    ambiente.data_inicio = formatted_date_entry
    ambiente.data_final = formatted_date_final

    logging.info('Comando get pronto para ser executado')

    ambiente.update_data()

update_button = customtkinter.CTkButton(janela, text='Atualizar', command=update_data)
#update_button.pack(padx=50, pady=20)
#update_button.grid(row=3, column=0, columnspan=2, pady=30)
update_button.place(x=160, y=300)

janela.mainloop()
