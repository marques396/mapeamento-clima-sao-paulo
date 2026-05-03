import os
import time
import requests  # Biblioteca para a API
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font
import tkinter as tk
from tkinter import ttk, messagebox
from dotenv import load_dotenv  # <--- Adicione este import

# 1. Carrega as configurações do arquivo .env
load_dotenv()

# 2. Captura a chave da variável de ambiente
API_KEY = os.getenv("OPENWEATHER_API_KEY")

# ADICIONE ESTA LINHA TEMPORARIAMENTE:
print(f">>> DEBUG: A chave carregada é: {API_KEY}")

# O restante do seu código (funções e interface) vem logo abaixo...
class ClimaApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Previsão do Tempo - São Paulo")
        self.root.geometry("400x350")
        
        self.style = ttk.Style()
        self.style.configure('TButton', font=('Arial', 12), padding=10)
        self.style.configure('TLabel', font=('Arial', 12), background='#f0f0f0')
        
        self.main_frame = ttk.Frame(root, padding="20")
        self.main_frame.pack(expand=True, fill=tk.BOTH)
        
        self.label_titulo = ttk.Label(self.main_frame, text="Previsão do Tempo", font=('Arial', 16, 'bold'))
        self.label_titulo.pack(pady=10)

        self.label_cidade = ttk.Label(self.main_frame, text="Digite a cidade:")
        self.label_cidade.pack(pady=(0, 5))
        self.entry_cidade = ttk.Entry(self.main_frame, width=30, font=('Arial', 12))
        self.entry_cidade.pack(pady=(0, 10))
        self.entry_cidade.insert(0, "São Paulo")

        self.btn_buscar = ttk.Button(self.main_frame, text="Buscar Previsão", command=self.executar_coleta)
        self.btn_buscar.pack(pady=(0, 20))
        
        self.frame_resultados = ttk.LabelFrame(self.main_frame, text="Resultados", padding=10)
        self.frame_resultados.pack(fill=tk.BOTH, expand=True)
        
        self.label_data = ttk.Label(self.frame_resultados, text="Data: -")
        self.label_data.pack(anchor='w')
        
        self.label_temp = ttk.Label(self.frame_resultados, text="Temperatura: -°C")
        self.label_temp.pack(anchor='w')
        
        self.label_umid = ttk.Label(self.frame_resultados, text="Umidade: -%")
        self.label_umid.pack(anchor='w')
        
        self.label_status = ttk.Label(self.frame_resultados, text="Condição: -")
        self.label_status.pack(anchor='w')
        
        self.label_status_oper = ttk.Label(self.main_frame, text="", foreground='blue')
        self.label_status_oper.pack(pady=5)

    def executar_coleta(self):
        try:
            self.btn_buscar['state'] = tk.DISABLED
            self.label_status_oper['text'] = "Conectando à API..."
            self.root.update()
            
            cidade = self.entry_cidade.get()
            dados = self.coletar_dados(cidade) # Agora não passa mais o 'driver'
            
            if dados:
                self.atualizar_interface(dados)
                self.salvar_dados_excel(dados, cidade)
                self.label_status_oper['text'] = "Dados coletados com sucesso!"
                messagebox.showinfo("Sucesso", "Dados coletados e salvos com sucesso!")
            else:
                self.label_status_oper['text'] = "Falha ao coletar dados"
                messagebox.showerror("Erro", "Não foi possível obter os dados. Verifique a chave da API ou a cidade.")
        
        except Exception as e:
            self.label_status_oper['text'] = f"Erro: {str(e)}"
            messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")
        finally:
            self.btn_buscar['state'] = tk.NORMAL
            self.root.update()

    def atualizar_interface(self, dados):
        self.label_data['text'] = f"Data: {dados['data']}"
        self.label_temp['text'] = f"Temperatura: {dados['temperatura']}°C"
        self.label_umid['text'] = f"Umidade: {dados['umidade']}%"
        self.label_status['text'] = f"Condição: {dados['status']}"

    def coletar_dados(self, cidade):
        try:
            api_key = os.getenv("OPENWEATHER_API_KEY")
            url = f"https://api.openweathermap.org/data/2.5/weather?q={cidade}&appid={api_key}&units=metric&lang=pt_br"
            response = requests.get(url)
            if response.status_code == 200:
                dados_json = response.json()
                return {
                    'data': time.strftime("%d/%m/%Y %H:%M"),
                    'temperatura': dados_json['main']['temp'],
                    'umidade': dados_json['main']['humidity'],
                    'status': dados_json['weather'][0]['description'].capitalize()
                }
            else:
                return None
        except Exception as e:
            print(f"Erro na API: {str(e)}")
            return None

    def salvar_dados_excel(self, dados, cidade, nome_arquivo="clima.xlsx"):
        try:
            caminho_completo_arquivo = os.path.join(os.path.dirname(os.path.abspath(__file__)), nome_arquivo)

            if os.path.exists(caminho_completo_arquivo):
                wb = load_workbook(caminho_completo_arquivo)
                sheet = wb.active
            else:
                wb = Workbook()
                sheet = wb.active
                sheet['A1'] = 'Cidade'
                sheet['B1'] = 'Data'
                sheet['C1'] = 'Temp (°C)'
                sheet['D1'] = 'Umidade (%)'
                sheet['E1'] = 'Status'
                for cell in sheet["1:1"]:
                    cell.font = Font(bold=True)
            
            nova_linha = [cidade, dados['data'], dados['temperatura'], dados['umidade'], dados['status']]
            sheet.append(nova_linha)
            
            # Ajuste automático de largura das colunas
            for column in sheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except: pass
                sheet.column_dimensions[column_letter].width = max_length + 2
            
            wb.save(caminho_completo_arquivo)
        except Exception as e:
            print(f"Erro ao salvar: {str(e)}")
            raise

if __name__ == "__main__":
    root = tk.Tk()
    app = ClimaApp(root)
    root.mainloop()