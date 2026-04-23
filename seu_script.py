import sys
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
import json
from openpyxl.styles import Alignment

# Bibliotecas de Nuvem
try:
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
except ImportError:
    gspread = None

CONFIG_FILE = "config_octalink.json"

def carregar_config():
    default = {
        "historicos": ["SAIDA USO ECO"], 
        "codigos_mi": ["S500"],
        "tutorial_ativo": True,
        "cloud": {
            "spreadsheet_id": "",
            "worksheet_name": "Sheet1",
            "creds_json": "" # O conteúdo do JSON colado aqui
        }
    }
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                # Garante que chaves novas existam em arquivos antigos
                for key in default:
                    if key not in data: data[key] = default[key]
                return data
        except: return default
    return default

def salvar_config(config):
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
    except: pass

class OctalinkApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Octalink Automator Pro + Cloud")
        self.root.geometry("750x850")
        self.config = carregar_config()
        self.dados_agrupados = None
        self.setup_ui()
        
        if self.config.get("tutorial_ativo"):
            self.exibir_tutorial()

    def exibir_tutorial(self):
        win = tk.Toplevel(self.root)
        win.title("Guia de Configuração Rápida")
        win.geometry("550x500")
        
        # --- CONFIGURAÇÃO DE DESTAQUE TOTAL ---
        win.attributes("-topmost", True) # Fica por cima de TODAS as janelas do Windows
        win.grab_set()                  # Bloqueia cliques na tela principal do App
        
        win.configure(padx=20, pady=20)
        
        texto_guia = (
            "👋 BEM-VINDO!\n\n"
            "Este guia é obrigatório para a primeira configuração:\n\n"
            "1. CONEXÃO:\n"
            "   Cole o ID da planilha e o texto do JSON no topo.\n\n"
            "2. VALIDAÇÃO:\n"
            "   Clique no botão de 'Validar Conexão' até ele ficar VERDE.\n\n"
            "3. OPERAÇÃO:\n"
            "   Carregue o TXT e clique em 'Finalizar'.\n"
        )

        tk.Label(win, text=texto_guia, justify="left", font=("Arial", 10, "bold"), wraplength=500).pack()

        var_tutorial = tk.BooleanVar(value=True)
        tk.Checkbutton(win, text="Mostrar este guia ao iniciar o App", variable=var_tutorial).pack(pady=20)

        def finalizar():
            self.config["tutorial_ativo"] = var_tutorial.get()
            salvar_config(self.config)
            win.destroy()

        tk.Button(win, text="ENTENDI, PODE LIBERAR O APP", bg="#4CAF50", fg="white", 
                  font=("Arial", 10, "bold"), command=finalizar, pady=10, padx=20).pack()


    def setup_ui(self):
        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=10)

        # Cabeçalho
        tk.Label(self.main_frame, text="OCTALINK AUTOMATOR PRO", font=("Arial", 14, "bold"), fg="#1565C0").pack(pady=10)

        # --- CONTAINER NUVEM REFORMULADO ---
        cloud_frame = tk.LabelFrame(self.main_frame, text="⚙️ CONFIGURAÇÃO ÚNICA DA NUVEM", fg="#D32F2F", font=("Arial", 9, "bold"))
        cloud_frame.pack(fill="x", pady=5, padx=5)

        # ID da Planilha com Ajuda
        tk.Label(cloud_frame, text="1. ID da Planilha Google:", font=("Arial", 8, "bold")).grid(row=0, column=0, sticky="w", padx=5)
        self.ent_sheet_id = tk.Entry(cloud_frame, width=55, fg="grey")
        self.ent_sheet_id.insert(0, self.config["cloud"]["spreadsheet_id"] or "Cole aqui o código longo da URL da sua planilha")
        self.ent_sheet_id.grid(row=0, column=1, pady=5, padx=5)
        self.ent_sheet_id.bind("<FocusIn>", lambda e: self.on_entry_click(self.ent_sheet_id, "Cole aqui"))

        # JSON com Ajuda
        tk.Label(cloud_frame, text="2. Conteúdo do JSON:", font=("Arial", 8, "bold")).grid(row=1, column=0, sticky="w", padx=5)
        self.ent_json_creds = tk.Entry(cloud_frame, width=55, show="*", fg="grey")
        self.ent_json_creds.insert(0, self.config["cloud"]["creds_json"] or "Abra o arquivo baixado no bloco de notas e cole o texto todo aqui")
        self.ent_json_creds.grid(row=1, column=1, pady=5, padx=5)
        self.ent_json_creds.bind("<FocusIn>", lambda e: self.on_entry_click(self.ent_json_creds, "Abra o arquivo"))

        # Botão de Teste (Crucial para o usuário não errar)
        self.btn_test_cloud = tk.Button(cloud_frame, text="✅ Validar Dados da Nuvem", bg="#607D8B", fg="white", 
                                        font=("Arial", 8, "bold"), command=self.testar_conexao_imediata)
        self.btn_test_cloud.grid(row=2, column=1, sticky="e", pady=5, padx=5)
        # --- CONTAINER OPERACIONAL (Onde o usuário foca no dia a dia) ---
        op_frame = tk.Frame(self.main_frame)
        op_frame.pack(fill="x", pady=10)

        # Seleção de Histórico
        tk.Label(op_frame, text="O que você está fazendo? (Histórico)", font=("Arial", 9, "bold")).pack()
        self.cb_hist = ttk.Combobox(op_frame, values=self.config["historicos"] + ["Outro..."], width=50, state="readonly")
        self.cb_hist.pack(pady=5)
        self.cb_hist.bind("<<ComboboxSelected>>", lambda e: self.toggle_outro(self.cb_hist, self.ent_outro_hist))
        self.cb_hist.current(0)
        
        self.ent_outro_hist = tk.Entry(op_frame, width=53, font=("Arial", 9, "italic"), fg="blue")
        # (Lógica de toggle_outro e placeholder se mantém)

        # Botão de Carregar (Impossível de não ver)
        self.btn_load = tk.Button(self.main_frame, text="1. SELECIONAR ARQUIVO TXT", bg="#2196F3", fg="white", 
                                  font=("Arial", 11, "bold"), command=self.ler_txt, height=2)
        self.btn_load.pack(fill="x", pady=10)

        # Preview com instrução
        tk.Label(self.main_frame, text="👀 Verifique se os dados abaixo estão corretos:", font=("Arial", 8, "italic")).pack()
        self.txt_preview = tk.Text(self.main_frame, height=10, width=85, state="disabled", font=("Consolas", 8), bg="#EEE")
        self.txt_preview.pack(pady=5)

        # Botão Final
        self.check_cloud = tk.BooleanVar(value=True)
        tk.Checkbutton(self.main_frame, text="Também salvar na Planilha Online", variable=self.check_cloud, font=("Arial", 9, "bold")).pack()

        self.btn_gerar = tk.Button(self.main_frame, text="2. CONCLUIR E GERAR TUDO", bg="#4CAF50", fg="white", 
                                   font=("Arial", 12, "bold"), state="disabled", command=self.fluxo_final, height=2)
        self.btn_gerar.pack(fill="x", pady=10)

    def on_entry_click(self, entry, text_check):
        """Limpa o texto de ajuda quando o usuário clica no campo"""
        if text_check in entry.get():
            entry.delete(0, tk.END)
            entry.config(fg="black", show="" if "ID" in str(entry) else "*")

    def testar_conexao_imediata(self):
        """Tenta conectar na hora para dar feedback ao usuário"""
        try:
            # Pega os dados atuais da tela
            id_plani = self.ent_sheet_id.get().strip()
            json_text = self.ent_json_creds.get().strip()
            
            if "Cole aqui" in id_plani or "Abra o arquivo" in json_text:
                raise ValueError("Você precisa preencher os campos com seus dados reais primeiro!")

            # Tenta autenticar
            scope = ["https://google.com", "https://googleapis.com"]
            creds = ServiceAccountCredentials.from_json_keyfile_dict(json.loads(json_text), scope)
            client = gspread.authorize(creds)
            
            # Tenta abrir a planilha
            client.open_by_key(id_plani)
            
            messagebox.showinfo("Sucesso!", "CONEXÃO ESTABELECIDA!\n\nO app conseguiu acessar sua planilha Google com sucesso.")
            self.btn_test_cloud.config(bg="#4CAF50", text="CONECTADO")
            
            # Já salva nas configs para não perder
            self.config["cloud"]["spreadsheet_id"] = id_plani
            self.config["cloud"]["creds_json"] = json_text
            salvar_config(self.config)

        except Exception as e:
            msg_erro = str(e)
            if "permission_denied" in msg_erro.lower():
                res = "Erro de Permissão: Você esqueceu de compartilhar a planilha com o e-mail que está dentro do JSON!"
            elif "not found" in msg_erro.lower():
                res = "Erro de ID: Esse ID de planilha não existe. Copie novamente da URL."
            else:
                res = f"Erro Técnico: Verifique se colou o JSON corretamente.\n\nDetalhe: {e}"
            
            messagebox.showerror("Falha na Configuração", res)
            self.btn_test_cloud.config(bg="#F44336", text="FALHA NA CONEXÃO")

    def toggle_outro(self, combo, entry):
        if combo.get() == "Outro...": entry.pack(pady=2)
        else: entry.pack_forget()

    def ler_txt(self):
        try:
            path = filedialog.askopenfilename(filetypes=[("Arquivos de Texto", "*.txt")])
            if not path: return

            with open(path, 'r', encoding='utf-8') as f:
                linhas = [l.strip() for l in f.readlines() if l.strip()]

            raw_data = []
            i = 0
            self.txt_preview.config(state="normal")
            self.txt_preview.delete(1.0, tk.END)

            while i < len(linhas):
                try:
                    num_ln = i + 1
                    # Validação de Bloco
                    p_item = linhas[i].upper().split(' - ')
                    if len(p_item) < 2: raise ValueError(f"Linha {num_ln}: Formato CÓD-NOME inválido.")
                    
                    p_alm = linhas[i+1].upper().split(' - ')
                    
                    # QTD Lógica (Data ou Direto)
                    idx_qtd = i + 4 if (i+3 < len(linhas) and "/" in linhas[i+3]) else i + 3
                    if idx_qtd >= len(linhas): raise ValueError(f"Linha {num_ln}: Dados incompletos.")
                    
                    raw_data.append({
                        "Item": p_item[0].strip(), 
                        "Desc": p_item[1].strip(), 
                        "Almox": p_alm[0].strip(), 
                        "Qtd": int(linhas[idx_qtd])
                    })
                    i = idx_qtd + 1
                except Exception as e:
                    self.txt_preview.insert(tk.END, f"❌ ERRO: {str(e)}\n", "erro")
                    self.btn_gerar.config(state="disabled")
                    return

            df = pd.DataFrame(raw_data)
            self.dados_agrupados = df.groupby(['Item', 'Desc', 'Almox']).agg({'Qtd': 'sum', 'Item': 'count'}).rename(columns={'Item': 'Pacotes'}).reset_index()
            
            # Renderizar Preview
            self.txt_preview.insert(tk.END, f"{'PRODUTO':<25} | {'CÓD':<10} | {'ALM':<5} | {'QTD':<5}\n", "header")
            for _, r in self.dados_agrupados.iterrows():
                n = (r['Desc'][:23] + "..") if len(r['Desc']) > 23 else r['Desc']
                self.txt_preview.insert(tk.END, f"{n:<25} | {r['Item']:<10} | {r['Almox']:<5} | {r['Qtd']:<5}\n")
            
            self.txt_preview.config(state="disabled")
            self.btn_gerar.config(state="normal")

        except Exception as e:
            messagebox.showerror("Erro Crítico", f"Falha ao carregar TXT: {e}")
            
    def obter_campos_finais(self):
        """Captura os valores atuais da UI (Combos ou Entradas 'Outro')"""
        h = self.ent_outro_hist.get().upper().strip() if self.cb_hist.get() == "Outro..." else self.cb_hist.get()
        m = self.ent_outro_mi.get().upper().strip() if self.cb_mi.get() == "Outro..." else self.cb_mi.get()
        
        if not h or "Digite" in h or not m or "Digite" in m:
            return None
        return {"hist": h, "mi": m}

    def salvar_novas_configuracoes(self, h, m):
        """Atualiza o JSON com novos históricos e dados de nuvem"""
        if self.cb_hist.get() == "Outro..." and h not in self.config["historicos"]:
            self.config["historicos"].append(h)
        if self.cb_mi.get() == "Outro..." and m not in self.config["codigos_mi"]:
            self.config["codigos_mi"].append(m)
            
        self.config["cloud"]["spreadsheet_id"] = self.ent_sheet_id.get().strip()
        self.config["cloud"]["worksheet_name"] = self.ent_sheet_name.get().strip()
        self.config["cloud"]["creds_json"] = self.ent_json_creds.get().strip()
        salvar_config(self.config)

    def fluxo_final(self):
        """Gerencia a execução local e nuvem com Try/Except em camadas"""
        campos = self.obter_campos_finais()
        if not campos:
            messagebox.showwarning("Aviso", "Preencha os campos de Histórico/MI corretamente.")
            return

        # Salva o estado atual das configurações
        self.salvar_novas_configuracoes(campos['hist'], campos['mi'])

        status_msg = []
        
        # 1. Tentar Envio para Nuvem (se habilitado)
        if self.check_cloud.get():
            try:
                self.enviar_para_nuvem(campos['hist'], campos['mi'])
                status_msg.append("✅ Dados enviados para a Nuvem!")
            except Exception as e:
                status_msg.append(f"❌ Erro Nuvem: {str(e)}")
                messagebox.showerror("Falha na Nuvem", f"Erro ao conectar com Google Sheets:\n{e}")

        # 2. Tentar Exportação Local (Sempre gera, exceto se você quiser criar um modo 'apenas nuvem')
        try:
            nome_local = self.exportar_excel_local(campos['hist'], campos['mi'])
            status_msg.append(f"✅ Arquivo local gerado: {nome_local}")
        except Exception as e:
            status_msg.append(f"❌ Erro Local: {str(e)}")

        messagebox.showinfo("Processamento Finalizado", "\n".join(status_msg))

    def enviar_para_nuvem(self, hist, mi):
        """Conecta ao Google Sheets usando o JSON embutido na config"""
        if not gspread:
            raise ImportError("Biblioteca gspread não encontrada. Instale com 'pip install gspread oauth2client'")

        creds_str = self.config["cloud"]["creds_json"]
        if not creds_str:
            raise ValueError("O campo JSON de Credenciais está vazio!")

        try:
            creds_dict = json.loads(creds_str)
            scope = ["https://google.com", "https://googleapis.com"]
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
            client = gspread.authorize(creds)
            
            sheet = client.open_by_key(self.config["cloud"]["spreadsheet_id"])
            worksheet = sheet.worksheet(self.config["cloud"]["worksheet_name"])
            
            data_hoje = datetime.now().strftime('%d/%m/%Y')
            linhas_envio = [[data_hoje, mi, hist, r['Item'], r['Desc'], r['Almox'], r['Qtd']] 
                            for _, r in self.dados_agrupados.iterrows()]
            
            worksheet.append_rows(linhas_envio, value_input_option='USER_ENTERED')

        except json.JSONDecodeError:
            raise ValueError("O conteúdo do JSON de credenciais está inválido (erro de digitação).")
        except gspread.exceptions.SpreadsheetNotFound:
            raise ValueError("ID da Planilha não encontrado. Verifique o ID e se compartilhou com o e-mail do JSON.")
        except (gspread.exceptions.APIError, Exception) as e:
            # Tratamento para falta de internet ou erro de servidor
            erro_str = str(e).lower()
            if "connection" in erro_str or "timeout" in erro_str:
                raise ConnectionError("Falha de conexão: Verifique sua internet ou se o Google está bloqueado.")
            else:
                raise Exception(f"Erro inesperado na nuvem: {e}")

    def exportar_excel_local(self, h, m):
        """Gera o arquivo .xlsx no computador"""
        h_limpo = "".join([c for c in h if c.isalnum() or c in (' ', '_', '-')]).strip()
        m_limpo = "".join([c for c in m if c.isalnum() or c in (' ', '_', '-')]).strip()
        nome_arq = f"IMPORT_{m_limpo}_{h_limpo}_{datetime.now().strftime('%H%M%S')}.xlsx"
        
        final_df = pd.DataFrame({
            'Data movimento': [datetime.now()] * len(self.dados_agrupados),
            'Cód. MI': m,
            'Histórico': h,
            'Cód. Item': self.dados_agrupados['Item'],
            'Descrição': self.dados_agrupados['Desc'],
            'Almoxarifado': self.dados_agrupados['Almox'],
            'Qtde': self.dados_agrupados['Qtd']
        })

        with pd.ExcelWriter(nome_arq, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False, sheet_name='Importacao')
            ws = writer.sheets['Importacao']
            for cell in ws['A']: cell.number_format = 'DD/MM/YYYY'
            for cell in ws[1]: cell.alignment = Alignment(horizontal='center')
        
        return nome_arq

if __name__ == "__main__":
    root = tk.Tk()
    app = OctalinkApp(root)
    root.mainloop()
