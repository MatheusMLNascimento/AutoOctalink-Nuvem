
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
import json
from openpyxl.styles import Alignment

try:
    import gspread
except ImportError:
    gspread = None

CONFIG_FILE = "config_octalink.json"

def carregar_config():
    default = {
        "historicos": ["SAIDA USO ECO"], 
        "codigos_mi": ["S500"],
        "tutorial_ativo": True,
        "cloud": {"spreadsheet_id": "", "worksheet_name": "Sheet1", "creds_json": ""}
    }
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
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
        self.root.title("Octalink Automator Pro - V3.0")
        self.root.geometry("800x900")
        self.config = carregar_config()
        self.dados_agrupados = None
        self.setup_ui()
        if self.config.get("tutorial_ativo"): self.exibir_tutorial()

    def exibir_tutorial(self):
        win = tk.Toplevel(self.root)
        win.title("📖 MANUAL DE CONFIGURAÇÃO (PASSO A PASSO)")
        win.geometry("650x700")
        win.attributes("-topmost", True)
        win.grab_set()
        
        # Container com scroll para caber toda a explicação
        canvas = tk.Canvas(win)
        scrollbar = ttk.Scrollbar(win, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        texto_tutorial = (
            "🚀 BEM-VINDO AO SEU AUTOMADOR OCTALINK PRO!\n"
            "Siga rigorosamente estes passos para ativar a Nuvem:\n\n"
            "------------------------------------------------------------------\n"
            "1️⃣ PASSO: OBTER A CHAVE (GOOGLE CLOUD)\n"
            "------------------------------------------------------------------\n"
            "1. Acesse: ://google.com\n"
            "2. No topo, crie um 'Novo Projeto'.\n"
            "3. No menu lateral, vá em 'APIs e Serviços' > 'Biblioteca'.\n"
            "4. Pesquise e ATIVE duas coisas: 'Google Sheets API' e 'Google Drive API'.\n"
            "5. Vá em 'Credenciais' > '+ Criar Credenciais' > 'Conta de Serviço'.\n"
            "6. Dê um nome qualquer e clique em 'Concluir'.\n"
            "7. Na lista de e-mails que aparecer, clique no e-mail azul que você criou.\n"
            "8. Vá na aba 'CHAVES' > 'Adicionar Chave' > 'Criar nova chave' > Escolha 'JSON'.\n"
            "9. Um arquivo será baixado. NÃO ABRA ELE, vamos usá-lo no Passo 2.\n\n"

            "------------------------------------------------------------------\n"
            "2️⃣ PASSO: CONFIGURAR O APLICATIVO\n"
            "------------------------------------------------------------------\n"
            "1. No App, clique em '📂 1. CARREGAR ARQUIVO JSON'.\n"
            "2. Escolha o arquivo que você baixou no passo anterior.\n"
            "3. O App mostrará um e-mail longo na tela. COPIE ESSE E-MAIL.\n\n"

            "------------------------------------------------------------------\n"
            "3️⃣ PASSO: DAR PERMISSÃO NA PLANILHA\n"
            "------------------------------------------------------------------\n"
            "1. Abra a sua Planilha do Google no seu navegador.\n"
            "2. Clique no botão azul 'COMPARTILHAR' no canto superior direito.\n"
            "3. Cole o e-mail que o App te deu e coloque-o como 'EDITOR'.\n"
            "4. Clique em 'Enviar'. (O bot não precisa aceitar nada, já está liberado!)\n\n"

            "------------------------------------------------------------------\n"
            "4️⃣ PASSO: O ID DA PLANILHA\n"
            "------------------------------------------------------------------\n"
            "1. Olhe para o endereço (URL) da sua planilha no navegador.\n"
            "2. Copie a URL e cole no campo 'Link da Planilha' no app.\n"
            "3. Clique em '✅ VALIDAR CONEXÃO'. Se ficar VERDE, está pronto!\n\n"
            "⚠️ DICA: Se mudar de planilha, basta trocar o link e Compartilhar de novo!"
        )

        tk.Label(scrollable_frame, text=texto_tutorial, justify="left", font=("Arial", 10), 
                 wraplength=580, padx=20, pady=20).pack()

        var_tutorial = tk.BooleanVar(value=True)
        tk.Checkbutton(scrollable_frame, text="Continuar mostrando este guia ao abrir o App", 
                       variable=var_tutorial, font=("Arial", 9, "bold")).pack(pady=10)
        
        def fechar():
            self.config["tutorial_ativo"] = var_tutorial.get()
            salvar_config(self.config)
            win.destroy()
        
        tk.Button(scrollable_frame, text="ENTENDI TUDO, VAMOS TRABALHAR!", bg="#4CAF50", fg="white", 
                  font=("Arial", 11, "bold"), command=fechar, pady=12, padx=30).pack(pady=20)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")



    def setup_ui(self):
        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=10)

        # --- SEÇÃO NUVEM ---
        cloud_frame = tk.LabelFrame(self.main_frame, text="⚙️ CONFIGURAÇÃO DA NUVEM (Fazer uma única vez)", fg="red", font=("Arial", 9, "bold"))
        cloud_frame.pack(fill="x", pady=5)

        # Botão para carregar arquivo JSON
        tk.Button(cloud_frame, text="📂 1. CARREGAR ARQUIVO JSON", command=self.importar_json_file, bg="#607D8B", fg="white").grid(row=0, column=0, columnspan=2, pady=5, padx=5, sticky="ew")

        tk.Label(cloud_frame, text="2. Link/ID da Planilha:").grid(row=1, column=0, padx=5, sticky="w")
        self.ent_sheet_id = tk.Entry(cloud_frame, width=50)
        self.ent_sheet_id.insert(0, self.config["cloud"]["spreadsheet_id"])
        self.ent_sheet_id.grid(row=1, column=1, pady=5, padx=5)

        tk.Label(cloud_frame, text="3. Nome da Aba:").grid(row=2, column=0, padx=5, sticky="w")
        self.ent_sheet_name = tk.Entry(cloud_frame, width=50)
        self.ent_sheet_name.insert(0, self.config["cloud"]["worksheet_name"])
        self.ent_sheet_name.grid(row=2, column=1, pady=5, padx=5)

        self.btn_test_cloud = tk.Button(cloud_frame, text="✅ VALIDAR CONEXÃO", bg="#2196F3", fg="white", command=self.testar_conexao)
        self.btn_test_cloud.grid(row=3, column=1, sticky="e", padx=5, pady=5)

        # --- SEÇÃO OPERACIONAL ---
        op_frame = tk.LabelFrame(self.main_frame, text="🚀 OPERAÇÃO DIÁRIA", font=("Arial", 10, "bold"))
        op_frame.pack(fill="x", pady=10)

        tk.Label(op_frame, text="Histórico:").pack()
        self.cb_hist = ttk.Combobox(op_frame, values=self.config["historicos"] + ["Outro..."], width=50, state="readonly")
        self.cb_hist.pack(pady=2)
        self.cb_hist.bind("<<ComboboxSelected>>", lambda e: self.toggle_outro(self.cb_hist, self.ent_outro_hist))
        self.cb_hist.current(0)
        self.ent_outro_hist = tk.Entry(op_frame, width=53, font=("Arial", 9, "italic"))

        tk.Label(op_frame, text="Código MI:").pack(pady=(5,0))
        self.cb_mi = ttk.Combobox(op_frame, values=self.config["codigos_mi"] + ["Outro..."], width=50, state="readonly")
        self.cb_mi.pack(pady=2)
        self.cb_mi.bind("<<ComboboxSelected>>", lambda e: self.toggle_outro(self.cb_mi, self.ent_outro_mi))
        self.cb_mi.current(0)
        self.ent_outro_mi = tk.Entry(op_frame, width=53, font=("Arial", 9, "italic"))

        self.btn_load = tk.Button(self.main_frame, text="📁 SELECIONAR ARQUIVO TXT", bg="#FF9800", fg="white", font=("Arial", 10, "bold"), command=self.ler_txt, height=2)
        self.btn_load.pack(fill="x", pady=10)

        # Preview aumentado
        self.txt_preview = tk.Text(self.main_frame, height=18, width=90, state="disabled", font=("Consolas", 8), bg="#F5F5F5")
        self.txt_preview.pack(pady=5)
        self.txt_preview.tag_configure("erro", foreground="red", font=("Consolas", 9, "bold"))
        self.txt_preview.tag_configure("header", foreground="blue", font=("Consolas", 8, "bold"))

        self.check_cloud = tk.BooleanVar(value=True)
        tk.Checkbutton(self.main_frame, text="ENVIAR PARA PLANILHA ONLINE", variable=self.check_cloud, font=("Arial", 10, "bold")).pack()

        self.btn_gerar = tk.Button(self.main_frame, text="🚀 FINALIZAR E GERAR TUDO", bg="#4CAF50", fg="white", font=("Arial", 12, "bold"), state="disabled", command=self.fluxo_final, height=2)
        self.btn_gerar.pack(fill="x", pady=10)

    def toggle_outro(self, combo, entry):
        if combo.get() == "Outro...": entry.pack(pady=2); entry.focus()
        else: entry.pack_forget()

    def importar_json_file(self):
        path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if not path: return
        try:
            with open(path, 'r', encoding='utf-8') as f:
                json_data = json.load(f)
                email_api = json_data.get("client_email", "")
                self.config["cloud"]["creds_json"] = json.dumps(json_data)
                
                # Pop-up impossível de ignorar com o e-mail
                win_email = tk.Toplevel(self.root)
                win_email.title("COPIE ESTE E-MAIL")
                win_email.attributes("-topmost", True)
                tk.Label(win_email, text="QUASE LÁ! Agora compartilhe sua planilha com:", font=("Arial", 10)).pack(pady=10, padx=20)
                
                ent_mail = tk.Entry(win_email, width=50, justify="center", font=("Arial", 10, "bold"), fg="blue")
                ent_mail.insert(0, email_api)
                ent_mail.pack(pady=10, padx=20)
                tk.Label(win_email, text="(Dê permissão de EDITOR na planilha Google)", font=("Arial", 8, "italic")).pack()
                
                tk.Button(win_email, text="OK, COPIEI!", command=win_email.destroy, bg="#2196F3", fg="white").pack(pady=15)
                
        except Exception as e:
            messagebox.showerror("Erro", f"Arquivo JSON inválido: {e}")

    def extrair_id(self, texto):
        """Extrai o ID da planilha se o usuário colar a URL inteira do navegador"""
        import re
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", texto)
        return match.group(1) if match else texto.strip()


    def ler_txt(self):
        path = filedialog.askopenfilename(filetypes=[("Arquivos de Texto", "*.txt")])
        if not path: return
        try:
            with open(path, 'r', encoding='utf-8') as f:
                linhas = [l.strip() for l in f.readlines() if l.strip()]
            
            raw = []
            i = 0
            self.txt_preview.config(state="normal")
            self.txt_preview.delete(1.0, tk.END)
            
            while i < len(linhas):
                try:
                    p_item = linhas[i].upper().split(' - ')
                    if len(p_item) < 2: raise ValueError(f"Linha {i+1}: Formato Código-Nome inválido.")
                    p_alm = linhas[i+1].upper().split(' - ')
                    idx_q = i+4 if (i+3 < len(linhas) and "/" in linhas[i+3]) else i+3
                    raw.append({"Item": p_item[0].strip(), "Desc": p_item[1].strip(), "Almox": p_alm[0].strip(), "Qtd": int(linhas[idx_q])})
                    i = idx_q + 1
                except Exception as e:
                    self.txt_preview.insert(tk.END, f"❌ ERRO: {e}\n", "erro")
                    messagebox.showerror("Ação Necessária", f"Erro no TXT: {e}")
                    return
            
            self.dados_agrupados = pd.DataFrame(raw).groupby(['Item', 'Desc', 'Almox']).agg({'Qtd':'sum','Item':'count'}).rename(columns={'Item':'Pacotes'}).reset_index()
            
            self.txt_preview.insert(tk.END, f"{'PRODUTO':<30} | {'CÓD':<12} | {'QTD':<5}\n", "header")
            self.txt_preview.insert(tk.END, "-"*55 + "\n")
            for _, r in self.dados_agrupados.iterrows():
                self.txt_preview.insert(tk.END, f"{r['Desc'][:28]:<30} | {r['Item']:<12} | {r['Qtd']}\n")
            
            self.btn_gerar.config(state="normal")
            self.txt_preview.config(state="disabled")
        except Exception as e:
            messagebox.showerror("Erro Crítico", str(e))

    def fluxo_final(self):
        h = self.ent_outro_hist.get().upper() if self.cb_hist.get() == "Outro..." else self.cb_hist.get()
        m = self.ent_outro_mi.get().upper() if self.cb_mi.get() == "Outro..." else self.cb_mi.get()
        id_plani = self.extrair_id(self.ent_sheet_id.get())
        
        self.config["cloud"]["spreadsheet_id"] = id_plani
        self.config["cloud"]["worksheet_name"] = self.ent_sheet_name.get().strip()
        salvar_config(self.config)

        # --- ENVIO NUVEM (Com 10 Colunas - Inclui Nome) ---
        if self.check_cloud.get():
            try:
                # Login moderno
                creds_dict = json.loads(self.config["cloud"]["creds_json"])
                client = gspread.service_account_from_dict(creds_dict)
                
                # Acesso à aba
                sheet = client.open_by_key(id_plani).worksheet(self.config["cloud"]["worksheet_name"])
                
                data_h = datetime.now().strftime('%d/%m/%Y')
                
                # Montagem das 10 colunas (A até J)
                envio = [[data_h, m, h, r['Item'], r['Almox'], "", "", r['Qtd'], "", r['Desc']] 
                         for _, r in self.dados_agrupados.iterrows()]
                
                sheet.append_rows(envio, value_input_option='USER_ENTERED')
                messagebox.showinfo("Nuvem", "🚀 Relatório enviado com sucesso!")
            except Exception as e: 
                messagebox.showerror("Falha na Nuvem", f"O arquivo local foi gerado, mas a nuvem falhou: {e}")


        # --- ARQUIVO LOCAL (Com 9 Colunas - SEM NOME para o Octa) ---
        nome_arq = f"IMPORT_{m}_{h}_{datetime.now().strftime('%H%M%S')}.xlsx"
        
        # Criando apenas as 9 colunas que o Octa aceita
        dados_locais = []
        for _, r in self.dados_agrupados.iterrows():
            dados_locais.append([
                datetime.now().strftime('%d/%m/%Y'), m, h, r['Item'], r['Almox'], "", "", r['Qtd'], ""
            ])
            
        df_local = pd.DataFrame(dados_locais, columns=[
            'Data movimento', 'Cód. MI', 'Histórico', 'Cód. Item', 
            'Almoxarifado', 'Almoxarifado transf.', 'Unidade Medida', 'Qtde', 'Valor'
        ])
        
        df_local.to_excel(nome_arq, index=False)
        messagebox.showinfo("Sucesso", f"Arquivo para o Octa gerado: {nome_arq}")


    def testar_conexao(self):
        try:
            # 1. Limpa o ID (caso seja URL)
            id_bruto = self.ent_sheet_id.get().strip()
            id_limpo = self.extrair_id(id_bruto)
            
            # 2. Pega o JSON da configuração
            if not self.config["cloud"]["creds_json"]:
                raise ValueError("Você precisa carregar o arquivo JSON primeiro!")
            
            creds_dict = json.loads(self.config["cloud"]["creds_json"])
            
            # 3. Autenticação Direta (Método Moderno - Sem oauth2client)
            client = gspread.service_account_from_dict(creds_dict)
            
            # 4. Tenta abrir a planilha
            client.open_by_key(id_limpo)
            
            # 5. Sucesso!
            messagebox.showinfo("Sucesso", "🚀 CONEXÃO ESTABELECIDA!\nO Google aceitou sua chave.")
            self.btn_test_cloud.config(bg="#4CAF50", text="✅ CONECTADO")
            
            # Atualiza interface e salva
            self.ent_sheet_id.delete(0, tk.END)
            self.ent_sheet_id.insert(0, id_limpo)
            self.config["cloud"]["spreadsheet_id"] = id_limpo
            salvar_config(self.config)

        except Exception as e:
            messagebox.showerror("Erro de Conexão", f"O Google recusou o acesso.\n\nDetalhe: {e}")
            self.btn_test_cloud.config(bg="#F44336", text="❌ FALHA NA CONEXÃO")

    def enviar_para_nuvem(self, hist, mi):
        """Versão corrigida: envia os dados para o Google Sheets"""
        if not self.config["cloud"]["creds_json"]:
            raise ValueError("Credenciais JSON não encontradas!")
            
        creds_dict = json.load(self.config["cloud"]["creds_json"])
        
        # Login moderno e estável
        client = gspread.service_account_from_dict(creds_dict)
        
        id_plani = self.extrair_id(self.ent_sheet_id.get())
        sheet = client.open_by_key(id_plani).worksheet(self.config["cloud"]["worksheet_name"])
        
        dt = datetime.now().strftime('%d/%m/%Y')
        
        # PREPARAÇÃO DOS DADOS (Ajustado: 'hist' e 'mi' agora batem com os nomes das variáveis)
        envio = []
        for _, r in self.dados_agrupados.iterrows():
            linha = [
                dt,          # A: Data
                mi,          # B: Cód. MI
                hist,        # C: Histórico
                r['Item'],   # D: Cód. Item
                r['Almox'],  # E: Almoxarifado
                "",          # F: Almoxarifado transf. (Vazio)
                "",          # G: Unidade Medida (Vazio)
                r['Qtd'],    # H: Qtde
                "",          # I: Valor (Vazio)
                r['Desc']    # J: Nome (O relatório da nuvem tem o nome!)
            ]
            envio.append(linha)
        
        # Envia tudo de uma vez para ser rápido
        sheet.append_rows(envio, value_input_option='USER_ENTERED')


if __name__ == "__main__":
    root = tk.Tk(); app = OctalinkApp(root); root.mainloop()