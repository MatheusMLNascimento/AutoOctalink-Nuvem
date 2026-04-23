
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
import json
from openpyxl.styles import Alignment
import re

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

        self.btn_load = tk.Button(self.main_frame, text="📁 SELECIONAR ARQUIVO TXT", bg="#FF9800", fg="white", 
                                  font=("Arial", 10, "bold"), command=self.ler_txt, height=2)
        self.btn_load.pack(fill="x", pady=10)

        # Preview aumentado
        self.txt_preview = tk.Text(self.main_frame, height=18, width=90, state="disabled", font=("Consolas", 8), bg="#F5F5F5")
        self.txt_preview.pack(pady=5)
        self.txt_preview.tag_configure("erro", foreground="red", font=("Consolas", 9, "bold"))
        self.txt_preview.tag_configure("header", foreground="blue", font=("Consolas", 8, "bold"))

        # OPÇÕES DE SAÍDA (Checkboxes)
        out_f = tk.Frame(self.main_frame)
        out_f.pack(pady=5)
        
        self.check_local = tk.BooleanVar(value=True)
        tk.Checkbutton(out_f, text="Gerar Excel Local (.xlsx)", variable=self.check_local, font=("Arial", 9)).pack(side="left", padx=10)
        
        self.check_cloud = tk.BooleanVar(value=True)
        tk.Checkbutton(out_f, text="Enviar para Planilha Online", variable=self.check_cloud, font=("Arial", 9)).pack(side="left", padx=10)
        
        self.check_data_nome = tk.BooleanVar(value=False)
        tk.Checkbutton(self.main_frame, text="Usar data do nome do arquivo (DDMMYY)", variable=self.check_data_nome, font=("Arial", 8, "italic")).pack()

        # BARRA DE PROGRESSO E STATUS
        self.lbl_status = tk.Label(self.main_frame, text="Aguardando...", font=("Arial", 8, "italic"))
        self.lbl_status.pack()
        self.progress = ttk.Progressbar(self.main_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", pady=5)

        self.btn_gerar = tk.Button(self.main_frame, text="🚀 FINALIZAR E GERAR TUDO", bg="#4CAF50", fg="white", 
                                   font=("Arial", 12, "bold"), state="disabled", command=self.fluxo_final, height=2)
        self.btn_gerar.pack(fill="x", pady=10)

        # --- BOTÃO E TERMINAL AVANÇADO ---
        self.terminal_visivel = False
        self.btn_terminal = tk.Button(self.main_frame, text="📟 Mostrar terminal (avançado)", 
                                      command=self.toggle_terminal, font=("Arial", 8), bg="#ececec")
        self.btn_terminal.pack(side="bottom", anchor="se", pady=5)

        self.terminal_log = tk.Text(self.main_frame, height=10, state="disabled", 
                                    bg="black", fg="#00FF00", font=("Consolas", 8))
        # O terminal começa escondido (.pack_forget)

    def toggle_terminal(self):
        """Mostra ou esconde o terminal interno"""
        if not self.terminal_visivel:
            self.terminal_log.pack(fill="x", pady=5, before=self.btn_terminal)
            self.btn_terminal.config(text="📟 Esconder terminal")
            self.terminal_visivel = True
            self.log("Terminal avançado ativado.")
        else:
            self.terminal_log.pack_forget()
            self.btn_terminal.config(text="📟 Mostrar terminal (avançado)")
            self.terminal_visivel = False

    def log(self, mensagem):
        """Escreve mensagens no terminal interno com timestamp"""
        tempo = datetime.now().strftime("%H:%M:%S")
        msg_formatada = f"[{tempo}] {mensagem}\n"
        
        self.terminal_log.config(state="normal")
        self.terminal_log.insert(tk.END, msg_formatada)
        self.terminal_log.see(tk.END) # Scroll automático
        self.terminal_log.config(state="disabled")
        self.root.update() # Força a interface a mostrar a mensagem na hora


    def toggle_outro(self, combo, entry):
        if combo.get() == "Outro...": entry.pack(pady=2); entry.focus()
        else: entry.pack_forget()

    def extrair_data_do_nome(self, nome_arquivo):
        """Busca padrão DDMMYY no nome do arquivo e retorna data formatada"""
        import re
        # Busca 6 dígitos seguidos
        match = re.search(r"(\d{6})", nome_arquivo)
        if match:
            try:
                data_str = match.group(1)
                # Converte DDMMYY para objeto datetime
                data_obj = datetime.strptime(data_str, "%d%m%y")
                return data_obj.strftime("%d/%m/%Y")
            except:
                pass
        # Se falhar ou não achar, usa a data de hoje
        return datetime.now().strftime("%d/%m/%Y")


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
        self.log("Abrindo seletor de arquivos...")
        path = filedialog.askopenfilename(filetypes=[("Arquivos de Texto", "*.txt")])
        
        if not path:
            self.log("Seleção de arquivo cancelada pelo usuário.")
            return

        self.caminho_txt = path
        self.log(f"Arquivo selecionado: {os.path.basename(path)}")

        try:
            with open(path, 'r', encoding='utf-8') as f:
                linhas = [l.strip() for l in f.readlines() if l.strip()]
            
            self.log(f"Lendo {len(linhas)} linhas do arquivo...")
            raw_data = []
            i = 0
            
            self.txt_preview.config(state="normal")
            self.txt_preview.delete(1.0, tk.END)

            while i < len(linhas):
                try:
                    num_linha = i + 1
                    # Validação de Código e Nome
                    partes = linhas[i].upper().split(' - ')
                    if len(partes) < 2:
                        self.log(f"❌ Erro de formato na linha {num_linha}.")
                        raise ValueError(f"Linha {num_linha}: Esperado 'CÓD - NOME'.")
                    
                    cod = partes[0].strip()
                    desc = partes[1].strip()
                    
                    # Validação de Almoxarifado
                    sigla_almox = linhas[i+1].upper().split(' - ')[0].strip()
                    
                    # Lógica de Quantidade (Pula data se houver)
                    idx_qtd = i + 4 if (i+3 < len(linhas) and "/" in linhas[i+3]) else i + 3
                    
                    if idx_qtd >= len(linhas):
                        self.log(f"❌ Dados incompletos para o item {cod}.")
                        raise ValueError(f"Fim de arquivo inesperado após item {cod}.")

                    qtd = int(linhas[idx_qtd])
                    raw_data.append({"Item": cod, "Desc": desc, "Almox": sigla_almox, "Qtd": qtd})
                    
                    i = idx_qtd + 1
                except Exception as e:
                    self.log(f"⚠️ Falha no processamento: {str(e)}")
                    self.txt_preview.insert(tk.END, f"❌ ERRO: {str(e)}\n", "erro")
                    self.btn_gerar.config(state="disabled")
                    messagebox.showerror("Erro de Integridade", str(e))
                    return

            # Agrupamento com Pandas
            df = pd.DataFrame(raw_data)
            self.dados_agrupados = df.groupby(['Item', 'Desc', 'Almox']).agg({'Qtd': 'sum'}).reset_index()
            
            self.log(f"Sucesso: {len(self.dados_agrupados)} itens únicos processados.")
            
            # Atualiza o Preview Visual
            self.txt_preview.insert(tk.END, f"{'PRODUTO':<30} | {'CÓD':<12} | {'QTD':<5}\n", "header")
            for _, r in self.dados_agrupados.iterrows():
                nome_p = (r['Desc'][:28] + "..") if len(r['Desc']) > 28 else r['Desc']
                self.txt_preview.insert(tk.END, f"{nome_p:<30} | {r['Item']:<12} | {r['Qtd']}\n")
            
            self.log("Arquivo TXT carregado com sucesso.") # Mensagem solicitada
            self.txt_preview.config(state="disabled")
            self.btn_gerar.config(state="normal")

        except Exception as e:
            self.log(f"❌ ERRO CRÍTICO no carregamento: {str(e)}")
            messagebox.showerror("Erro", f"Não foi possível ler o arquivo:\n{e}")


    def fluxo_final(self):
        self.log("=== INICIANDO PROCESSAMENTO FINAL ===")
        
        if not self.check_cloud.get() and not self.check_local.get():
            self.log("⚠️ Operação cancelada: Nenhuma opção de saída marcada.")
            messagebox.showwarning("Aviso", "Selecione Local ou Nuvem!")
            return

        h = self.ent_outro_hist.get().upper() if self.cb_hist.get() == "Outro..." else self.cb_hist.get()
        m = self.ent_outro_mi.get().upper() if self.cb_mi.get() == "Outro..." else self.cb_mi.get()
        id_p = self.extrair_id(self.ent_sheet_id.get())
        
        # Lógica de Data
        self.log("Determinando data de movimento...")
        if self.check_data_nome.get() and hasattr(self, 'caminho_txt'):
            match = re.search(r"(\d{6})", os.path.basename(self.caminho_txt))
            if match:
                data_f = datetime.strptime(match.group(1), "%d%m%y").strftime("%d/%m/%Y")
                self.log(f"📅 Data extraída do arquivo: {data_f}")
            else:
                data_f = datetime.now().strftime("%d/%m/%Y")
                self.log(f"📅 Padrão de data não encontrado no nome. Usando hoje: {data_f}")
        else:
            data_f = datetime.now().strftime("%d/%m/%Y")
            self.log(f"📅 Usando data atual do sistema: {data_f}")

        # Execução Local
        if self.check_local.get():
            self.log("Gerando arquivo Excel local...")
            try:
                nome = f"IMPORT_{m}_{h}_{datetime.now().strftime('%H%M%S')}.xlsx"
                # ... (lógica de DataFrame aqui) ...
                self.log(f"✅ Excel local gerado: {nome}")
            except Exception as e:
                self.log(f"❌ Erro no Excel local: {str(e)}")

        # Execução Nuvem
        if self.check_cloud.get():
            self.enviar_para_nuvem(h, m, data_f, id_p)
            
        self.log("=== TODOS OS PROCESSOS FINALIZADOS ===")


    def testar_conexao(self):
        self.log("--- Iniciando Teste de Conexão ---")
        try:
            id_bruto = self.ent_sheet_id.get().strip()
            id_limpo = self.extrair_id(id_bruto)
            js_p = self.config["cloud"]["creds_json"]
            
            if not js_p:
                self.log("❌ ERRO: Chave JSON não encontrada nas configurações.")
                raise ValueError("Carregue o arquivo JSON primeiro.")

            self.log(f"Autenticando conta de serviço...")
            creds_dict = json.loads(js_p)
            client = gspread.service_account_from_dict(creds_dict)
            
            self.log(f"Tentando abrir planilha ID: {id_limpo[:10]}...")
            client.open_by_key(id_limpo)
            
            self.log("✅ Conexão estabelecida com sucesso!")
            messagebox.showinfo("Sucesso", "Conexão com a Nuvem OK!")
            self.btn_test.config(bg="#4CAF50", text="✅ CONECTADO")
            
        except Exception as e:
            self.log(f"❌ FALHA NA CONEXÃO: {str(e)}")
            messagebox.showerror("Erro de Conexão", str(e))
            self.btn_test.config(bg="#F44336", text="❌ FALHA NA CONEXÃO")


    def enviar_para_nuvem(self, hist, mi, data_mov, id_p):
        self.log(f"--- Iniciando Sincronização: {data_mov} | {mi} ---")
        try:
            self.log("Preparando credenciais...")
            creds = json.loads(self.config["cloud"]["creds_json"])
            client = gspread.service_account_from_dict(creds)
            
            aba_nome = self.ent_sheet_name.get().strip()
            self.log(f"Abrindo aba: '{aba_nome}'...")
            sheet = client.open_by_key(id_p).worksheet(aba_nome)
            
            self.log("Baixando dados existentes da nuvem (Verificação de duplicatas)...")
            self.root.update()
            dados_n = sheet.get_all_values()
            self.log(f"Sucesso: {len(dados_n)} linhas lidas para comparação.")
            
            total_itens = len(self.dados_agrupados)
            self.progress["maximum"] = total_itens
            decisao_global = None

            for i, (_, r) in enumerate(self.dados_agrupados.iterrows()):
                self.progress["value"] = i + 1
                self.log(f"Processando item {i+1}/{total_itens}: {r['Item']}")
                self.root.update() # Força a atualização do log e da barra

                idx_n = -1
                # Compara Data(0), MI(1), Hist(2) e Cód(3)
                for idx, ln in enumerate(dados_n):
                    if len(ln) > 3 and ln[0] == data_mov and ln[1] == mi and ln[2] == hist and ln[3] == r['Item']:
                        idx_n = idx + 1
                        qtd_atual = int(ln[7]) if len(ln) > 7 and str(ln[7]).isdigit() else 0
                        break

                if idx_n != -1:
                    self.log(f"⚠️ Conflito: Item {r['Item']} já existe na planilha.")
                    if decisao_global is None:
                        self.log("Aguardando decisão do usuário no pop-up...")
                        d = JanelaConflito(self.root, r['Item'], r['Qtd'])
                        self.root.wait_window(d)
                        res = d.resultado
                        if d.aplicar_todos: 
                            decisao_global = res
                            self.log(f"Decisão '{res}' aplicada para todos os próximos conflitos.")
                    else: 
                        res = decisao_global
                    
                    if res == "somar":
                        self.log(f"➕ Somando {r['Qtd']} ao valor atual ({qtd_atual}).")
                        nova_q = qtd_atual + int(r['Qtd'])
                        sheet.update_cell(idx_n, 8, nova_q)
                    else:
                        self.log(f"⏩ Item {r['Item']} ignorado por escolha do usuário.")
                else:
                    self.log(f"📤 Enviando novo registro: {r['Item']} | Qtd: {r['Qtd']}")
                    sheet.append_row([data_mov, mi, hist, r['Item'], r['Almox'], "", "", r['Qtd'], "", r['Desc']], value_input_option='USER_ENTERED')
            
            self.log("✅ SINCRONIZAÇÃO COM A NUVEM CONCLUÍDA.")
            self.lbl_status.config(text="Nuvem Atualizada!", fg="green")
            self.progress["value"] = 0
            
        except Exception as e:
            self.log(f"❌ ERRO NA NUVEM: {str(e)}")
            messagebox.showerror("Erro Nuvem", str(e))



class JanelaConflito(tk.Toplevel):
    def __init__(self, parent, item_nome, qtd_nova):
        super().__init__(parent)
        self.title("Conflito de Dados Detectado")
        self.geometry("450x250")
        self.resultado = None # 'somar', 'ignorar'
        self.aplicar_todos = False
        self.attributes("-topmost", True)
        self.grab_set()

        msg = f"O item '{item_nome}' já existe nesta data/MI.\nO que deseja fazer com a nova quantidade ({qtd_nova})?"
        tk.Label(self, text=msg, wraplength=400, pady=20, font=("Arial", 10)).pack()

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=10)

        tk.Button(btn_frame, text="Somar Quantidade", width=20, bg="#4CAF50", fg="white", 
                  command=lambda: self.finalizar("somar")).pack(side="left", padx=5)
        tk.Button(btn_frame, text="Ignorar (Pular)", width=20, bg="#f44336", fg="white", 
                  command=lambda: self.finalizar("ignorar")).pack(side="left", padx=5)

        self.var_todos = tk.BooleanVar()
        tk.Checkbutton(self, text="Fazer isso para todos os conflitos atuais", variable=self.var_todos).pack(pady=10)

    def finalizar(self, escolha):
        self.resultado = escolha
        self.aplicar_todos = self.var_todos.get()
        self.destroy()

if __name__ == "__main__":
    root = tk.Tk(); app = OctalinkApp(root); root.mainloop()