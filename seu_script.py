import subprocess
import sys

def instalar_dependencias():
    # Liste aqui o nome das bibliotecas que você usa (conforme aparecem no pip)
    pacotes = ["pandas", "openpyxl", "gspread"]
    
    for pacote in pacotes:
        try:
            # Tenta importar para ver se já existe
            __import__(pacote.replace('-', '_'))
        except ImportError:
            print(f"Instalando {pacote}...")
            # Chama o pip de forma silenciosa
            subprocess.check_call([sys.executable, "-m", "pip", "install", pacote])

# Executa a verificação antes de carregar o resto do código
instalar_dependencias()
import os
import json
import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
from openpyxl.styles import Alignment

# Tenta importar gspread
try:
    import gspread
except ImportError:
    gspread = None

# --- CONFIGURAÇÕES E BACKUP INTERNO ---
B_ID = "1n0sBeEQaBfe-Onh0GRNNX1ftJH35Fk0hGz1l_ZwQSBY"
B_ABA = "Abril"
B_JSON = """ {} """ # Cole seu JSON aqui entre as chaves se desejar

CONFIG_FILE = "config_octalink.json"

def carregar_config():
    default = {
        "historicos": ["SAIDA USO ECO"], "codigos_mi": ["S500"], "tutorial_ativo": True,
        "cloud": {"spreadsheet_id": B_ID, "worksheet_name": B_ABA, "creds_json": B_JSON.strip()}
    }
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                for k, v in default.items():
                    if k not in data: data[k] = v
                return data
        except: return default
    return default

def salvar_config(config):
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
    except: pass

class JanelaConflito(tk.Toplevel):
    def __init__(self, parent, item_nome, qtd_nova):
        super().__init__(parent)
        self.title("Conflito Detectado")
        self.geometry("450x250")
        self.resultado = 'ignorar'
        self.aplicar_todos = False
        self.attributes("-topmost", True)
        self.grab_set()
        tk.Label(self, text=f"O item '{item_nome}' já existe.\nSomar nova quantidade ({qtd_nova})?", 
                 wraplength=400, pady=20, font=("Arial", 10, "bold")).pack()
        f = tk.Frame(self); f.pack(pady=10)
        tk.Button(f, text="Somar", width=15, bg="green", fg="white", command=lambda: self.fim("somar")).pack(side="left", padx=5)
        tk.Button(f, text="Ignorar", width=15, bg="red", fg="white", command=lambda: self.fim("ignorar")).pack(side="left", padx=5)
        self.v = tk.BooleanVar(); tk.Checkbutton(self, text="Aplicar a todos os conflitos", variable=self.v).pack()
    def fim(self, r):
        self.resultado = r; self.aplicar_todos = self.v.get(); self.destroy()

class OctalinkApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Octalink Automator Pro v5.1")
        self.root.geometry("750x950")
        self.config = carregar_config()
        self.dados_agrupados = None
        self.caminho_txt = ""
        self.terminal_visivel = False
        self.setup_ui()
        if self.config.get("tutorial_ativo"): self.exibir_tutorial()

    def setup_ui(self):
        # Container Horizontal Principal
        self.pane = tk.Frame(self.root)
        self.pane.pack(fill="both", expand=True)

        # LADO ESQUERDO: Conteúdo Principal (Sempre visível)
        self.left_frame = tk.Frame(self.pane, padx=20, pady=10)
        self.left_frame.pack(side="left", fill="both", expand=True)

        # --- SEÇÃO NUVEM ---
        cloud_f = tk.LabelFrame(self.left_frame, text="⚙️ CONFIGURAÇÃO DA NUVEM", fg="red", font=("Arial", 9, "bold"))
        cloud_f.pack(fill="x", pady=5)
        tk.Button(cloud_f, text="📂 1. CARREGAR ARQUIVO JSON", command=self.importar_json_file, bg="#607D8B", fg="white").grid(row=0, columnspan=2, sticky="ew", pady=5, padx=5)
        self.ent_sheet_id = tk.Entry(cloud_f, width=40); self.ent_sheet_id.insert(0, self.config["cloud"]["spreadsheet_id"])
        self.ent_sheet_id.grid(row=1, column=1, pady=2, padx=5)
        tk.Label(cloud_f, text="2. Link/ID Planilha:").grid(row=1, column=0, padx=5, sticky="w")
        self.ent_sheet_name = tk.Entry(cloud_f, width=40); self.ent_sheet_name.insert(0, self.config["cloud"]["worksheet_name"])
        self.ent_sheet_name.grid(row=2, column=1, pady=2, padx=5)
        tk.Label(cloud_f, text="3. Nome da Aba:").grid(row=2, column=0, padx=5, sticky="w")
        self.btn_test_cloud = tk.Button(cloud_f, text="✅ VALIDAR CONEXÃO", bg="#2196F3", fg="white", command=self.testar_conexao)
        self.btn_test_cloud.grid(row=3, column=1, sticky="e", padx=5, pady=5)

        # --- OPERAÇÃO ---
        op_f = tk.LabelFrame(self.left_frame, text="🚀 OPERAÇÃO DIÁRIA"); op_f.pack(fill="x", pady=10)
        tk.Label(op_f, text="Histórico:").pack()
        self.cb_hist = ttk.Combobox(op_f, values=self.config["historicos"] + ["Outro..."], width=45, state="readonly")
        self.cb_hist.pack(pady=2); self.cb_hist.current(0)
        self.ent_outro_hist = tk.Entry(op_f, width=48); self.cb_hist.bind("<<ComboboxSelected>>", lambda e: self.toggle_outro(self.cb_hist, self.ent_outro_hist))
        self.cb_mi = ttk.Combobox(op_f, values=self.config["codigos_mi"] + ["Outro..."], width=45, state="readonly")
        self.cb_mi.pack(pady=2); self.cb_mi.current(0)
        self.ent_outro_mi = tk.Entry(op_f, width=48); self.cb_mi.bind("<<ComboboxSelected>>", lambda e: self.toggle_outro(self.cb_mi, self.ent_outro_mi))

        tk.Button(self.left_frame, text="📁 SELECIONAR ARQUIVO TXT", bg="#FF9800", fg="white", command=self.ler_txt, height=2).pack(fill="x", pady=10)
        self.txt_preview = tk.Text(self.left_frame, height=8, state="disabled", bg="#F5F5F5")
        self.txt_preview.pack(fill="x")
        self.txt_preview.tag_configure("header", foreground="blue", font=("Consolas", 8, "bold"))

        # Opções e Barra de Progresso
        opts_f = tk.Frame(self.left_frame); opts_f.pack(pady=5)
        self.check_local = tk.BooleanVar(value=True); tk.Checkbutton(opts_f, text="Gerar Excel", variable=self.check_local).pack(side="left")
        self.check_cloud = tk.BooleanVar(value=True); tk.Checkbutton(opts_f, text="Enviar Nuvem", variable=self.check_cloud).pack(side="left")
        self.check_data_nome = tk.BooleanVar(value=False); tk.Checkbutton(self.left_frame, text="Data pelo nome (DDMMYY ou DDMMYYYY)", variable=self.check_data_nome).pack()

        self.lbl_status = tk.Label(self.left_frame, text="Aguardando..."); self.lbl_status.pack()
        self.progress = ttk.Progressbar(self.left_frame, orient="horizontal", mode="determinate"); self.progress.pack(fill="x")
        self.btn_gerar = tk.Button(self.left_frame, text="🚀 FINALIZAR E GERAR TUDO", bg="#4CAF50", fg="white", state="disabled", command=self.fluxo_final, height=2)
        self.btn_gerar.pack(fill="x", pady=10)

        # Botão para expandir terminal
        self.btn_terminal = tk.Button(self.left_frame, text="📟 Mostrar Log Avançado ➡️", command=self.toggle_terminal, font=("Arial", 7))
        self.btn_terminal.pack(side="bottom", anchor="se")

        # 3. LADO DIREITO (Terminal / Log)
        # Ele é criado aqui, mas o .pack() só acontece no toggle_terminal
        self.right_frame = tk.Frame(self.pane, bg="black", padx=2, width=300)
        
        self.terminal_log = tk.Text(
            self.right_frame, 
            bg="black", 
            fg="#00FF00", # Verde Matrix
            font=("Consolas", 9), 
            padx=10, 
            pady=10, 
            state="disabled",
            wrap="word",
            insertbackground="white"
        )
        self.terminal_log.pack(fill="both", expand=True)

    def toggle_terminal(self):
        """Expandir ou Recolher o painel lateral de Log"""
        if not self.terminal_visivel:
            self.right_frame.pack(side="right", fill="both", expand=False)
            self.btn_terminal.config(text="⬅️ Ocultar Log")
            self.terminal_visivel = True
            self.root.geometry("1050x750") # Alarga a janela para mostrar o log
            self.log("Painel de log expandido.")
        else:
            self.right_frame.pack_forget()
            self.btn_terminal.config(text="📟 Mostrar Log Avançado ➡️")
            self.terminal_visivel = False
            self.root.geometry("650x750") # Retorna ao tamanho original


    def toggle_outro(self, combo, entry):
        if combo.get() == "Outro...": entry.pack(); entry.focus()
        else: entry.pack_forget()

    def extrair_id(self, texto):
        m = re.search(r"/d/([a-zA-Z0-9-_]+)", texto)
        return m.group(1) if m else texto.strip()

    def importar_json_file(self):
        p = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if p:
            with open(p, 'r') as f: d = json.load(f)
            self.config["cloud"]["creds_json"] = json.dumps(d)
            messagebox.showinfo("Sucesso", f"Compartilhe com: {d['client_email']}")

    def ler_txt(self):
        p = filedialog.askopenfilename(filetypes=[("TXT", "*.txt")])
        if not p: return
        self.caminho_txt = p; self.log(f"Lendo: {os.path.basename(p)}")
        try:
            with open(p, 'r', encoding='utf-8') as f:
                lns = [l.strip() for l in f.readlines() if l.strip()]
            raw = []
            i = 0
            while i < len(lns):
                p_i = lns[i].upper().split(' - '); p_a = lns[i+1].upper().split(' - ')
                idx_q = i+4 if (i+3 < len(lns) and "/" in lns[i+3]) else i+3
                raw.append({"Item": p_i[0].strip(), "Desc": p_i[1].strip() if len(p_i)>1 else "", "Almox": p_a[0].strip(), "Qtd": int(lns[idx_q])})
                i = idx_q + 1
            self.dados_agrupados = pd.DataFrame(raw).groupby(['Item', 'Desc', 'Almox']).agg({'Qtd':'sum'}).reset_index()
            self.txt_preview.config(state="normal"); self.txt_preview.delete(1.0, tk.END)
            self.txt_preview.insert(tk.END, f"{'PRODUTO':<30} | {'CÓD':<10} | {'QTD':<5}\n", "header")
            for _, r in self.dados_agrupados.iterrows():
                self.txt_preview.insert(tk.END, f"{r['Desc'][:28]:<30} | {r['Item']:<10} | {r['Qtd']}\n")
            self.btn_gerar.config(state="normal"); self.log("TXT OK.")
        except Exception as e: messagebox.showerror("Erro", str(e))
        self.txt_preview.config(state="disabled")

    def log(self, mensagem):
        """Escreve no terminal lateral com timestamp e auto-scroll"""
        tempo = datetime.now().strftime("%H:%M:%S")
        msg_formatada = f"[{tempo}] {mensagem}\n"
        
        # Garante que o widget possa ser editado
        self.terminal_log.config(state="normal")
        self.terminal_log.insert(tk.END, msg_formatada)
        
        # Auto-scroll para a última linha
        self.terminal_log.see(tk.END)
        
        # Bloqueia edição manual do usuário
        self.terminal_log.config(state="disabled")
        
        # Força o Tkinter a redesenhar a tela imediatamente
        self.root.update_idletasks()


    def fluxo_final(self):
        h = self.ent_outro_hist.get().upper() if self.cb_hist.get() == "Outro..." else self.cb_hist.get()
        m = self.ent_outro_mi.get().upper() if self.cb_mi.get() == "Outro..." else self.cb_mi.get()
        id_p = self.extrair_id(self.ent_sheet_id.get())
        
        # Lógica de Data Melhorada
        if self.check_data_nome.get():
            # Tenta encontrar 8 dígitos (DDMMYYYY) ou 6 dígitos (DDMMYY)
            ma8 = re.search(r"(\d{8})", os.path.basename(self.caminho_txt))
            ma6 = re.search(r"(\d{6})", os.path.basename(self.caminho_txt))
            
            if ma8:
                try:
                    data_f = datetime.strptime(ma8.group(1), "%d%m%Y").strftime("%d/%m/%Y")
                    self.log(f"Data extraída (8 dígitos): {data_f}")
                except: data_f = datetime.now().strftime("%d/%m/%Y")
            elif ma6:
                try:
                    data_f = datetime.strptime(ma6.group(1), "%d%m%y").strftime("%d/%m/%Y")
                    self.log(f"Data extraída (6 dígitos): {data_f}")
                except: data_f = datetime.now().strftime("%d/%m/%Y")
            else:
                data_f = datetime.now().strftime("%d/%m/%Y")
                self.log("Nenhum padrão de data encontrado no nome. Usando data de hoje.")
        else:
            data_f = datetime.now().strftime("%d/%m/%Y")

        if self.check_cloud.get(): self.enviar_para_nuvem(h, m, data_f, id_p)
        if self.check_local.get():
            nome = f"IMPORT_{m}_{h}_{datetime.now().strftime('%H%M%S')}.xlsx"
            pd.DataFrame([[data_f, m, h, r['Item'], r['Almox'], "", "", r['Qtd'], ""] for _, r in self.dados_agrupados.iterrows()], 
                         columns=['Data movimento', 'Cód. MI', 'Histórico', 'Cód. Item', 'Almoxarifado', 'Almoxarifado transf.', 'Unidade Medida', 'Qtde', 'Valor']).to_excel(nome, index=False)
            self.log(f"Excel salvo: {nome}")
        messagebox.showinfo("Sucesso", "Processamento concluído!")

    def enviar_para_nuvem(self, hist, mi, data_mov, id_p):
        try:
            self.log("Conectando..."); client = gspread.service_account_from_dict(json.loads(self.config["cloud"]["creds_json"]))
            sheet = client.open_by_key(id_p).worksheet(self.ent_sheet_name.get().strip())
            dados_n = sheet.get_all_values()
            self.progress["maximum"] = len(self.dados_agrupados); decisao = None
            for i, (_, r) in enumerate(self.dados_agrupados.iterrows()):
                self.progress["value"] = i + 1; self.log(f"Sinc: {r['Item']}"); self.root.update()
                idx_n = -1
                for ix, ln in enumerate(dados_n):
                    if len(ln)>3 and ln[0]==data_mov and ln[1]==mi and ln[2]==hist and ln[3]==r['Item']:
                        idx_n = ix + 1; qtd_at = int(ln[7]) if ln[7].isdigit() else 0; break
                if idx_n != -1:
                    if decisao is None:
                        d = JanelaConflito(self.root, r['Item'], r['Qtd']); self.root.wait_window(d)
                        res = d.resultado; 
                        if d.aplicar_todos: decisao = res
                    else: res = decisao
                    if res == "somar": sheet.update_cell(idx_n, 8, qtd_at + int(r['Qtd']))
                else: sheet.append_row([data_mov, mi, hist, r['Item'], r['Almox'], "", "", r['Qtd'], "", r['Desc']], value_input_option='USER_ENTERED')
            self.log("Nuvem OK!")
        except Exception as e: self.log(f"Erro Nuvem: {e}"); messagebox.showerror("Erro Nuvem", str(e))

    def testar_conexao(self):
        self.log("Testando conexão..."); 
        try:
            id_p = self.extrair_id(self.ent_sheet_id.get())
            c = json.loads(self.config["cloud"]["creds_json"])
            gspread.service_account_from_dict(c).open_by_key(id_p)
            self.btn_test_cloud.config(bg="#4CAF50", text="✅ CONECTADO"); self.log("Sucesso!")
        except Exception as e: self.btn_test_cloud.config(bg="#f44336"); messagebox.showerror("Erro", str(e))

    def exibir_tutorial(self):
        w = tk.Toplevel(self.root); w.attributes("-topmost", True); w.grab_set()
        tk.Label(w, text="Tutorial: 1. Carregue JSON 2. Compartilhe 3. Valide 4. Processe", padx=20, pady=20).pack()
        tk.Button(w, text="OK", command=lambda: [self.config.update({"tutorial_ativo":False}), salvar_config(self.config), w.destroy()]).pack()

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = OctalinkApp(root)
        root.mainloop()
    except Exception as e:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, f"Erro Fatal: {e}", "Erro Crítico", 0x10)
