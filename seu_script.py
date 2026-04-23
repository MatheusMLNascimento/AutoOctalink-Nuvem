import sys
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
import json
from openpyxl.styles import Alignment

# Evita erro de console no Windows quando compilado com --noconsole
if sys.executable.endswith("pythonw.exe"):
    sys.stdout = open(os.devnull, "w")
    sys.stderr = open(os.devnull, "w")

CONFIG_FILE = "config_octalink.json"

def carregar_config():
    default = {"historicos": ["SAIDA USO ECO"], "codigos_mi": ["S500"]}
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
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
        self.root.title("Octalink Automator v2.0")
        self.root.geometry("650x650")
        self.config = carregar_config()
        self.dados_agrupados = None
        self.setup_ui()

    def setup_ui(self):
        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=10)

        tk.Label(self.main_frame, text="Automação de Movimentos Octalink", font=("Arial", 12, "bold")).pack(pady=10)

        # --- SEÇÃO HISTÓRICO ---
        tk.Label(self.main_frame, text="Histórico:").pack()
        self.cb_hist = ttk.Combobox(self.main_frame, values=self.config["historicos"] + ["Outro..."], width=50, state="readonly")
        self.cb_hist.pack(pady=2)
        self.cb_hist.bind("<<ComboboxSelected>>", lambda e: self.toggle_outro(self.cb_hist, self.ent_outro_hist))
        self.cb_hist.current(0)

        self.ent_outro_hist = tk.Entry(self.main_frame, width=53, font=("Arial", 9, "italic"), fg="grey")
        self.ent_outro_hist.insert(0, "Digite o novo histórico aqui...")
        self.ent_outro_hist.bind("<FocusIn>", lambda e: self.clear_placeholder(self.ent_outro_hist))

        # --- SEÇÃO CÓDIGO MI ---
        tk.Label(self.main_frame, text="Código MI:", font=("Arial", 10)).pack(pady=(10, 0))
        self.cb_mi = ttk.Combobox(self.main_frame, values=self.config["codigos_mi"] + ["Outro..."], width=50, state="readonly")
        self.cb_mi.pack(pady=2)
        self.cb_mi.bind("<<ComboboxSelected>>", lambda e: self.toggle_outro(self.cb_mi, self.ent_outro_mi))
        self.cb_mi.current(0)

        self.ent_outro_mi = tk.Entry(self.main_frame, width=53, font=("Arial", 9, "italic"), fg="grey")
        self.ent_outro_mi.insert(0, "Digite o novo Código MI aqui...")
        self.ent_outro_mi.bind("<FocusIn>", lambda e: self.clear_placeholder(self.ent_outro_mi))

        # --- BOTÃO SELECIONAR ---
        self.btn_load = tk.Button(self.main_frame, text="📂 Selecionar Arquivo TXT", bg="#2196F3", fg="white", 
                                  font=("Arial", 10, "bold"), command=self.ler_txt, height=2, width=30)
        self.btn_load.pack(pady=20)

        # --- PREVIEW ---
        tk.Label(self.main_frame, text="Preview da Importação:").pack()
        self.txt_preview = tk.Text(self.main_frame, height=15, width=85, state="disabled", font=("Consolas", 8), bg="#f4f4f4")
        self.txt_preview.pack(pady=5)
        
        self.txt_preview.tag_configure("erro", foreground="red", font=("Consolas", 9, "bold"))
        self.txt_preview.tag_configure("header", foreground="blue", font=("Consolas", 8, "bold"))

        # --- BOTÃO GERAR ---
        self.btn_gerar = tk.Button(self.main_frame, text="🚀 Gerar Planilha .xlsx", bg="#4CAF50", fg="white", 
                                   font=("Arial", 11, "bold"), state="disabled", command=self.exportar_excel, height=2, width=30)
        self.btn_gerar.pack(pady=10)

    def toggle_outro(self, combo, entry):
        if combo.get() == "Outro...":
            entry.pack(pady=2, after=combo)
        else:
            entry.pack_forget()

    def clear_placeholder(self, entry):
        if "Digite" in entry.get():
            entry.delete(0, tk.END)
            entry.config(fg="black", font=("Arial", 9, "normal"))

    def ler_txt(self):
        try:
            path = filedialog.askopenfilename(filetypes=[("Arquivos de Texto", "*.txt")])
            if not path: return

            with open(path, 'r', encoding='utf-8') as f:
                linhas = [l.strip() for l in f.readlines()]

            raw_data = []
            i = 0
            self.txt_preview.config(state="normal")
            self.txt_preview.delete(1.0, tk.END)

            while i < len(linhas):
                if not linhas[i]:
                    i += 1; continue
                try:
                    num_linha = i + 1
                    partes = linhas[i].upper().split(' - ')
                    if len(partes) < 2: raise ValueError(f"Linha {num_linha}: Formato inválido.")
                    cod, desc = partes[0].strip(), partes[1].strip()

                    if i + 1 >= len(linhas): raise ValueError(f"Linha {num_linha}: Dados incompletos.")
                    partes_alm = linhas[i+1].upper().split(' - ')
                    sigla_almox = partes_alm[0].strip()

                    idx_qtd = i + 4 if (i+3 < len(linhas) and "/" in linhas[i+3]) else i + 3
                    qtd = int(linhas[idx_qtd])

                    raw_data.append({"Item": cod, "Desc": desc, "Almox": sigla_almox, "Qtd": qtd})
                    i = idx_qtd + 1
                except Exception as e:
                    self.txt_preview.insert(tk.END, f"❌ ERRO NA LINHA {i+1}:\n{str(e)}\n", "erro")
                    self.btn_gerar.config(state="disabled")
                    messagebox.showerror("Erro no TXT", f"Erro na linha {i+1}")
                    return

            self.dados_agrupados = pd.DataFrame(raw_data).groupby(['Item', 'Desc', 'Almox']).agg({'Qtd': 'sum', 'Item': 'count'}).rename(columns={'Item': 'Pacotes'}).reset_index()
            
            header = f"{'PRODUTO':<25} | {'CÓD':<10} | {'ALM':<5} | {'QTD':<5} | {'PKTS'}\n"
            self.txt_preview.insert(tk.END, header, "header")
            self.txt_preview.insert(tk.END, ("-"*75) + "\n")
            for _, row in self.dados_agrupados.iterrows():
                nome = (row['Desc'][:23] + "..") if len(row['Desc']) > 23 else row['Desc']
                self.txt_preview.insert(tk.END, f"{nome:<25} | {row['Item']:<10} | {row['Almox']:<5} | {row['Qtd']:<5} | {row['Pacotes']}\n")
            
            self.txt_preview.config(state="disabled")
            self.btn_gerar.config(state="normal")

        except Exception as e:
            messagebox.showerror("Erro", str(e))

    def exportar_excel(self):
        try:
            h = self.ent_outro_hist.get().upper().strip() if self.cb_hist.get() == "Outro..." else self.cb_hist.get()
            m = self.ent_outro_mi.get().upper().strip() if self.cb_mi.get() == "Outro..." else self.cb_mi.get()

            if not h or "Digite" in h or not m or "Digite" in m:
                messagebox.showwarning("Atenção", "Preencha os campos corretamente.")
                return

            # Limpa nome do arquivo
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
                for cell in ws['A']: 
                    cell.number_format = 'DD/MM/YYYY'
                    cell.alignment = Alignment(horizontal='right')

            # SÓ SALVA NA CONFIG SE TUDO DEU CERTO ATÉ AQUI
            if self.cb_hist.get() == "Outro..." and h not in self.config["historicos"]:
                self.config["historicos"].append(h)
            if self.cb_mi.get() == "Outro..." and m not in self.config["codigos_mi"]:
                self.config["codigos_mi"].append(m)
            salvar_config(self.config)

            messagebox.showinfo("Sucesso", f"Gerado: {nome_arq}")
            
        except Exception as e:
            messagebox.showerror("Erro ao salvar", f"Feche o Excel ou verifique o nome.\nErro: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = OctalinkApp(root)
    root.mainloop()
