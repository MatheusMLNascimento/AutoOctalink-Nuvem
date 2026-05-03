import os
import json
import re
import base64
import hashlib
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
from datetime import datetime
from io import BytesIO
from PIL import Image, ImageTk, ImageGrab

# --- CONFIGURAÇÕES MESTRAS (DNA DO APP) ---
# Se o arquivo sumir, o app usa esses padrões para se auto-recuperar.
CONFIG_FILE = "config_octalink.json"

def carregar_config():
    """DNA: Estrutura base de dados e preferências do sistema"""
    default = {
        "tutorial_ativo": True,
        "historicos": ["SAIDA USO ECO", "SAIDA USO CC", "SAIDA USO MI"],
        "codigos_mi": ["S500"],
        "coluna_qtd": 8,
        "cloud": {
            "spreadsheet_id": "",
            "worksheet_name": "Sheet1",
            "creds_json": "{}",
            "folder_backup_id": ""
        },
        "tutorial_dinamico": [], # Onde a 'Forma de Gelo' guarda as etapas
        "textos_ui": {
            "titulo": "Octalink Automator Pro v6.0",
            "msg_boas_vindas": "Bem-vindo! Carregue um TXT para começar."
        }
    }
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                # Garante que novas chaves existam se o usuário estiver vindo de uma versão antiga
                for k, v in default.items():
                    if k not in data: data[k] = v
                return data
        except: 
            return default
    return default

def salvar_config(config):
    """Persistência de dados segura"""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"Erro ao salvar configuração: {e}")
# --- LÓGICA DE INTEGRIDADE E SEGURANÇA (HASH MD5) ---

def calcular_hash_md5_binario(conteudo_binario):
    """
    Gera uma 'impressão digital' única para blocos de dados.
    Essencial para comparar o arquivo do PC com o da Nuvem.
    """
    return hashlib.md5(conteudo_binario).hexdigest()

def calcular_hash_arquivo_local(caminho_arquivo):
    """
    Lê o arquivo local em pedaços (chunks) para não travar a memória
    e gera o código Hash MD5 de integridade.
    """
    if not os.path.exists(caminho_arquivo):
        return None
    
    hasher = hashlib.md5()
    try:
        with open(caminho_arquivo, "rb") as f:
            # Lê o arquivo em blocos de 4KB para eficiência
            for chunk in iter(lambda: f.read(4096), b""):
                hasher.update(chunk)
        return hasher.hexdigest()
    except Exception:
        return None

def validar_arquivo_saudavel(caminho_local):
    """
    Melhoria de Robustez: Verifica se o arquivo não está vazio 
    ou corrompido antes de qualquer operação de envio.
    """
    if not os.path.exists(caminho_local):
        return False, "Arquivo não encontrado."
    
    tamanho = os.path.getsize(caminho_local)
    
    # Arquivos com menos de 5 bytes são considerados inválidos/vazios
    if tamanho < 5:
        return False, "O arquivo parece estar vazio ou corrompido."
    
    return True, "OK"

# --- AUXILIARES DE LIMPEZA DE DADOS ---

def extrair_id_limpo(texto):
    """
    Melhoria de UX: Se o usuário colar o link inteiro da planilha,
    o robô limpa e extrai apenas o ID necessário entre as barras.
    """
    match = re.search(r"/d/([a-zA-Z0-9-_]+)", texto)
    if match:
        return match.group(1)
    return texto.strip()

class Tooltip:
    """
    Cria balões explicativos que surgem ao passar o mouse.
    Corrigido para evitar que fiquem 'órfãos' na tela.
    """
    def __init__(self, widget, texto):
        self.widget = widget
        self.texto = texto
        self.tip_window = None
        self.id_espera = None
        
        # Vincula os eventos do mouse
        self.widget.bind("<Enter>", self.agendar_exibicao)
        self.widget.bind("<Leave>", self.esconder_tip)
        self.widget.bind("<ButtonPress>", self.esconder_tip)

    def agendar_exibicao(self, event=None):
        """ Aguarda 500ms antes de mostrar para não poluir a visão """
        self.cancelar_espera()
        if not self.texto:
            return
        self.id_espera = self.widget.after(500, self.mostrar_tip)

    def mostrar_tip(self):
        """ Cria a janelinha amarela de dica """
        if self.tip_window or not self.texto:
            return
            
        # Calcula a posição (logo abaixo do mouse)
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25
        
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True) # Remove bordas do Windows
        tw.wm_geometry(f"+{x}+{y}")
        tw.attributes("-topmost", True) # Sempre por cima
        
        # Design do Balão
        lbl = tk.Label(tw, text=self.texto, background="#FFFFE1", relief="solid", 
                       borderwidth=1, font=("Arial", 8, "normal"), padx=5, pady=2)
        lbl.pack()

    def esconder_tip(self, event=None):
        """ Remove o balão e cancela qualquer agendamento """
        self.cancelar_espera()
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None

    def cancelar_espera(self):
        """ Evita que a dica apareça depois que o mouse já saiu """
        if self.id_espera:
            self.widget.after_cancel(self.id_espera)
            self.id_espera = None

# --- FUNÇÃO AUXILIAR PARA CRIAR TOOLTIPS RAPIDAMENTE ---
def add_tip(widget, texto):
    return Tooltip(widget, texto)

class MascaraSelecao(tk.Toplevel):
    def __init__(self, parent, callback):
        super().__init__(parent)
        self.parent = parent 
        self.callback = callback
        
        # 1. Inicialização Única de Atributos
        self.start_x: int | None = None
        self.start_y: int | None = None
        self.rect: int | None = None
        
        # 2. Configurações de Janela
        self.attributes("-alpha", 0.5) 
        self.attributes("-topmost", True)
        self.overrideredirect(True) 
        
        # 3. Sincronização e Limites do App
        self.parent.update_idletasks()
        self.x_off = self.parent.winfo_rootx()
        self.y_off = self.parent.winfo_rooty()
        self.w_app = self.parent.winfo_width()
        self.h_app = self.parent.winfo_height()
        self.geometry(f"{self.w_app}x{self.h_app}+{self.x_off}+{self.y_off}")

        # 4. Interface e Binds
        self.canvas = tk.Canvas(self, cursor="cross", bg="black", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)

        self.canvas.bind("<ButtonPress-1>", self.ao_clicar)
        self.canvas.bind("<B1-Motion>", self.ao_arrastar)
        self.canvas.bind("<ButtonRelease-1>", self.ao_soltar)
        self.bind("<Escape>", lambda e: self.destroy()) 
        
        # 5. Blindagem de Foco
        self.focus_force()
        self.grab_set()


    def ao_clicar(self, event: tk.Event):
        """Inicia o retângulo garantindo tipos numéricos"""
        # Extrai como int para garantir que não seja None
        ex, ey = int(event.x), int(event.y)
        
        self.start_x, self.start_y = ex, ey
        
        # Agora o editor sabe que ex e ey são floats/ints válidos
        self.rect = self.canvas.create_rectangle(
            ex, ey, ex + 1, ey + 1, 
            outline="yellow", 
            width=2, 
            dash=(4, 4)
        )

    def ao_arrastar(self, event: tk.Event):
        """Atualiza o retângulo com travas e verificação de None"""
        # Type Guard: Só executa se todas as variáveis foram inicializadas
        if (self.rect is not None and 
            self.start_x is not None and 
            self.start_y is not None):
            
            # Trava dentro dos limites do App (0 até largura/altura)
            cur_x = max(0, min(int(event.x), self.w_app))
            cur_y = max(0, min(int(event.y), self.h_app))
            
            # Coords recebe (x0, y0, x1, y1)
            self.canvas.coords(self.rect, self.start_x, self.start_y, cur_x, cur_y)


    def ao_soltar(self, event):
        """Captura o print e finaliza com proteções de tipo"""
        # 1. TRAVA DE SEGURANÇA (Resolve o erro 'Unknown | None')
        # Se o usuário apenas clicar e soltar sem arrastar, ou se o clique falhou
        if self.start_x is None or self.start_y is None: 
            self.grab_release()
            self.destroy()
            return
        
        # 2. LIMITA AS COORDENADAS (Trava dentro do App)
        ex = max(0, min(event.x, self.w_app))
        ey = max(0, min(event.y, self.h_app))
        
        # 3. CÁLCULOS (Agora o editor sabe que são números válidos)
        sw = abs(ex - self.start_x)
        sh = abs(ey - self.start_y)
        sx = min(self.start_x, ex)
        sy = min(self.start_y, ey)

        # 4. PROCESSAMENTO DO PRINT
        if sw > 10 and sh > 10:
            # Coordenadas reais para o Print (relativas à tela)
            ax = self.winfo_rootx() + sx
            ay = self.winfo_rooty() + sy
            
            try:
                screenshot = ImageGrab.grab(bbox=(ax, ay, ax + sw, ay + sh))
                self.callback((sx, sy, sw, sh), screenshot)
            except Exception as e:
                if hasattr(self.parent, 'log'):
                    self.parent.log(f"Erro na captura: {e}")
        
        # 5. LIMPEZA FINAL
        self.grab_release()
        self.destroy()



class JanelaCompressaoTutorial(tk.Toplevel):
    """
    Interface para revisar o print, definir qualidade e 
    converter para texto (Base64) embutido no código.
    """
    def __init__(self, parent, imagem_pil, coords, callback_salvar):
        super().__init__(parent)
        self.title("Otimizar Destaque do Tutorial")
        self.geometry("400x550")
        self.imagem_pil = imagem_pil
        self.coords = coords
        self.callback_salvar = callback_salvar
        self.img_b64_final = ""
        
        self.setup_ui()
        self.atualizar_preview()

    def setup_ui(self):
        # Preview da Captura
        tk.Label(self, text="Visualização do Destaque:", font=("Arial", 9, "bold")).pack(pady=10)
        self.lbl_img = tk.Label(self, bg="#dcdcdc", relief="sunken")
        self.lbl_img.pack(pady=5)

        # Controle de Qualidade (Melhoria 15)
        tk.Label(self, text="Qualidade da Compressão (%):").pack(pady=(10, 0))
        self.slider = tk.Scale(self, from_=10, to=100, orient="horizontal", command=lambda x: self.atualizar_preview())
        self.slider.set(60) # Padrão 60% para bom equilíbrio entre peso e nitidez
        self.slider.pack(fill="x", padx=40)

        self.lbl_peso = tk.Label(self, text="Tamanho estimado: 0 KB", fg="#1a73e8")
        self.lbl_peso.pack()

        # Input do Texto do Balão
        tk.Label(self, text="Instrução para o Usuário:", font=("Arial", 9, "bold")).pack(pady=(15, 0))
        self.ent_texto = tk.Entry(self, width=45)
        self.ent_texto.pack(pady=5, padx=20)
        self.ent_texto.insert(0, "Clique aqui para...")

        # Botão Finalizar
        tk.Button(self, text="💾 SALVAR ETAPA NO TUTORIAL", bg="#1a73e8", fg="white", 
                  font=("Arial", 10, "bold"), command=self.concluir).pack(pady=20, fill="x", padx=40)

    def atualizar_preview(self):
        """ Simula a compressão JPEG e calcula o peso em Base64 """
        qualidade = self.slider.get()
        
        # Converte para JPEG em memória para comprimir
        buffer = BytesIO()
        # Necessário converter para RGB se for RGBA (PNG)
        img_rgb = self.imagem_pil.convert("RGB")
        img_rgb.save(buffer, format="JPEG", quality=qualidade)
        
        # Calcula peso
        dados_binarios = buffer.getvalue()
        peso_kb = len(dados_binarios) / 1024
        
        # Gera Base64
        self.img_b64_final = base64.b64encode(dados_binarios).decode('utf-8')
        
        # Atualiza Preview Visual
        img_preview = Image.open(buffer)
        img_preview.thumbnail((300, 200))
        photo = ImageTk.PhotoImage(img_preview)
        
        # 1. Aplica a imagem ao widget
        self.lbl_img.config(image=photo)
        
        # 2. Salva a referência para a imagem não sumir (Silenciando o erro do editor)
        setattr(self.lbl_img, "image", photo) 
        
        self.lbl_peso.config(text=f"Peso no Arquivo: {peso_kb:.1f} KB")

    def concluir(self):
        """ Valida e envia os dados para o banco de dados do app """
        texto = self.ent_texto.get().strip()
        if not texto:
            messagebox.showwarning("Atenção", "O passo do tutorial precisa de um texto explicativo!")
            return
            
        dados_finais = {
            "texto": texto,
            "coords": self.coords,
            "img_b64": self.img_b64_final
        }
        self.callback_salvar(dados_finais)
        self.destroy()

class HolofoteTutorial(tk.Toplevel):
    """
    Cria uma máscara escura sobre todo o app e 'recorta' uma área
    clara para destacar um botão ou campo específico.
    """
    def __init__(self, parent, coords, texto, img_b64=None):
        super().__init__(parent)
        self.parent = parent
        self.coords = coords # (x, y, largura, altura)
        self.texto = texto
        self.img_b64 = img_b64
        
        # Configuração da Janela de Máscara
        self.overrideredirect(True)
        self.attributes("-alpha", 0.75) # Nível de escuridão (75%)
        self.attributes("-topmost", True)
        
        # Sincroniza posição com o App
        x_app = self.parent.winfo_rootx()
        y_app = self.parent.winfo_rooty()
        w_app = self.parent.winfo_width()
        h_app = self.parent.winfo_height()
        self.geometry(f"{w_app}x{h_app}+{x_app}+{y_app}")
        
        # Canvas para desenhar a escuridão e o 'furo'
        self.canvas = tk.Canvas(self, bg="black", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)
        
        # Janela de Informação (O Balão de Texto)
        self.janela_info = None
        
        self.desenhar_holofote()
        self.exibir_balao_informativo()
        
        # Clique na máscara fecha o tutorial
        self.canvas.bind("<Button-1>", lambda e: self.encerrar())

    def desenhar_holofote(self):
        """ Desenha retângulos pretos ao redor da área clara para simular o foco """
        self.canvas.delete("all")
        
        x, y, w, h = self.coords
        w_total = self.parent.winfo_width()
        h_total = self.parent.winfo_height()
        
        # Retângulos que formam a máscara (Cima, Baixo, Esquerda, Direita)
        self.canvas.create_rectangle(0, 0, w_total, y, fill="black", outline="") # Topo
        self.canvas.create_rectangle(0, y, x, y + h, fill="black", outline="") # Esquerda
        self.canvas.create_rectangle(x + w, y, w_total, y + h, fill="black", outline="") # Direita
        self.canvas.create_rectangle(0, y + h, w_total, h_total, fill="black", outline="") # Base
        
        # Borda de destaque no buraco (Amarelo Neon para chamar atenção)
        self.canvas.create_rectangle(x, y, x + w, y + h, outline="#ccff00", width=3)

    def exibir_balao_informativo(self):
        """ Cria um balão flutuante próximo ao holofote com o texto e imagem """
        x, y, w, h = self.coords
        
        self.janela_info = tk.Toplevel(self)
        self.janela_info.overrideredirect(True)
        self.janela_info.attributes("-topmost", True)
        self.janela_info.config(bg="white", padx=15, pady=15, relief="solid", borderwidth=1)
        
        # Tenta posicionar o balão abaixo do holofote, se não couber, coloca acima
        pos_y = self.parent.winfo_rooty() + y + h + 10
        if pos_y + 200 > self.winfo_screenheight():
            pos_y = self.parent.winfo_rooty() + y - 210
            
        self.janela_info.geometry(f"+{self.parent.winfo_rootx() + x}+{pos_y}")
        
        # Texto da Instrução
        tk.Label(self.janela_info, text=self.texto, wraplength=250, bg="white", 
                 font=("Arial", 10, "bold"), justify="left").pack()
        
        # Imagem de Auxílio (se existir no Base64)
        if self.img_b64:
            try:
                dados_img = base64.b64decode(self.img_b64)
                img = Image.open(BytesIO(dados_img))
                img.thumbnail((200, 150))
                self.photo = ImageTk.PhotoImage(img)
                tk.Label(self.janela_info, image=self.photo, bg="white").pack(pady=5)
            except: pass
            
        tk.Label(self.janela_info, text="[ Clique na tela para fechar ]", 
                 font=("Arial", 7, "italic"), bg="white", fg="grey").pack(pady=(10,0))

    def encerrar(self):
        """ Fecha o holofote e o balão de informação """
        if self.janela_info:
            self.janela_info.destroy()
        self.destroy()

class GradeFormaDeGelo(tk.Frame):
    """
    Painel que organiza as etapas do tutorial em uma grade flexível.
    Permite visualizar, excluir e testar cada 'cubo' (etapa).
    """
    def __init__(self, parent, config, callback_remover, callback_testar):
        super().__init__(parent, bg="#f5f5f5")
        self.config = config
        self.callback_remover = callback_remover
        self.callback_testar = callback_testar
        
        # Container com Scroll (Para caso de muitos passos)
        self.canvas = tk.Canvas(self, bg="#f5f5f5", highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scroll_frame = tk.Frame(self.canvas, bg="#f5f5f5")
        
        self.scroll_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        self.renderizar_grade()

    def renderizar_grade(self):
        """ Desenha os 'cubos de gelo' baseados no tutorial_dinamico """
        # Limpa a grade atual
        for widget in self.scroll_frame.winfo_children():
            widget.destroy()
            
        passos = self.config.get("tutorial_dinamico", [])
        
        if not passos:
            tk.Label(self.scroll_frame, text="Nenhuma etapa criada.\nUse o botão 'Adicionar' acima.", 
                     bg="#f5f5f5", fg="grey", font=("Arial", 10, "italic")).pack(pady=50)
            return

        colunas_max = 3
        for i, passo in enumerate(passos):
            linha = i // colunas_max
            coluna = i % colunas_max
            
            # O 'Cubo de Gelo' (Card)
            card = tk.Frame(self.scroll_frame, bg="white", bd=1, relief="solid", padx=5, pady=5)
            card.grid(row=linha, column=coluna, padx=10, pady=10, sticky="nsew")
            
            # Cabeçalho do Card
            tk.Label(card, text=f"Etapa {i+1}", font=("Arial", 8, "bold"), bg="white", fg="#1a73e8").pack()
            
            # Miniatura da Imagem (Base64)
            if passo.get("img_b64"):
                try:
                    dados_img = base64.b64decode(passo["img_b64"])
                    img_pil = Image.open(BytesIO(dados_img))
                    img_pil.thumbnail((120, 80))
                    photo = ImageTk.PhotoImage(img_pil)
                    
                    lbl_img = tk.Label(card, image=photo, bg="white")
        
                    setattr(lbl_img, "image", photo) 
                    
                    lbl_img.pack(pady=2)
                except Exception:
                    tk.Label(card, text="[Erro Imagem]", bg="white", fg="red", font=("Arial", 7)).pack()

            
            # Texto Resumido
            resumo = (passo["texto"][:35] + '..') if len(passo["texto"]) > 37 else passo["texto"]
            tk.Label(card, text=resumo, font=("Arial", 7), bg="white", wraplength=130).pack()
            
            # Botões de Ação
            btn_f = tk.Frame(card, bg="white")
            btn_f.pack(fill="x", side="bottom", pady=2)
            
            # Botão Testar (Olho)
            btn_test = tk.Button(btn_f, text="👁️", font=("Arial", 8), bg="#e8f0fe", relief="flat",
                                 command=lambda p=passo: self.callback_testar(p))
            btn_test.pack(side="left", expand=True, fill="x", padx=1)
            add_tip(btn_test, "Testar como o usuário verá este passo")
            
            # Botão Remover (Lixeira)
            btn_del = tk.Button(btn_f, text="🗑️", font=("Arial", 8), bg="#fce8e6", relief="flat",
                                command=lambda idx=i: self.callback_remover(idx))
            btn_del.pack(side="left", expand=True, fill="x", padx=1)
            add_tip(btn_del, "Excluir esta etapa definitivamente")

class OctalinkApp:

    def verificar_dependencias_criticas(self):
        """ 
        Verifica se as bibliotecas pesadas estão carregadas.
        """
        modulos_faltantes = []
        try: import pandas
        except ImportError: modulos_faltantes.append("pandas")
        try: import PIL
        except ImportError: modulos_faltantes.append("pillow")
        try: import gspread
        except ImportError: modulos_faltantes.append("gspread")
        
        if modulos_faltantes:
            self.log(f"⚠️ Alerta: Módulos não encontrados: {', '.join(modulos_faltantes)}")
            return False
        return True

    def verificar_tutorial_inicial(self):
        """ 
        Verifica se deve rodar o tutorial de boas-vindas.
        Chamado na inicialização se 'tutorial_ativo' for True no JSON.
        """
        if not self.config.get("tutorial_dinamico"):
            self.log("Nenhum tutorial dinâmico encontrado para exibir.")
            return
            
        self.log("🎬 Iniciando tutorial interativo...")
        self.executar_sequencia_tutorial(0)

    def executar_sequencia_tutorial(self, indice):
        """ 
        Percorre as etapas do tutorial uma a uma usando o Holofote.
        """
        passos = self.config.get("tutorial_dinamico", [])
        
        if indice >= len(passos):
            # Fim do Tutorial: Desativa para a próxima abertura
            self.config["tutorial_ativo"] = False
            salvar_config(self.config)
            self.log("🏁 Tutorial concluído com sucesso.")
            messagebox.showinfo("Tutorial", "Guia finalizado!\n\nAgora você está pronto para operar.")
            return

        passo = passos[indice]
        
        # Aciona o Holofote (A classe HolofoteTutorial deve estar na Parte 6)
        guia = HolofoteTutorial(
            self.root, 
            passo["coords"], 
            passo["texto"], 
            passo["img_b64"]
        )
        
        # Espera o usuário fechar o destaque atual para chamar o próximo
        self.root.wait_window(guia)
        self.root.after(300, lambda: self.executar_sequencia_tutorial(indice + 1))

    def __init__(self, root):
        self.root = root
        self.config = carregar_config()
        self.dados_agrupados = None
        self.caminho_txt = ""
        self.terminal_visivel = False
        
        # Variáveis de Controle (Liberdade total do usuário - v6.0)
        self.var_local = tk.BooleanVar(value=True)
        self.var_cloud = tk.BooleanVar(value=True)
        self.var_data_nome = tk.BooleanVar(value=False)
        
        # Configuração da Janela Principal
        self.root.title(self.config["textos_ui"]["titulo"])
        self.root.geometry("900x800")
        
        # Melhora a nitidez em telas High DPI (Windows 10/11)
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass # Caso não esteja no Windows ou falte permissão

        self.setup_ui()
        
        # Inicia tutorial automático se for o primeiro acesso
        if self.config.get("tutorial_ativo"):
            self.root.after(1000, self.verificar_tutorial_inicial)

    # --- FUNÇÕES DE SUPORTE AO SCROLL (MÓDULO UI) ---

    def _on_frame_configure(self, event):
        """ 
        Atualiza a região de scroll sempre que o conteúdo muda de tamanho.
        Resolve o erro 'Attribute unknown'.
        """
        self.canvas_scroll.configure(scrollregion=self.canvas_scroll.bbox("all"))

    def _on_canvas_configure(self, event):
        """ 
        Faz com que a largura do conteúdo interno acompanhe a largura da janela.
        """
        canvas_width = event.width
        self.canvas_scroll.itemconfig(self.scroll_window, width=canvas_width)

    def _on_mousewheel(self, event):
        """ 
        Habilita a rodinha do mouse para rolar a tela principal.
        """
        # No Windows, event.delta é 120 ou -120
        self.canvas_scroll.yview_scroll(int(-1 * (event.delta / 120)), "units")


    def setup_ui(self):
        """ Interface Principal com Tela Rolável (v6.0) """
        self.pane_principal = tk.Frame(self.root)
        self.pane_principal.pack(fill="both", expand=True)

        # --- LADO ESQUERDO: ÁREA COM SCROLL ---
        self.left_container = tk.Frame(self.pane_principal, padx=5, pady=5)
        self.left_container.pack(side="left", fill="both", expand=True)

        # Criando o Canvas e a Scrollbar
        self.canvas_scroll = tk.Canvas(self.left_container, highlightthickness=0)
        self.scrollbar_v = ttk.Scrollbar(self.left_container, orient="vertical", command=self.canvas_scroll.yview)
        
        # Este é o frame onde REALMENTE colocaremos os widgets
        self.left_frame = tk.Frame(self.canvas_scroll, padx=20, pady=15)
        
        # Configuração de ancoragem do Scroll
        self.scroll_window = self.canvas_scroll.create_window((0, 0), window=self.left_frame, anchor="nw")
        self.canvas_scroll.configure(yscrollcommand=self.scrollbar_v.set)

        # Empacotamento do sistema de Scroll
        self.canvas_scroll.pack(side="left", fill="both", expand=True)
        self.scrollbar_v.pack(side="right", fill="y")

        # Eventos para ajustar o tamanho do scroll
        self.left_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas_scroll.bind("<Configure>", self._on_canvas_configure)
        
        # Habilitar scroll com a bolinha do mouse
        self.canvas_scroll.bind_all("<MouseWheel>", self._on_mousewheel)

        # ---------------------------------------------------------
        # A PARTIR DAQUI, CONTINUE COLANDO SEUS WIDGETS NO self.left_frame
        # ---------------------------------------------------------

        # 1. CABEÇALHO (Aqui criamos o lbl_titulo)
        self.lbl_titulo = tk.Label(
            self.left_frame, 
            text="Estoque Nuvem Octa", 
            font=("Arial", 12, "bold"), 
            fg="#2c3e50"
        )
        self.lbl_titulo.pack(pady=5)

        # Botão de Engrenagem (Posicionamento absoluto)
        self.btn_config = tk.Button(self.left_frame, text="⚙️", font=("Arial", 16), 
                                     command=self.abrir_painel_config, relief="flat", cursor="hand2")
        self.btn_config.place(relx=1.0, x=0, y=-5, anchor="ne")

        # 2. IDENTIFICAÇÃO DA CARGA (Chama Parte 10)
        self.setup_operacao_ui()

        # 3. TABELA DE PREVIEW (Chama Parte 11)
        self.setup_tabela_ui()

        # 4. OPÇÕES DE SAÍDA (Aqui residem os seletores de liberdade do usuário)
        out_f = tk.LabelFrame(self.left_frame, text="⚙️ OPÇÕES DE SAÍDA", padx=10, pady=10)
        out_f.pack(fill="x", pady=10)

        tk.Checkbutton(out_f, text="Gerar Excel Local", variable=self.var_local).grid(row=0, column=0, sticky="w")
        tk.Checkbutton(out_f, text="Enviar para Nuvem", variable=self.var_cloud).grid(row=0, column=1, sticky="w")
        tk.Checkbutton(out_f, text="Data pelo Nome (ddmmyy ou mmddyyyy)", variable=self.var_data_nome).grid(row=1, column=0, sticky="w")

        # 5. BOTÕES DE AÇÃO (Criando atributos críticos)
        self.btn_carregar = tk.Button(self.left_frame, text="📁 1. CARREGAR RELATÓRIO TXT", bg="#FF9800", fg="white", 
                                       font=("Arial", 10, "bold"), height=2, command=self.ler_txt)
        self.btn_carregar.pack(fill="x", pady=5)

        self.btn_processar = tk.Button(self.left_frame, text="🚀 2. FINALIZAR E ENVIAR TUDO", bg="#4CAF50", fg="white", 
                                        font=("Arial", 11, "bold"), height=2, state="disabled", command=self.fluxo_final)
        self.btn_processar.pack(fill="x", pady=5)

        # 6. STATUS E LOG
        self.bar_progresso = ttk.Progressbar(self.left_frame, orient="horizontal", mode="determinate")
        self.bar_progresso.pack(fill="x", pady=10)
        
        self.adicionar_botao_log_principal() # Parte 9

        # --- REATIVANDO TOOLTIPS ---
        add_tip(self.lbl_titulo, "Versão 6.0 Pro - Máxima Resiliência")
        add_tip(self.btn_config, "Configurações, Drive e Criador de Tutorial")
        add_tip(self.btn_carregar, "Selecione o arquivo exportado pelo sistema (TXT)")
        add_tip(self.btn_processar, "Inicia a soma dos itens, gera o Excel e sincroniza com o Google")
        add_tip(self.btn_toggle_log, "Veja os detalhes técnicos de cada ação do robô")

        # --- LADO DIREITO: TERMINAL ---
        self.right_frame = tk.Frame(self.pane_principal, bg="#1e1e1e", width=320)

    def iniciar_criador_tutorial(self):
        """ 
        Melhoria 5: Ativa a Máscara para selecionar área e criar novo passo.
        """
        def callback_criacao(coords, img_pil):
            # Chama a janela de compressão da Parte 5
            JanelaCompressaoTutorial(self.root, img_pil, coords, self.salvar_passo_tutorial)
            
        # Ativa o seletor estilo Windows (MascaraSelecao deve estar definida no arquivo)
        MascaraSelecao(self.root, callback_criacao)

    def salvar_passo_tutorial(self, dados_passo):
        """ Registra o novo passo no DNA do app e atualiza a grade visual """
        if "tutorial_dinamico" not in self.config:
            self.config["tutorial_dinamico"] = []
        
        self.config["tutorial_dinamico"].append(dados_passo)
        salvar_config(self.config)
        self.log("✅ Nova etapa adicionada ao tutorial.")
        self.renderizar_grade_tutorial()

    def renderizar_grade_tutorial(self):
        """ 
        Melhoria 7: Organiza as etapas em blocos visuais na Engrenagem.
        """
        # Proteção: Verifica se o frame_grid existe (se a Engrenagem está aberta)
        grid = getattr(self, "frame_grid", None)
        if grid is None:
            return

        # Limpa a grade atual
        for child in grid.winfo_children():
            child.destroy()
            
        passos = self.config.get("tutorial_dinamico", [])
        
        if not passos:
            tk.Label(grid, text="Nenhuma etapa criada.\nUse o botão acima.", 
                     bg="#f5f5f5", fg="grey", font=("Arial", 9, "italic")).pack(pady=40)
            return

        # Monta a grade (3 colunas)
        for i, passo in enumerate(passos):
            linha = i // 3
            coluna = i % 3
            
            card = tk.Frame(grid, bd=1, relief="solid", width=170, height=210, padx=5, pady=5, bg="white")
            card.grid(row=linha, column=coluna, padx=8, pady=8)
            card.grid_propagate(False)
            
            tk.Label(card, text=f"ETAPA {i+1}", font=("Arial", 8, "bold"), bg="white", fg="#1a73e8").pack()
            
            if passo.get("img_b64"):
                try:
                    img_dados = base64.b64decode(passo["img_b64"])
                    img_pil = Image.open(BytesIO(img_dados)).resize((140, 90))
                    foto = ImageTk.PhotoImage(img_pil)
                    lbl_img = tk.Label(card, image=foto, bg="white")
                    setattr(lbl_img, "image", foto) 
                    lbl_img.pack(pady=2)
                except Exception:
                    tk.Label(card, text="[Erro Imagem]", bg="white", font=("Arial", 7)).pack()
            
            txt_curto = (passo["texto"][:40] + '..') if len(passo["texto"]) > 43 else passo["texto"]
            tk.Label(card, text=txt_curto, font=("Arial", 7), bg="white", wraplength=150).pack()
            
            btn_f = tk.Frame(card, bg="white")
            btn_f.pack(side="bottom", fill="x")
            
            tk.Button(btn_f, text="👁️", font=("Arial", 8), command=lambda p=passo: self.testar_holofote_individual(p)).pack(side="left", expand=True, fill="x")
            tk.Button(btn_f, text="🗑️", font=("Arial", 8), command=lambda idx=i: self.remover_etapa_tutorial(idx)).pack(side="left", expand=True, fill="x")

    def testar_holofote_individual(self, passo):
        """ Executa o efeito de escurecer a tela para validar o passo """
        HolofoteTutorial(self.root, passo["coords"], passo["texto"], passo["img_b64"])

    def remover_etapa_tutorial(self, indice):
        """ Exclui uma etapa da lista """
        if messagebox.askyesno("Confirmar", f"Excluir etapa {indice+1}?"):
            self.config["tutorial_dinamico"].pop(indice)
            salvar_config(self.config)
            self.renderizar_grade_tutorial()

    def log(self, mensagem):
        """ Escreve no terminal de log com timestamp """
        tempo = datetime.now().strftime("%H:%M:%S")
        msg_formatada = f"[{tempo}] {mensagem}\n"
        
        # Garante que o widget exista (será criado na Parte 10) antes de escrever
        if hasattr(self, 'txt_log'):
            self.txt_log.config(state="normal")
            self.txt_log.insert(tk.END, msg_formatada)
            self.txt_log.see(tk.END)
            self.txt_log.config(state="disabled")
        
        # Força o terminal a printar no console do Python também para debug
        print(msg_formatada.strip())

    def setup_log_ui(self):
        """ 
        Monta os widgets internos do painel de log lateral.
        Este painel ajuda no suporte técnico e dá transparência ao usuário.
        """
        # Cabeçalho do Log
        tk.Label(
            self.right_frame, 
            text="📟 LOG DE SISTEMA", 
            bg="#1e1e1e", 
            fg="#888888", 
            font=("Arial", 9, "bold")
        ).pack(pady=10)

        # Widget de Texto do Log
        self.txt_log = tk.Text(
            self.right_frame, 
            bg="#1e1e1e", 
            fg="#00FF00", # Verde Matrix
            font=("Consolas", 9), 
            state="disabled", 
            wrap="word", 
            padx=10,
            borderwidth=0,
            highlightthickness=0
        )
        self.txt_log.pack(fill="both", expand=True)

        # Botão para fechar o log (dentro do próprio log)
        btn_fechar = tk.Button(
            self.right_frame, 
            text="OCULTAR LOG ➡️", 
            bg="#333333", 
            fg="white", 
            font=("Arial", 7),
            command=self.toggle_terminal,
            relief="flat",
            padx=10
        )
        btn_fechar.pack(fill="x", pady=10, padx=20)

    def toggle_terminal(self):
        """ 
        Melhoria 12: Expande ou recolhe o painel de log lateralmente.
        Ajusta o tamanho da janela principal para comportar o novo painel.
        """
        if not self.terminal_visivel:
            # Se ainda não configuramos a UI interna do log, fazemos agora
            if not hasattr(self, 'txt_log'):
                self.setup_log_ui()
            
            self.right_frame.pack(side="right", fill="both", expand=False)
            self.btn_toggle_log.config(text="⬅️ Ocultar Log")
            self.terminal_visivel = True
            
            # Alarga a janela para 1220px para mostrar o log sem apertar a UI
            self.root.geometry("1220x800")
            self.log("Painel de diagnósticos expandido.")
        else:
            self.right_frame.pack_forget()
            self.btn_toggle_log.config(text="📟 Mostrar Log Avançado ➡️")
            self.terminal_visivel = False
            
            # Retorna ao tamanho original de 900px
            self.root.geometry("900x800")

    def adicionar_botao_log_principal(self):
        """ Posiciona o botão de gatilho do Log no rodapé da área esquerda """
        self.btn_toggle_log = tk.Button(
            self.left_frame, 
            text="📟 Mostrar Log Avançado ➡️", 
            command=self.toggle_terminal, 
            font=("Arial", 8),
            relief="groove"
        )
        self.btn_toggle_log.pack(side="bottom", anchor="se", pady=5)
        add_tip(self.btn_toggle_log, "Veja os detalhes técnicos do que o robô está fazendo")

    def setup_operacao_ui(self):
        """ 
        Cria a área de identificação da carga. 
        Implementa a Melhoria 13: Seletores que permitem novas entradas.
        """
        self.op_frame = tk.LabelFrame(
            self.left_frame, 
            text="📋 IDENTIFICAÇÃO DA MOVIMENTAÇÃO", 
            padx=15, 
            pady=15,
            font=("Arial", 9, "bold"),
            fg="#1a237e"
        )
        self.op_frame.pack(fill="x", pady=10)

        # 1. SELETOR DE HISTÓRICO
        tk.Label(self.op_frame, text="Histórico de Movimento:", font=("Arial", 9)).grid(row=0, column=0, sticky="w")
        
        self.cb_hist = ttk.Combobox(
            self.op_frame, 
            values=self.config["historicos"] + ["Outro..."], 
            width=35, 
            state="readonly"
        )
        self.cb_hist.grid(row=0, column=1, pady=8, padx=10)
        self.cb_hist.set(self.config["historicos"][0])
        self.cb_hist.bind("<<ComboboxSelected>>", self.gerenciar_outro_historico)
        add_tip(self.cb_hist, "Escolha o tipo de saída ou entrada de material")

        # 2. SELETOR DE CÓDIGO MI
        tk.Label(self.op_frame, text="Código MI (Destino):", font=("Arial", 9)).grid(row=1, column=0, sticky="w")
        
        self.cb_mi = ttk.Combobox(
            self.op_frame, 
            values=self.config["codigos_mi"] + ["Outro..."], 
            width=35, 
            state="readonly"
        )
        self.cb_mi.grid(row=1, column=1, pady=8, padx=10)
        self.cb_mi.set(self.config["codigos_mi"][0])
        self.cb_mi.bind("<<ComboboxSelected>>", self.gerenciar_outro_mi)
        add_tip(self.cb_mi, "Código do centro de custo ou projeto de destino")

    def gerenciar_outro_historico(self, event):
        """ Se 'Outro...' for selecionado, abre prompt para novo histórico """
        if self.cb_hist.get() == "Outro...":
            novo_val = simpledialog.askstring("Novo Histórico", "Digite a descrição do novo histórico:")
            if novo_val:
                novo_val = novo_val.strip().upper()
                if novo_val not in self.config["historicos"]:
                    self.config["historicos"].insert(0, novo_val)
                    salvar_config(self.config)
                    # Atualiza a lista do combo
                    self.cb_hist['values'] = self.config["historicos"] + ["Outro..."]
                self.cb_hist.set(novo_val)
                self.log(f"Novo histórico adicionado: {novo_val}")
            else:
                self.cb_hist.current(0) # Volta para o primeiro se cancelar

    def gerenciar_outro_mi(self, event):
        """ Se 'Outro...' for selecionado, abre prompt para novo Código MI """
        if self.cb_mi.get() == "Outro...":
            novo_val = simpledialog.askstring("Novo Código MI", "Digite o código MI:")
            if novo_val:
                novo_val = novo_val.strip().upper()
                if novo_val not in self.config["codigos_mi"]:
                    self.config["codigos_mi"].insert(0, novo_val)
                    salvar_config(self.config)
                    # Atualiza a lista do combo
                    self.cb_mi['values'] = self.config["codigos_mi"] + ["Outro..."]
                self.cb_mi.set(novo_val)
                self.log(f"Novo Código MI adicionado: {novo_val}")
            else:
                self.cb_mi.current(0)

    def setup_tabela_ui(self):
        """ 
        Configura a Treeview para conferência de itens.
        Implementa o visual de Checkout de Mercado com colunas reais.
        """
        self.lbl_preview = tk.Label(
            self.left_frame, 
            text="🔍 CONFERÊNCIA DE ITENS (Preview):", 
            font=("Arial", 9, "bold"),
            fg="#555555"
        )
        self.lbl_preview.pack(anchor="w", pady=(10, 0))

        # Container para a Tabela e Scrollbar
        self.table_container = tk.Frame(self.left_frame)
        self.table_container.pack(fill="both", expand=True, pady=5)

        # Configuração de Estilo (Visual Zebra)
        style = ttk.Style()
        # 'clam' permite customizar as cores do cabeçalho de forma mais fácil no Windows
        style.theme_use("clam") 
        style.configure("Treeview", 
                        rowheight=28, 
                        font=("Arial", 9),
                        background="#ffffff",
                        fieldbackground="#ffffff")
        style.map("Treeview", background=[('selected', '#3498db')]) # Cor ao clicar na linha

        # Definição das Colunas
        # Nome = Descrição, Cod Mi = Part Number, Unid = Soma Total, Qtd = Pacotes
        cols = ("Nome", "Cod Mi", "Unid", "Qtd")
        self.tree = ttk.Treeview(self.table_container, columns=cols, show="headings", height=10)
        
        # Cabeçalhos
        self.tree.heading("Nome", text="DESCRIÇÃO DO PRODUTO")
        self.tree.heading("Cod Mi", text="CÓD. ITEM")
        self.tree.heading("Unid", text="UNID. TOTAL")
        self.tree.heading("Qtd", text="VOLUMES")

        # Alinhamento e Largura das Colunas
        self.tree.column("Nome", width=350, anchor="w")
        self.tree.column("Cod Mi", width=100, anchor="center")
        self.tree.column("Unid", width=100, anchor="center")
        self.tree.column("Qtd", width=100, anchor="center")

        # Tags para cores intercaladas (Melhoria 9)
        self.tree.tag_configure('oddrow', background="#f2f2f2")
        self.tree.tag_configure('evenrow', background="#ffffff")

        # Barra de Rolagem Vertical
        self.scroll_y = ttk.Scrollbar(self.table_container, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.scroll_y.set)

        self.tree.pack(side="left", fill="both", expand=True)
        self.scroll_y.pack(side="right", fill="y")

    def limpar_tabela(self):
        """ Remove todos os dados atuais da tabela para um novo carregamento """
        for i in self.tree.get_children():
            self.tree.delete(i)

    def popular_tabela(self):
        """ 
        Insere os dados processados pelo Pandas na Treeview.
        Aplica o efeito Zebra dinamicamente usando um contador manual.
        """
        self.limpar_tabela()
        
        if self.dados_agrupados is not None:
            # i é o índice original do Pandas, r é a linha
            # Criamos um contador manual (idx) para evitar erros de tipagem
            for idx, (_, r) in enumerate(self.dados_agrupados.iterrows()):
                
                # Agora usamos 'idx', que é garantidamente um número inteiro (int)
                tag = 'evenrow' if idx % 2 == 0 else 'oddrow'
                
                # Garante que a descrição seja tratada como string
                desc_bruta = str(r.get('Desc', 'SEM DESCRIÇÃO'))
                desc_curta = (desc_bruta[:42] + "...") if len(desc_bruta) > 45 else desc_bruta
                
                # Insere na tabela
                self.tree.insert("", "end", values=(
                    desc_curta, 
                    r.get('Item', '0'), 
                    r.get('Unid_Total', 0), 
                    f"x{r.get('Qtd_Pacotes', 1)}"
                ), tags=(tag,))
            
            self.log(f"Visualização atualizada com {len(self.dados_agrupados)} itens únicos.")

    def extrair_data_dinamica(self, nome_arquivo):
        """ 
        Melhoria 11: Busca padrões de data no nome do arquivo (RegEx).
        Suporta DDMMYY ou DDMMYYYY.
        """
        # Procura 8 dígitos (DDMMYYYY) ou 6 dígitos (DDMMYY)
        padrao_8 = re.search(r"(\d{8})", nome_arquivo)
        padrao_6 = re.search(r"(\d{6})", nome_arquivo)
        
        try:
            if padrao_8:
                data_str = padrao_8.group(1)
                return datetime.strptime(data_str, "%d%m%Y").strftime("%d/%m/%Y")
            elif padrao_6:
                data_str = padrao_6.group(1)
                return datetime.strptime(data_str, "%d%m%y").strftime("%d/%m/%Y")
        except Exception as e:
            self.log(f"Erro ao converter data do nome: {e}")
        
        # Se não encontrar nada ou der erro, retorna a data de hoje como padrão
        return datetime.now().strftime("%d/%m/%Y")

    def ler_txt(self):
        """ 
        Processa o arquivo TXT bruto e transforma em dados estruturados.
        Implementa a Melhoria 10 (Agrupamento de Checkout).
        """
        caminho = filedialog.askopenfilename(
            title="Selecionar Relatório de Movimentação",
            filetypes=[("Arquivos de Texto", "*.txt"), ("Todos os arquivos", "*.*")]
        )
        
        if not caminho:
            return

        # Verifica se o arquivo é válido e não está vazio (Melhoria de Integridade)
        status_arq, msg_arq = validar_arquivo_saudavel(caminho)
        if not status_arq:
            messagebox.showerror("Erro de Integridade", msg_arq)
            return

        self.caminho_txt = caminho
        nome_arq = os.path.basename(caminho)
        self.log(f"Processando arquivo: {nome_arq}")

        # Extração de Data Automática (Se habilitado pelo usuário)
        if self.var_data_nome.get():
            data_sugerida = self.extrair_data_dinamica(nome_arq)
            self.log(f"Data identificada no nome: {data_sugerida}")
        
        try:
            import pandas as pd
            
            with open(caminho, 'r', encoding='utf-8') as f:
                # Remove espaços em branco e linhas vazias
                linhas = [l.strip() for l in f.readlines() if l.strip()]
            
            dados_brutos = []
            i = 0
            
            # Varredura inteligente por blocos
            while i < len(linhas):
                # O padrão esperado é: CÓDIGO - DESCRIÇÃO
                partes = linhas[i].upper().split(' - ')
                
                # Procura a quantidade: geralmente 3 ou 4 linhas abaixo do código
                idx_qtd = i+4 if (i+3 < len(linhas) and "/" in linhas[i+3]) else i+3
                
                if idx_qtd < len(linhas):
                    try:
                        # Extrai apenas os números da linha de quantidade (Sanitização)
                        qtd_limpa = re.sub(r'\D', '', linhas[idx_qtd])
                        
                        if qtd_limpa:
                            dados_brutos.append({
                                "Item": partes[0].strip(), # Corrigido: Acesso ao índice da lista
                                "Desc": partes[1].strip() if len(partes) > 1 else "SEM DESCRIÇÃO",
                                "Qtd": int(qtd_limpa) # CORRIGIDO: de qtd_qtd para qtd_limpa
                            })
                    except (ValueError, IndexError):
                        pass
                
                i = idx_qtd + 1

            if not dados_brutos:
                messagebox.showwarning("Aviso", "Nenhum item válido encontrado no TXT.\nVerifique o formato do relatório.")
                return

            # --- MOTOR PANDAS (CONSOLIDAÇÃO) ---
            # Agrupa por código e descrição
            df = pd.DataFrame(dados_brutos)
            self.dados_agrupados = df.groupby(['Item', 'Desc']).agg(
                Unid_Total=('Qtd', 'sum'),      # Soma matemática (Unidades)
                Qtd_Pacotes=('Qtd', 'count')    # Quantidade de volumes (Frequência)
            ).reset_index()

            # Atualiza a Treeview (Zebra) com os novos dados
            self.popular_tabela()
            
            # Habilita o botão de processamento final
            self.btn_processar.config(state="normal")
            self.log(f"Checkout pronto: {len(self.dados_agrupados)} itens únicos consolidados.")

        except Exception as e:
            self.log(f"Erro crítico na leitura: {str(e)}")
            messagebox.showerror("Erro de Leitura", f"Falha ao processar o arquivo:\n{e}")

    def conectar_google_sheets(self):
        """ 
        Estabelece a conexão com a API do Google Sheets.
        Requer gspread instalado e o JSON de credenciais configurado.
        """
        try:
            import gspread
            from google.oauth2.service_account import Credentials
            
            # 1. Carrega as credenciais do DNA do App (Parte 1 e 8)
            creds_info = json.loads(self.config["cloud"]["creds_json"])
            if not creds_info or creds_info == "{}":
                raise Exception("Credenciais JSON não encontradas. Configure na Engrenagem.")

            escopos = ['https://googleapis.com']
            creds = Credentials.from_service_account_info(creds_info, scopes=escopos)
            
            # 2. Autoriza o cliente
            cliente = gspread.authorize(creds)
            
            # 3. Abre a planilha pelo ID limpo (Parte 2)
            planilha_id = self.config["cloud"]["spreadsheet_id"]
            if not planilha_id:
                raise Exception("ID da Planilha não configurado.")
                
            planilha = cliente.open_by_key(planilha_id)
            aba = planilha.worksheet(self.config["cloud"]["worksheet_name"])
            
            return aba
            
        except Exception as e:
            self.log(f"Falha na conexão Google: {e}")
            raise e

    def validar_schema_planilha(self, aba):
        """ 
        Melhoria 2: Validação de Schema (Pré-Voo).
        Garante que a coluna de 'Quantidade' na planilha online é a correta.
        """
        self.log("Validando estrutura da planilha Google...")
        try:
            # Lê a primeira linha (cabeçalho)
            cabecalho = aba.row_values(1)
            coluna_alvo = self.config["coluna_qtd"] # Padrão: 8 (H)

            if len(cabecalho) < coluna_alvo:
                return False, f"A planilha tem apenas {len(cabecalho)} colunas. O app espera a coluna {coluna_alvo}."

            # Verifica se o nome da coluna faz sentido (evita escrever data em local de qtd)
            nome_coluna_atual = cabecalho[coluna_alvo - 1].upper()
            palavras_chave = ["QTD", "UNID", "QUANT", "SAIDA", "TOTAL"]
            
            if not any(pc in nome_coluna_atual for pc in palavras_chave):
                return False, f"Atenção: A coluna {coluna_alvo} na planilha chama-se '{nome_coluna_atual}'. Isso pode estar errado."

            return True, "Estrutura validada."
            
        except Exception as e:
            return False, f"Erro ao ler cabeçalho: {e}"

    def enviar_para_nuvem(self, data_mov, mi, hist):
        """ 
        Coordena o envio dos dados processados para o Google Sheets.
        Versão corrigida para evitar erros de tipagem (Hashable e Literal).
        """
        # Type Guard: Garante que os dados existem
        if self.dados_agrupados is None or self.dados_agrupados.empty:
            self.log("⚠️ Erro: Não há dados para enviar.")
            return False

        try:
            aba = self.conectar_google_sheets()
            
            # Executa a Melhoria 2 (Validação Estrutural)
            status_schema, msg_schema = self.validar_schema_planilha(aba)
            
            if not status_schema:
                self.log(f"⚠️ BLOQUEIO DE SEGURANÇA: {msg_schema}")
                if not messagebox.askyesno("Erro de Estrutura", f"{msg_schema}\n\nDeseja forçar o envio mesmo assim?"):
                    return False

            total_itens = len(self.dados_agrupados)
            self.log(f"Iniciando envio de {total_itens} itens...")
            
            # CORREÇÃO: Usamos enumerate para garantir que 'idx' seja int (Resolve o erro Hashable)
            for idx, (_, r) in enumerate(self.dados_agrupados.iterrows()):
                
                # Prepara a linha
                linha = [data_mov, mi, hist, str(r.get('Item', '')), "", "", "", r.get('Unid_Total', 0), "", r.get('Desc', '')]
                
                # CORREÇÃO: Passamos a opção de entrada. 
                # Se o erro 'Literal' persistir no editor, use: # type: ignore ao final da linha
                aba.append_row(linha, value_input_option='USER_ENTERED') # type: ignore
                
                # Atualiza progresso visual usando idx (que é int)
                perc = int(((idx + 1) / total_itens) * 100)
                self.bar_progresso['value'] = perc
                self.root.update_idletasks()

            self.log("✅ Sincronização com Google Sheets concluída!")
            return True

        except Exception as e:
            self.log(f"❌ Falha no envio para Nuvem: {e}")
            messagebox.showerror("Erro de Sincronização", f"Não foi possível enviar para o Google:\n{e}")
            return False

    def conectar_google_drive(self):
        """ 
        Estabelece conexão com a API do Google Drive v3.
        Utilizado para gerenciar o Carrossel de Versões e Backups de TXT.
        """
        try:
            from googleapiclient.discovery import build
            from google.oauth2.service_account import Credentials
            
            creds_info = json.loads(self.config["cloud"]["creds_json"])
            if not creds_info or creds_info == "{}":
                raise Exception("Credenciais JSON não encontradas.")

            escopos = ['https://googleapis.com']
            creds = Credentials.from_service_account_info(creds_info, scopes=escopos)
            
            # Constrói o serviço do Drive
            servico = build('drive', 'v3', credentials=creds)
            return servico
            
        except Exception as e:
            self.log(f"Falha na conexão Drive: {e}")
            return None

    def sincronizar_arquivo_drive(self, caminho_local):
        """ 
        Envia o arquivo TXT para a pasta de backup no Drive.
        Usa MD5 para evitar uploads desnecessários.
        """
        servico = self.conectar_google_drive()
        if not servico: return

        pasta_id = self.config["cloud"]["folder_backup_id"]
        if not pasta_id:
            self.log("⚠️ Backup Drive pulado: ID da pasta não configurado.")
            return

        nome_arquivo = os.path.basename(caminho_local)
        
        try:
            # 1. Calcula Hash MD5 local (Parte 2)
            hash_local = calcular_hash_arquivo_local(caminho_local)
            
            # 2. Busca se o arquivo já existe na pasta de backup
            query = f"name = '{nome_arquivo}' and '{pasta_id}' in parents and trashed = false"
            resposta = servico.files().list(q=query, fields="files(id, md5Checksum)").execute()
            arquivos_remotos = resposta.get('files', [])

            from googleapiclient.http import MediaFileUpload
            media = MediaFileUpload(caminho_local, mimetype='text/plain', resumable=True)

            if arquivos_remotos:
                file_id = arquivos_remotos[0]['id']
                hash_remoto = arquivos_remotos[0].get('md5Checksum')

                # Melhoria 1: Só faz upload se o conteúdo mudou
                if hash_local == hash_remoto:
                    self.log(f"☁️ Arquivo '{nome_arquivo}' já está sincronizado (MD5 idêntico).")
                    return

                # Atualiza arquivo existente (Gera nova versão para o Carrossel)
                servico.files().update(fileId=file_id, media_body=media).execute()
                self.log(f"✅ Versão atualizada no Drive: {nome_arquivo}")
            else:
                # Cria novo arquivo no Drive
                metadados = {'name': nome_arquivo, 'parents': [pasta_id]}
                servico.files().create(body=metadados, media_body=media, fields='id').execute()
                self.log(f"✅ Novo backup criado no Drive: {nome_arquivo}")

        except Exception as e:
            self.log(f"❌ Erro ao sincronizar com Drive: {e}")

    def baixar_arquivo_drive(self, arquivo_id, caminho_destino):
        """ 
        Faz o download binário de uma versão específica do Drive.
        Base para o Carrossel de Versões (Parte 17).
        """
        try:
            # 1. Obtém o serviço
            servico = self.conectar_google_drive()
            
            # TRAVA DE SEGURANÇA: Verifica se o serviço foi criado com sucesso
            if servico is None:
                self.log("❌ Erro: Serviço do Google Drive indisponível.")
                return False

            # 2. Prepara a requisição de download
            requisicao = servico.files().get_media(fileId=arquivo_id)
            
            # 3. Executa o download binário
            # Usamos BytesIO para capturar o fluxo de dados com segurança
            fh = BytesIO()
            from googleapiclient.http import MediaIoBaseDownload
            
            downloader = MediaIoBaseDownload(fh, requisicao)
            done = False
            while done is False:
                status, done = downloader.next_chunk()
                if status:
                    self.log(f"Baixando: {int(status.progress() * 100)}%")

            # 4. Grava no disco
            with open(caminho_destino, "wb") as f:
                f.write(fh.getvalue())
                
            self.log(f"✅ Versão restaurada com sucesso em: {caminho_destino}")
            return True

        except Exception as e:
            self.log(f"❌ Erro fatal no download do Drive: {e}")
            return False

    def fluxo_final(self):
        """ 
        Orquestra a exportação e sincronização total.
        Garante a liberdade do usuário respeitando os seletores (Local/Nuvem).
        """
        # 1. Validação de segurança inicial
        if self.dados_agrupados is None or self.dados_agrupados.empty:
            messagebox.showwarning("Aviso", "Não há dados para processar. Carregue um TXT primeiro.")
            return

        # 2. Captura definições da UI
        hist = self.cb_hist.get().upper()
        mi = self.cb_mi.get().upper()
        
        # Melhoria 11: Usa data do nome ou data atual
        data_mov = self.extrair_data_dinamica(os.path.basename(self.caminho_txt)) if self.var_data_nome.get() else datetime.now().strftime("%d/%m/%Y")
        
        self.log(f"Iniciando ciclo de exportação: {data_mov} | {mi} | {hist}")
        self.bar_progresso['value'] = 0
        self.btn_processar.config(state="disabled") # Evita cliques duplos

        sucesso_total = True

        try:
            # 3. EXPORTAÇÃO EXCEL LOCAL (Liberdade do Usuário)
            if self.var_local.get():
                self.log("Gerando arquivo Excel local...")
                import pandas as pd
                
                # Prepara DataFrame para o formato padrão de importação
                df_export = pd.DataFrame([[data_mov, mi, hist, r['Item'], r['Unid_Total']] for _, r in self.dados_agrupados.iterrows()],
                                         columns=['Data Movimento', 'Cod MI', 'Histórico', 'Cód. Item', 'Quantidade'])
                
                nome_arq = f"IMPORT_{mi}_{datetime.now().strftime('%H%M%S')}.xlsx"
                df_export.to_excel(nome_arq, index=False)
                self.log(f"✅ Excel local gerado: {nome_arq}")

            # 4. BACKUP DO ARQUIVO BRUTO NO DRIVE (Melhoria 1)
            # Independente da planilha, o TXT original deve ser salvo se houver pasta configurada
            if self.config["cloud"]["folder_backup_id"]:
                self.log("Sincronizando backup do TXT no Drive...")
                self.sincronizar_arquivo_drive(self.caminho_txt)

            # 5. SINCRONIZAÇÃO GOOGLE SHEETS (Liberdade do Usuário)
            if self.var_cloud.get():
                if not self.enviar_para_nuvem(data_mov, mi, hist):
                    sucesso_total = False
                    self.log("⚠️ Falha na sincronização com a nuvem.")
            
            # 6. FINALIZAÇÃO
            if sucesso_total:
                self.bar_progresso['value'] = 100
                self.log("🏁 Ciclo de movimentação finalizado com sucesso!")
                messagebox.showinfo("Sucesso", "Processamento concluído!\n\nDados protegidos e sincronizados.")
            else:
                self.log("❌ O processo terminou com avisos. Verifique o Log.")
                messagebox.showwarning("Concluído com Erros", "O processo terminou, mas a nuvem não foi atualizada.")

        except Exception as e:
            self.log(f"❌ ERRO CRÍTICO NO FLUXO: {e}")
            messagebox.showerror("Erro Fatal", f"Ocorreu um erro inesperado:\n{e}")
        
        finally:
            self.btn_processar.config(state="normal")

    def abrir_painel_config(self):
        """ 
        Abre a central de configurações. 
        Organiza Nuvem, Drive e Tutorial em abas profissionais.
        """
        win = tk.Toplevel(self.root)
        win.title("Configurações do Sistema ⚙️")
        win.geometry("650x750")
        win.grab_set() # Foca apenas nesta janela
        
        tabs = ttk.Notebook(win)
        tabs.pack(expand=True, fill="both", padx=10, pady=10)

        # --- ABA 1: CONECTIVIDADE NUVEM ---
        f_nuvem = tk.Frame(tabs, padx=20, pady=20)
        tabs.add(f_nuvem, text="🌐 Google Cloud & Drive")

        # ID da Planilha
        tk.Label(f_nuvem, text="Link ou ID da Planilha Google:", font=("Arial", 9, "bold")).pack(anchor="w")
        ent_sheet_id = tk.Entry(f_nuvem, width=65)
        ent_sheet_id.insert(0, self.config["cloud"]["spreadsheet_id"])
        ent_sheet_id.pack(pady=5)
        add_tip(ent_sheet_id, "Cole a URL da planilha ou apenas o código ID")

        # Nome da Aba
        tk.Label(f_nuvem, text="Nome da Aba (ex: Abril):", font=("Arial", 9, "bold")).pack(anchor="w", pady=(10, 0))
        ent_aba = tk.Entry(f_nuvem, width=30)
        ent_aba.insert(0, self.config["cloud"]["worksheet_name"])
        ent_aba.pack(pady=5, anchor="w")

        # JSON de Credenciais
        tk.Label(f_nuvem, text="Credenciais (Arquivo JSON):", font=("Arial", 9, "bold")).pack(anchor="w", pady=(10, 0))
        btn_json = tk.Button(f_nuvem, text="📂 SELECIONAR ARQUIVO JSON", bg="#607D8B", fg="white", 
                             command=self.importar_json_manual)
        btn_json.pack(pady=5, fill="x")
        
        self.lbl_json_status = tk.Label(f_nuvem, text="Status: JSON Carregado ✅" if len(self.config["cloud"]["creds_json"]) > 10 else "Status: Não configurado", 
                                        fg="green" if len(self.config["cloud"]["creds_json"]) > 10 else "red")
        self.lbl_json_status.pack()

        # ID da Pasta do Drive (Exclusivo aqui na Engrenagem)
        tk.Label(f_nuvem, text="ID da Pasta de Backup (Google Drive):", font=("Arial", 9, "bold")).pack(anchor="w", pady=(20, 0))
        ent_drive = tk.Entry(f_nuvem, width=65)
        ent_drive.insert(0, self.config["cloud"]["folder_backup_id"])
        ent_drive.pack(pady=5)
        add_tip(ent_drive, "ID da pasta onde os TXTs originais serão salvos")

        # Botão para o Carrossel (Parte 17)
        tk.Button(f_nuvem, text="📦 ABRIR CARROSSEL DE VERSÕES", bg="#1a73e8", fg="white",
                  command=lambda: self.abrir_carrossel_versoes(ent_drive.get())).pack(pady=15, fill="x")

        # --- ABA 2: GERENCIADOR DE TUTORIAL (FORMA DE GELO) ---
        f_tut_editor = tk.Frame(tabs, padx=10, pady=10)
        tabs.add(f_tut_editor, text="🧊 Gerenciar Tutorial")
        
        tk.Button(f_tut_editor, text="➕ ADICIONAR NOVA ETAPA", bg="#27ae60", fg="white",
                  font=("Arial", 9, "bold"), command=self.iniciar_criador_tutorial).pack(fill="x", pady=10)

        # CRIAÇÃO DO ATRIBUTO: Importante usar self. para ser acessível globalmente na classe
        self.frame_grid = tk.Frame(f_tut_editor, bg="#f5f5f5")
        self.frame_grid.pack(expand=True, fill="both")

        # RENDERIZAÇÃO INICIAL:
        self.renderizar_grade_tutorial()

        # --- RODAPÉ: BOTÃO SALVAR ---
        def salvar_e_fechar():
            self.config["cloud"]["spreadsheet_id"] = extrair_id_limpo(ent_sheet_id.get())
            self.config["cloud"]["worksheet_name"] = ent_aba.get().strip()
            self.config["cloud"]["folder_backup_id"] = ent_drive.get().strip()
            
            salvar_config(self.config)
            self.log("⚙️ Configurações de nuvem atualizadas.")
            messagebox.showinfo("Sucesso", "Configurações salvas com sucesso!")
            win.destroy()

        tk.Button(win, text="💾 SALVAR TUDO", bg="#2ecc71", fg="white", height=2, font=("Arial", 10, "bold"),
                  command=salvar_e_fechar).pack(fill="x", padx=20, pady=15)

    def importar_json_manual(self):
        """ Lê o arquivo JSON e embuti no config para independência total """
        caminho = filedialog.askopenfilename(filetypes=[("Arquivo JSON", "*.json")])
        if caminho:
            try:
                with open(caminho, 'r', encoding='utf-8') as f:
                    conteudo = f.read()
                    json.loads(conteudo) # Valida se é um JSON real
                    self.config["cloud"]["creds_json"] = conteudo
                    self.lbl_json_status.config(text="Status: JSON Carregado ✅", fg="green")
                    self.log("✅ Novo arquivo de credenciais vinculado.")
            except Exception as e:
                messagebox.showerror("Erro JSON", f"Arquivo inválido ou corrompido: {e}")

    def abrir_carrossel_versoes(self, folder_id):
        """ Inicializa a interface de recuperação de arquivos do Drive """
        if not folder_id:
            messagebox.showwarning("Aviso", "Por favor, insira o ID da pasta do Drive na Engrenagem primeiro!")
            return
        
        if len(self.config["cloud"]["creds_json"]) < 10:
            messagebox.showerror("Erro", "Credenciais JSON não configuradas na Engrenagem!")
            return

        try:
            servico = self.conectar_google_drive()
            if not servico: return
            
            # Pergunta qual arquivo deseja visualizar
            escolha = simpledialog.askstring("Recuperar Dados", "Qual arquivo TXT deseja buscar? (Ex: estoque.txt)")
            if escolha:
                if not escolha.lower().endswith(".txt"): escolha += ".txt"
                # No OctalinkApp, altere para passar o 'self.log' no final:
                JanelaCarrossel(self.root, servico, escolha, folder_id, self.baixar_arquivo_drive, self.log)
                
        except Exception as e:
            self.log(f"Erro ao acessar Drive: {e}")
            messagebox.showerror("Erro de Conexão", f"Não foi possível ler o histórico do Drive:\n{e}")

class JanelaCarrossel(tk.Toplevel):
    """ Interface estilo galeria para navegar pelos backups históricos """
    # 1. ADICIONE 'log_func=None' AQUI NOS ARGUMENTOS:
    def __init__(self, parent, drive_service, file_name, folder_id, callback_download, log_func=None):
        super().__init__(parent)
        
        # 2. AGORA O PYTHON SABE QUEM É log_func E PODE GUARDAR NO self.
        self.log_func = log_func 
        
        self.title(f"📦 Histórico: {file_name}")
        self.geometry("420x500")
        self.drive_service = drive_service
        self.file_name = file_name
        self.folder_id = folder_id
        self.callback_download = callback_download
        self.versoes = []
        self.indice = 0
        self.trava_id = None
        self.frame_grid = None
        
        self.setup_ui()
        self.buscar_versoes()

    def setup_ui(self):
        self.main_f = tk.Frame(self, padx=20, pady=20)
        self.main_f.pack(expand=True, fill="both")

        self.frame_grid = tk.Frame(self, bg="white") 
        self.frame_grid.pack(fill="both", expand=True)
        
        # O 'Cubo de Gelo' Visual (Card de Informação)
        self.card = tk.Frame(self.main_f, bd=1, relief="solid", padx=20, pady=25, bg="white")
        self.card.pack(fill="x", pady=20)
        
        self.lbl_info = tk.Label(self.card, text="Buscando versões...", bg="white", justify="left", font=("Consolas", 10))
        self.lbl_info.pack()
        
        # Controles de Navegação
        ctrl_f = tk.Frame(self.main_f)
        ctrl_f.pack(fill="x")
        
        self.btn_ant = tk.Button(ctrl_f, text="⬅️ Mais Antiga", command=self.anterior, state="disabled")
        self.btn_ant.pack(side="left", expand=True, fill="x", padx=2)
        
        self.btn_prox = tk.Button(ctrl_f, text="Mais Recente ➡️", command=self.proximo, state="disabled")
        self.btn_prox.pack(side="left", expand=True, fill="x", padx=2)
        
        # Trava de Segurança de 3 Segundos (Melhoria de UX)
        self.btn_restaurar = tk.Button(self, text="SEGURE PARA RESTAURAR (3s)", bg="#ff9800", fg="white", 
                                       font=("Arial", 10, "bold"), height=2)
        self.btn_restaurar.pack(side="bottom", fill="x", padx=20, pady=20)
        
        # Eventos da Trava
        self.btn_restaurar.bind("<ButtonPress-1>", self.iniciar_contagem)
        self.btn_restaurar.bind("<ButtonRelease-1>", self.parar_contagem)

    def renderizar_grade_tutorial(self):
        """ 
        Melhoria 7: Organiza as etapas em blocos visuais.
        Blindado contra erro 'winfo_children of None'.
        """
        # 1. PEGA O GRID E VALIDA (Resolve o erro do Linter)
        grid = getattr(self, "frame_grid", None)
        if grid is None:
            return

        # 2. LIMPA A GRADE (Agora o editor sabe que grid não é None)
        for child in grid.winfo_children():
            child.destroy()
            
        # 3. EXTRAI OS PASSOS (Resolve erro de MethodType)
        cfg = self.config if isinstance(self.config, dict) else {}
        passos_raw = cfg.get("tutorial_dinamico", [])
        passos = list(passos_raw)
        
        if not passos:
            tk.Label(grid, text="Nenhuma etapa configurada.", 
                     fg="grey", font=("Arial", 9, "italic"), bg="white").pack(pady=40)
            return

        # ... restante do código (use 'grid' no lugar de 'self.frame_grid')

        for i, p in enumerate(passos):
            # Garante que 'p' seja um dicionário para o editor não reclamar do .get()
            passo = dict(p) if isinstance(p, dict) else {}
            
            linha, coluna = i // 3, i % 3
            
            card = tk.Frame(grid, bd=1, relief="solid", width=170, height=210, padx=5, pady=5, bg="white")
            card.grid(row=linha, column=coluna, padx=8, pady=8)
            card.grid_propagate(False)
            
            tk.Label(card, text=f"ETAPA {i+1}", font=("Arial", 8, "bold"), bg="white", fg="#1a73e8").pack()
            
            # Miniatura da Imagem
            img_b64 = passo.get("img_b64")
            if img_b64:
                try:
                    img_dados = base64.b64decode(str(img_b64))
                    img_pil = Image.open(BytesIO(img_dados))
                    img_pil.thumbnail((140, 90))
                    foto = ImageTk.PhotoImage(img_pil)
                    
                    lbl_img = tk.Label(card, image=foto, bg="white")
                    # CORREÇÃO: setattr resolve o erro 'image is unknown' e mantém a referência
                    setattr(lbl_img, "image", foto) 
                    lbl_img.pack(pady=2)
                except Exception:
                    tk.Label(card, text="[Erro na Imagem]", bg="white", font=("Arial", 7)).pack()
            
            # Texto Resumido
            txt_original = str(passo.get("texto", ""))
            txt_curto = (txt_original[:40] + '..') if len(txt_original) > 43 else txt_original
            tk.Label(card, text=txt_curto, font=("Arial", 7), bg="white", wraplength=150).pack()
            
            # Botões de Ação (Acessando o pai para funções de remoção/teste se necessário)
            btn_f = tk.Frame(card, bg="white")
            btn_f.pack(side="bottom", fill="x", pady=2)
            
            tk.Button(btn_f, text="👁️", font=("Arial", 8), bg="#e8f0fe", relief="flat",
                      command=lambda p_step=passo: self.testar_holofote_individual(p_step)).pack(side="left", expand=True, fill="x", padx=1)
            
            tk.Button(btn_f, text="🗑️", font=("Arial", 8), bg="#fce8e6", relief="flat",
                      command=lambda idx=i: self.remover_etapa_tutorial(idx)).pack(side="left", expand=True, fill="x", padx=1)

    def buscar_versoes(self):
        try:
            query = f"name = '{self.file_name}' and '{self.folder_id}' in parents and trashed = false"
            res = self.drive_service.files().list(q=query, fields="files(id, name, modifiedTime, size)").execute()
            self.versoes = sorted(res.get('files', []), key=lambda x: x['modifiedTime'], reverse=True)[:5]
            
            if self.versoes: self.atualizar_visual()
            else: self.lbl_info.config(text=f"Nenhum backup encontrado para:\n{self.file_name}")
        except Exception as e:
            self.lbl_info.config(text=f"Erro ao buscar: {e}")

    def atualizar_visual(self):
        v = self.versoes[self.indice]
        dt = datetime.strptime(v['modifiedTime'], "%Y-%m-%dT%H:%M:%S.%fZ").strftime("%d/%m/%Y às %H:%M")
        tam = f"{int(v['size'])/1024:.2f} KB"
        self.lbl_info.config(text=f"📄 {self.file_name}\n\nREVISÃO {self.indice+1} de {len(self.versoes)}\n\n📅 DATA: {dt}\n⚖️ TAMANHO: {tam}")
        
        self.btn_ant.config(state="normal" if self.indice < len(self.versoes)-1 else "disabled")
        self.btn_prox.config(state="normal" if self.indice > 0 else "disabled")

    def anterior(self): self.indice += 1; self.atualizar_visual()
    def proximo(self): self.indice -= 1; self.atualizar_visual()

    def iniciar_contagem(self, event):
        self.btn_restaurar.config(bg="red", text="⏳ MANTENHA PRESSIONADO...")
        self.trava_id = self.after(3000, self.executar_restauracao)

    def parar_contagem(self, event):
        if self.trava_id:
            self.after_cancel(self.trava_id)
            self.trava_id = None
            self.btn_restaurar.config(bg="#ff9800", text="SEGURE PARA RESTAURAR (3s)")

    def executar_restauracao(self):
        """ Restauração com Lixeira Segura """
        v = self.versoes[self.indice]
        # Backup de Despedida (Renomeia o local antes de baixar o antigo)
        if os.path.exists(self.file_name):
            timestamp = datetime.now().strftime("%H%M%S")
            os.rename(self.file_name, f"OLD_{timestamp}_{self.file_name}")
        
        if self.callback_download(v['id'], self.file_name):
            messagebox.showinfo("Sucesso", f"Versão restaurada!\nO arquivo anterior foi salvo como 'OLD_...'")
            self.destroy()

    def remover_etapa_tutorial(self, indice):
        """ 
        Remove a etapa da lista e atualiza a grade visual.
        Corrigido para evitar conflito entre self.config e tk.config().
        """
        # Força o reconhecimento como dicionário para evitar erro de 'Overload'
        dados_config = self.config
        
        # Se por algum motivo o editor ainda reclamar, usamos esta proteção:
        if not isinstance(dados_config, dict):
            return

        pergunta = f"Deseja excluir a Etapa {indice + 1} definitivamente?"
        if messagebox.askyesno("Confirmar", pergunta):
            try:
                # Acessa a lista de forma segura
                lista_tutorial = dados_config.get("tutorial_dinamico", [])
                
                if 0 <= indice < len(lista_tutorial):
                    lista_tutorial.pop(indice)
                    
                    # Salva usando a função global da Parte 1
                    salvar_config(dados_config)
                    
                    # Atualiza a interface
                    self.renderizar_grade_tutorial()
                    
                    # CORREÇÃO: Usamos self.log_func (o argumento que passamos no init)
                    if self.log_func:
                        self.log_func(f"Etapa {indice + 1} removida do tutorial.")
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível remover: {e}")

    def testar_holofote_individual(self, passo):
        """ Aciona o Holofote apenas para validação da etapa """
        # Usamos self.master pois ele é a referência para o OctalinkApp (root)
        HolofoteTutorial(self.master, passo["coords"], passo["texto"], passo.get("img_b64"))

    def verificar_tutorial_inicial(self):
        """ 
        Verifica se deve rodar o tutorial de boas-vindas.
        Corrigido: Uso de self.log_func e tratamento de MethodType no config.
        """
        # 1. Resolve o erro de MethodType: Garante que cfg seja um dicionário
        cfg = self.config if isinstance(self.config, dict) else getattr(self, "config", {})
        
        # 2. Resolve o erro de .get(): Acessa a lista de forma segura
        passos = cfg.get("tutorial_dinamico", []) if hasattr(cfg, "get") else []
        
        if not passos:
            if self.log_func:
                self.log_func("Nenhum tutorial dinâmico encontrado para exibir.")
            return
            
        # 3. Resolve o erro de 'Attribute log unknown': Usa apenas self.log_func
        if self.log_func:
            self.log_func("🎬 Iniciando tutorial interativo...")
            
        self.executar_sequencia_tutorial(0)


    def executar_sequencia_tutorial(self, indice: int):
        """ 
        Percorre as etapas do tutorial uma a uma.
        Corrigido: Uso de self.log_func, self.master e tratamento de MethodType.
        """
        # 1. Garante que config seja tratado como dicionário (Resolve MethodType)
        cfg = self.config if isinstance(self.config, dict) else getattr(self, "config", {})
        passos = list(cfg.get("tutorial_dinamico", [])) if hasattr(cfg, "get") else []
        
        if indice >= len(passos):
            # Fim do Tutorial
            if isinstance(cfg, dict):
                cfg["tutorial_ativo"] = False
                salvar_config(cfg)
            
            if self.log_func:
                self.log_func("🏁 Tutorial concluído com sucesso.")
            
            messagebox.showinfo("Tutorial", "Guia finalizado!\n\nAgora você está pronto.")
            return

        # 2. Obtém o passo atual (Garante que seja dicionário para o editor)
        passo = dict(passos[indice]) if isinstance(passos[indice], dict) else {}
        
        # 3. Aciona o Holofote
        # CORREÇÃO: Usamos self.master (que é o root do App) e self.log_func
        guia = HolofoteTutorial(
            self.master, 
            passo.get("coords", (0,0,0,0)), 
            str(passo.get("texto", "")), 
            passo.get("img_b64")
        )
        
        # 4. Espera o fechamento para avançar
        self.master.wait_window(guia)
        self.after(300, lambda: self.executar_sequencia_tutorial(indice + 1))


    def iniciar_criador_tutorial(self):
        """ 
        Melhoria 5: Abre a Máscara de Seleção para criar um novo passo.
        O fluxo é: Desenhar na tela -> Comprimir -> Salvar na Grade.
        """
        def pos_selecao(coords, imagem_pil):
            # Abre a janela de compressão da Parte 5
            JanelaCompressaoTutorial(self.master, imagem_pil, coords, self.salvar_passo_no_dna)
            
        # Ativa o seletor estilo Windows
        MascaraSelecao(self.master, pos_selecao)

    def salvar_passo_no_dna(self, dados_passo: dict):
        """ 
        Insere o novo passo na lista e atualiza a interface visual.
        Versão corrigida para a classe JanelaCarrossel.
        """
        # 1. Força o reconhecimento do dicionário
        cfg = self.config if isinstance(self.config, dict) else getattr(self, "config", {})
        
        # 2. Garante que a lista existe e adiciona o dado
        if "tutorial_dinamico" not in cfg:
            cfg["tutorial_dinamico"] = []
        
        cfg["tutorial_dinamico"].append(dados_passo)
        
        # 3. Salva no arquivo físico
        salvar_config(cfg)
        
        # 4. Atualiza a grade visual
        self.renderizar_grade_tutorial()
        
        # 5. CORREÇÃO DO ERRO: Usamos 'self.log_func' (o argumento do init)
        if self.log_func:
            texto = str(dados_passo.get('texto', ''))[:20]
            self.log_func(f"✅ Novo passo salvo: {texto}...")


    def verificar_dependencias_criticas(self):
        """ 
        Verifica bibliotecas e reporta ao log_func da JanelaCarrossel.
        """
        modulos_faltantes = []
        try: import pandas
        except ImportError: modulos_faltantes.append("pandas")
        try: import PIL
        except ImportError: modulos_faltantes.append("pillow")
        try: import gspread
        except ImportError: modulos_faltantes.append("gspread")
        
        if modulos_faltantes:
            # CORREÇÃO: Usamos self.log_func (o atributo da classe JanelaCarrossel)
            if self.log_func:
                self.log_func(f"⚠️ Alerta: Módulos faltantes: {', '.join(modulos_faltantes)}")
            return False
        return True


# --- BLOCO DE EXECUÇÃO PRINCIPAL ---

if __name__ == "__main__":
    # 1. Configurações de DPI para Windows 10/11 (Evita interface embaçada)
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    # 2. Inicialização com Captura de Erros Fatais
    try:
        root = tk.Tk()
        
        # Define o estilo visual global para 'clam' (mais moderno no Windows)
        estilo = ttk.Style()
        estilo.theme_use("clam")
        
        # Instancia o aplicativo
        app = OctalinkApp(root)
        
        # Registra o sucesso da inicialização no log
        app.log("🚀 Sistema Octalink Pro v6.0 carregado.")
        
        # Verifica se há dependências instaladas
        if not app.verificar_dependencias_criticas():
            messagebox.showwarning("Aviso de Ambiente", 
                "Algumas bibliotecas podem estar faltando. Isso pode afetar o envio para nuvem.")

        # Inicia o loop da interface gráfica
        root.mainloop()

    except Exception as e:
        # Tratamento de erro nível 'Sênior': 
        # Se a janela principal nem abrir, exibe uma mensagem nativa do Windows.
        import ctypes
        import traceback
        
        erro_detalhado = traceback.format_exc()
        mensagem_erro = (
            f"O Octalink Pro encontrou um erro crítico e não pôde iniciar:\n\n"
            f"ERRO: {str(e)}\n\n"
            f"--- LOG TÉCNICO ---\n{erro_detalhado}"
        )
        
        # 0x10 é o ícone de 'Erro' no Windows MessageBox
        ctypes.windll.user32.MessageBoxW(0, mensagem_erro, "Erro de Inicialização Fatal", 0x10)
        sys.exit(1)
