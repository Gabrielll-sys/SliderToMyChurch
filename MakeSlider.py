import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import os
import platform
import subprocess
# import tkinter.font as tkFont # Manter comentado por enquanto
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN, MSO_AUTO_SIZE

# --- Constantes de Configuração ---
# (Copie as constantes da sua versão funcional mais recente - v23)
LARGURA_SLIDE = Inches(16)
ALTURA_SLIDE = Inches(9)
MARGEM_TEXTO = Inches(0.1)
NOME_FONTE_PADRAO = 'Arial'
FONTES_COMUNS_PPT = sorted([ # Lista da v23
    "Arial", "Arial Black", "Arial Narrow", "Bahnschrift", "Calibri", "Calibri Light",
    "Cambria", "Cambria Math", "Candara", "Candara Light", "Century", "Century Gothic",
    "Century Schoolbook", "Comic Sans MS", "Consolas", "Constantia", "Corbel",
    "Corbel Light", "Courier New", "Ebrima", "Franklin Gothic Medium",
    "Franklin Gothic Book", "Gabriola", "Gadugi", "Georgia", "Gill Sans MT", "Impact",
    "Ink Free", "Leelawadee UI", "Lucida Console", "Lucida Sans Unicode",
    "Malgun Gothic", "Marlett", "Microsoft Himalaya", "Microsoft JhengHei",
    "Microsoft JhengHei UI", "Microsoft New Tai Lue", "Microsoft PhagsPa",
    "Microsoft Sans Serif", "Microsoft Tai Le", "Microsoft YaHei", "Microsoft YaHei UI",
    "Microsoft Yi Baiti", "MingLiU-ExtB", "PMingLiU-ExtB", "MingLiU_HKSCS-ExtB",
    "Mongolian Baiti", "Montserrat", "MS Gothic", "MS UI Gothic", "MS PGothic", "MV Boli",
    "Myanmar Text", "Nirmala UI", "Palatino Linotype", "Rockwell", "Segoe Print",
    "Segoe Script", "Segoe UI", "Segoe UI Black", "Segoe UI Emoji", "Segoe UI Historic",
    "Segoe UI Semibold", "Segoe UI Semilight", "Segoe UI Symbol", "SimSun", "NSimSun",
    "SimSun-ExtB", "Sitka Banner", "Sitka Display", "Sitka Heading", "Sitka Small",
    "Sitka Subheading", "Sitka Text", "Sylfaen", "Symbol", "Tahoma", "Times New Roman",
    "Trebuchet MS", "Verdana", "Webdings", "Wingdings", "Wingdings 2", "Wingdings 3"
])
DEFAULT_TAMANHO_FONTE_MUSICA_REFRAO = Pt(80)
DEFAULT_TAMANHO_FONTE_MUSICA_VERSO = Pt(80)
DEFAULT_TAMANHO_FONTE_ACLAMACAO = Pt(70)
DEFAULT_TAMANHO_FONTE_ANTIFONA = Pt(66)
DEFAULT_TAMANHO_FONTE_LEITURA_TITULO_AMARELO = Pt(90)
DEFAULT_TAMANHO_FONTE_LEITURA_TEXTO_BRANCO = Pt(90)
DEFAULT_TAMANHO_FONTE_PALAVRA = Pt(80)
TAMANHO_TITULO_PARTE = Pt(60)
TAMANHO_FONTE_TITULO_INICIAL = Pt(90)
TAMANHO_FONTE_ORACAO = Pt(36)
LINHAS_POR_SLIDE_VERSO = 4
LINHAS_POR_SLIDE_ORACAO = 5
LINHAS_POR_SLIDE_LEITURA = 5
LINHAS_POR_SLIDE_PALAVRA = 6
LINHAS_POR_SLIDE_ACLAMACAO_TXT = 4
LINHAS_POR_SLIDE_ANTIFONA_TXT = 4
COR_REFRAO = RGBColor(255, 192, 0); COR_VERSO = RGBColor(255, 255, 255)
COR_TITULO = RGBColor(255, 192, 0); COR_FUNDO_PRETO = RGBColor(0, 0, 0)

# --- Textos Padrão (Fallback da GUI) ---
DEFAULT_TEXTS = {
    # ===========================================================
    # COLE O DICIONÁRIO DEFAULT_TEXTS COMPLETO AQUI
     "Entrada": {"titulo": "CANTO DE ENTRADA", "refrao": ["SENHOR, EIS AQUI O TEU","POVO QUE VEM IMPLORAR","TEU PERDÃO","É GRANDE O NOSSO","PECADO, PORÉM É MAIOR O","TEU CORAÇÃO"], "versos": [["SABENDO QUE","ACOLHESTE ZAQUEU, O","COBRADOR E ASSIM LHE","DEVOLVESTE TUA PAZ E","TEU AMOR TAMBÉM"],["NOS COLOCAMOS AO","LADO DOS QUE VÃO","BUSCAR NO TEU ALTAR A","GRAÇA DO PERDÃO"],["REVENDO EM MADALENA","A NOSSA PRÓPRIA FÉ","CHORANDO NOSSAS","PENAS DIANTE DOS TEUS","PÉS TAMBÉM"],["NÓS DESEJAMOS O","NOSSO AMOR TE DAR","PORQUE SÓ MUITO","AMOR NOS PODE","LIBERTAR"],["MOTIVOS TEMOS NÓS","DE SEMPRE CONFIAR,","DE ERGUER A NOSSA VOZ,","DE NÃO DESESPERAR,","OLHANDO AQUELE GESTO"],["QUE O BOM LADRÃO","SALVOU,","NÃO FOI, TAMBÉM, POR","NÓS,","TEU SANGUE QUE JORROU?"]]},
     "Ato Penitencial": {"titulo": "ATO PENITENCIAL", "refrao": [], "versos": []},
     "Palavra": {"titulo": "PALAVRA", "texto": ["DESÇA COMO A CHUVA A TUA","PALAVRA. QUE SE ESPALHE","COMO ORVALHO. COMO","CHUVISCO NA RELVA. COMO","AGUACEIRO NA GRAMA.","AMÉM!"]},
     "1ª Leitura": {"titulo_amarelo": ["PRIMEIRA LEITURA"], "texto_branco": ["Josue 5,9a.10-12"] },
     "Salmo": {"titulo_amarelo": ["SALMO 33 (34)"], "texto_branco": ["-Louvo a Vós Senhor"] },
     "2ª Leitura": {"titulo_amarelo": ["SEGUNDA LEITURA"], "texto_branco": ["2Corintíos 5,17-21"] },
     "Aclamação": {"titulo": "ACLAMAÇÃO DO EVANGELHO", "aclamacao_texto": ["Louvor e honra a vós, Senhor Jesus."], "antifona_texto": ["Vou levantar-me e vou a meu pai","e lhe direi: Meu pai, eu pequei","contra o céu e contra ti.","(Lc 15,1-3.11-32)"]},
     "Oferendas": {"titulo": "PREPARAÇÃO DAS OFERENDAS", "refrao": ["CONFIEI NO TEU AMOR E","VOLTEI, SIM, AQUI É MEU","LUGAR. EU GASTEI TEUS","BENS, Ó PAI, E TE DOU","ESTE PRANTO EM MINHAS","MÃOS"], "versos": [["MUITO ALEGRE EU TE","PEDI O QUE ERA MEU","PARTI, UM SONHO TÃO","NORMAL"],["DISSIPEI MEUS BENS E","O CORAÇÃO TAMBÉM","NO FIM, MEU MUNDO","ERA IRREAL"],["MIL AMIGOS CONHECI,","DISSERAM ADEUS","CAIU A SOLIDÃO EM","MIM"],["UM PATRÃO CRUEL","LEVOU-ME A REFLETIR","MEU PAI NÃO TRATA","UM SERVO ASSIM"],["NEM DEIXASTE-ME","FALAR DA INGRATIDÃO","MORREU NO ABRAÇO","O MAL QUE EU FIZ"],["FESTA, ROUPA NOVA,","ANEL, SANDÁLIA AOS","PÉS","VOLTEI À VIDA, SOU","FELIZ"]]},
     "Comunhão": {"titulo": "COMUNHÃO", "refrao": ["PROVAI E VEDE COMO DEUS É","BOM FELIZ DE QUEM NO SEU","AMOR CONFIA EM JESUS","CRISTO, SE FAZ GRAÇA E DOM","SE FAZ PALAVRA E PÃO ΝΑ","EUCARISTIA"], "versos": [["Ó PAI, TEU POVO BUSCA VIDA","NOVA NA DIREÇÃO DA PÁSCOA","DE JESUS EM NOSSA FRONTE, O","SINAL DAS CINZAS NA","CAMINHADA,","VEM SER FORÇA E LUZ"],["QUANDO, NA VIDA, ANDAMOS","NO DESERTO E A TENTAÇÃO","VEM NOS TIRAR A PAZ A","FORTALEZA E A PALAVRA","CERTA EM TI BUSCAMOS, DEUS","DE NOSSOS PAIS"],["PEREGRINAMOS ENTRE LUZ E","SOMBRAS A CRUZ NOS PESA, O","MAL NOS DESFIGURA MAS NA","ORAÇÃO E NA PALAVRA","ACHAMOS A TUA GRAÇA, QUE","NOS TRANSFIGURA"],["Ó DEUS, CONHECES NOSSO","SOFRIMENTO HÁ MUITA DOR, É","GRANDE A AFLIÇÃO","TRANSFORMA EM FESTA NOSSA","DOR-LAMENTO ACOLHE OS","FRUTOS BONS DA CONVERSÃO"],["QUANDO O PECADO NOS","CONSOME E FERE E EM TI","BUSCAMOS A PAZ DO PERDÃO","O NOSSO RIO DE AFLIÇÃO SE","PERDE NO MAR PROFUNDO DO","TEU CORAÇÃO"],["POR QUE FICAR EM COISAS JÁ","PASSADAS? O TEU PERDÃO","LIBERTA E NOS RENOVA O TEU","AMOR NOS ABRE NOVA","ESTRADA TRAZ ALEGRIA E PAZ,","NOS REVIGORA"]]},
     "Pós-Comunhão": {"titulo": "CANTO PÓS-COMUNHÃO", "refrao": [], "versos": []},
     "Final": {"titulo": "CANTO FINAL", "refrao": [], "versos": []},
    # ===========================================================
}

# --- Textos Fixos ---
# (Cole os textos fixos completos)
TEXTO_CREDO = [ "CREIO EM DEUS PAI TODO PODEROSO,", "CRIADOR DO CÉU E DA TERRA.", # ... etc
               "E EM JESUS CRISTO, SEU ÚNICO FILHO,", "NOSSO SENHOR,",
               "QUE FOI CONCEBIDO PELO PODER DO ESPÍRITO SANTO;", "NASCEU DA VIRGEM MARIA;",
               "PADECEU SOB PÔNCIO PILATOS,", "FOI CRUCIFICADO, MORTO E SEPULTADO.",
               "DESCEU À MANSÃO DOS MORTOS;", "RESSUSCITOU AO TERCEIRO DIA;",
               "SUBIU AOS CÉUS, ESTÁ SENTADO À DIREITA", "DE DEUS PAI TODO-PODEROSO,",
               "DONDE HÁ DE VIR A JULGAR OS VIVOS E OS MORTOS.",
               "CREIO NO ESPÍRITO SANTO,", "NA SANTA IGREJA CATÓLICA,",
               "NA COMUNHÃO DOS SANTOS,", "NA REMISSÃO DOS PECADOS,",
               "NA RESSURREIÇÃO DA CARNE,", "NA VIDA ETERNA.", "AMÉM." ]
TEXTO_ORACAO_SANTA_LUZIA = [ "Ó VIRGEM ADMIRÁVEL.", "CHEIA DE FIRMEZA E DE", "CONSTÂNCIA, QUE NEM", # ... etc
                            "AS POMPAS HUMANAS", "PUDERAM SEDUZIR,",
                            "NEM AS PROMESSAS,", "NEM AS AMEAÇAS,", "NEM A FORÇA BRUTA", "PUDERAM ABALAR,",
                            "PORQUE SOUBESTES SER", "O TEMPLO VIVO DO", "DIVINO ESPÍRITO SANTO.",
                            "O MUNDO CRISTÃO VOS", "PROCLAMOU ADVOGADA",
                            "DA LUZ DOS NOSSOS", "OLHOS. DEFENDEI-NOS,", "POIS, DE TODA MOLÉSTIA",
                            "QUE POSSA PREJUDICAR", "A NOSSA VISTA.",
                            "ALCANÇAI-NOS A LUZ", "SOBRENATURAL DA FÉ,", "ESPERANÇA E CARIDADE",
                            "PARA QUE NOS", "DESAPEGUEMOS", "DAS COISAS MATERIAIS", "E TERRESTRES",
                            "E TENHAMOS A FORÇA", "PARA VENCER O INIMIGO", "E ASSIM POSSAMOS",
                            "CONTEMPLAR-VOS NA", "GLÓRIA CELESTE. AMÉM." ]
TEXTO_AVISOS = ["Santuário Arquidiocesano Santa Luzia", "santuariosantaluziamg", "Santuário Arquidiocesano Santa Luzia MG"]


# --- Função Auxiliar (adiciona_texto_com_divisao - igual v23) ---
def adiciona_texto_com_divisao(prs, layout, linhas_originais, cor, tamanho_fonte, font_name, bold_state, italic_state, max_linhas, use_auto_size=True):
    # (Função igual à v23 - OMITIDA PARA BREVIDADE)
    if not linhas_originais or all(not s or s.isspace() for s in linhas_originais): return False
    linhas_validas = [linha for linha in linhas_originais if linha and not linha.isspace()]
    if not linhas_validas: return False
    linhas_restantes = linhas_validas[:]; slides_criados = 0
    while linhas_restantes:
        linhas_para_este_bloco = linhas_restantes[:max_linhas]; linhas_restantes = linhas_restantes[max_linhas:]
        texto_bloco_continuo = " ".join(linhas_para_este_bloco)
        if not texto_bloco_continuo.strip(): continue
        slide = prs.slides.add_slide(layout); slides_criados += 1
        esquerda = MARGEM_TEXTO; topo = MARGEM_TEXTO; largura = LARGURA_SLIDE - (2 * MARGEM_TEXTO); altura = ALTURA_SLIDE - (2 * MARGEM_TEXTO)
        caixa_texto = slide.shapes.add_textbox(esquerda, topo, largura, altura)
        frame_texto = caixa_texto.text_frame; frame_texto.clear(); frame_texto.word_wrap = True; frame_texto.vertical_anchor = MSO_ANCHOR.MIDDLE
        if use_auto_size: frame_texto.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        else: frame_texto.auto_size = MSO_AUTO_SIZE.NONE
        p = frame_texto.add_paragraph(); p.text = texto_bloco_continuo; p.alignment = PP_ALIGN.CENTER;
        p.font.name = font_name; p.font.size = tamanho_fonte; p.font.color.rgb = cor; p.font.bold = bold_state; p.font.italic = italic_state;
        try:
            caixa_texto.left = esquerda; caixa_texto.top = topo; caixa_texto.width = largura; caixa_texto.height = altura
            frame_texto.margin_bottom = Inches(0.05); frame_texto.margin_top = Inches(0.05); frame_texto.margin_left = Inches(0.1); frame_texto.margin_right = Inches(0.1)
        except Exception as e_resize: print(f"Aviso: resize caixa: {e_resize}")
    return slides_criados > 0


# --- Classe Principal da Aplicação GUI ---
class MassSlideGeneratorApp:
    def __init__(self, master):
        self.master = master
        master.title("Gerador de Slides v24 (Font Preview)")
        master.geometry("900x850") # Aumentar largura para preview
        title_frame = ttk.Frame(master, padding="10"); title_frame.pack(fill="x", padx=10, pady=(5, 0))
        ttk.Label(title_frame, text="Título Inicial da Apresentação:", font=('Arial', 11, 'bold')).pack(anchor='w')
        self.initial_title_widget = scrolledtext.ScrolledText(title_frame, height=3, width=90, wrap=tk.WORD, font=('Arial', 10)); self.initial_title_widget.pack(fill="x", expand=True, pady=(2, 5))
        self.initial_title_widget.insert(tk.END, "4º DOMINGO DA\nQUARESMA")
        self.notebook = ttk.Notebook(master)
        self.ordem_gui = [ "Entrada", "Ato Penitencial", "Palavra", "1ª Leitura", "Salmo", "2ª Leitura", "Aclamação", "Oferendas", "Comunhão", ]
        self.widgets_gui = {}

        for nome_parte in self.ordem_gui:
            frame = ttk.Frame(self.notebook, padding="10")
            self.notebook.add(frame, text=nome_parte)
            self.widgets_gui[nome_parte] = {}
            # Chama a função correta para criar widgets
            if nome_parte in ["1ª Leitura", "Salmo", "2ª Leitura"]:
                 self._criar_widgets_leitura_estilos(frame, nome_parte, self.widgets_gui[nome_parte])
            elif nome_parte == "Aclamação":
                 self._criar_widgets_aclamacao_estilos(frame, nome_parte, self.widgets_gui[nome_parte])
            elif nome_parte == "Palavra":
                 self._criar_widgets_palavra_estilos(frame, nome_parte, self.widgets_gui[nome_parte])
            else: # Assume musical padrão
                 self._criar_widgets_musica_estilos(frame, nome_parte, self.widgets_gui[nome_parte])

        self.notebook.pack(expand=True, fill="both", padx=10, pady=5)
        bottom_frame = ttk.Frame(master, padding="5"); bottom_frame.pack(fill="x", side="bottom", pady=(0, 5))
        self.status_label = ttk.Label(bottom_frame, text="Pronto."); self.status_label.pack(side="left", padx=10)
        self.generate_button = ttk.Button(bottom_frame, text="Gerar PowerPoint", command=self.gerar_apresentacao); self.generate_button.pack(side="right", padx=10)

    # --- Funções para Criar Widgets (Atualizadas com Preview) ---

    def _criar_controles_estilo(self, parent, data_dict, prefix_key, default_size_pt, default_bold=True, default_italic=False):
        """Cria controles de estilo (Fonte, Tamanho, N, I) e um Label de Preview."""
        style_frame = ttk.Frame(parent); style_frame.pack(fill='x', pady=2)

        # --- Controles (Esquerda) ---
        controls_frame = ttk.Frame(style_frame); controls_frame.pack(side='left', fill='x', expand=True)

        # Linha 1: Fonte + Tamanho
        line1_frame = ttk.Frame(controls_frame); line1_frame.pack(fill='x')
        ttk.Label(line1_frame, text="Fonte:", font=('Arial', 9)).pack(side='left', padx=(0, 5))
        font_var = tk.StringVar(value=NOME_FONTE_PADRAO)
        font_combo = ttk.Combobox(line1_frame, textvariable=font_var, values=FONTES_COMUNS_PPT, width=20, state='readonly')
        font_combo.pack(side='left', padx=(0, 10))
        data_dict[f"{prefix_key}_font_combo"] = font_combo

        ttk.Label(line1_frame, text="Tam:", font=('Arial', 9)).pack(side='left', padx=(5, 5))
        size_spinbox = ttk.Spinbox(line1_frame, from_=10, to=100, increment=1, width=4, justify='right', wrap=True)
        size_spinbox.set(int(default_size_pt.pt)); size_spinbox.pack(side='left', padx=(0, 10))
        data_dict[f"{prefix_key}_font_spinbox"] = size_spinbox

        # Linha 2: Negrito + Itálico
        line2_frame = ttk.Frame(controls_frame); line2_frame.pack(fill='x', pady=(2,0))
        bold_var = tk.BooleanVar(value=default_bold)
        bold_check = ttk.Checkbutton(line2_frame, text="Negrito", variable=bold_var); bold_check.pack(side='left', padx=(0, 5))
        data_dict[f"{prefix_key}_bold_var"] = bold_var
        italic_var = tk.BooleanVar(value=default_italic)
        italic_check = ttk.Checkbutton(line2_frame, text="Itálico", variable=italic_var); italic_check.pack(side='left')
        data_dict[f"{prefix_key}_italic_var"] = italic_var

        # --- Preview (Direita) ---
        preview_label = tk.Label(style_frame, text="Amostra", font=(NOME_FONTE_PADRAO, 12), width=15, relief="groove", borderwidth=1)
        preview_label.pack(side='right', padx=(10, 0), fill='y')
        data_dict[f"{prefix_key}_preview_label"] = preview_label

        # --- Função para Atualizar Preview ---
        def update_preview(*args):
            try:
                fname = font_var.get()
                fsize = int(size_spinbox.get())
                fbold = "bold" if bold_var.get() else "normal"
                fitalic = "italic" if italic_var.get() else "roman"
                preview_label.config(font=(fname, fsize if fsize < 18 else 18, fbold, fitalic), text=fname) # Limita tamanho no preview
            except (ValueError, tk.TclError):
                preview_label.config(font=(NOME_FONTE_PADRAO, 12), text="?") # Fallback

        # Vincula atualização aos controles
        font_var.trace_add("write", update_preview)
        size_spinbox.config(command=update_preview) # Atualiza ao clicar nas setas
        # Para atualizar ao digitar no spinbox (precisa de validação extra, opcional)
        # size_spinbox_var = tk.StringVar()
        # size_spinbox.config(textvariable=size_spinbox_var)
        # size_spinbox_var.trace_add("write", update_preview)
        bold_var.trace_add("write", update_preview)
        italic_var.trace_add("write", update_preview)

        update_preview() # Chama uma vez para definir estado inicial


    def _criar_widgets_musica_estilos(self, parent_frame, nome_parte, data_dict):
        # (Função igual à v23, mas chama _criar_controles_estilo)
        data_dict["tipo"] = "musica"; data_dict["titulo_geracao"] = DEFAULT_TEXTS.get(nome_parte, {}).get("titulo", nome_parte.upper())
        top_frame = ttk.Frame(parent_frame); top_frame.pack(fill='x', expand=False)
        data_dict["iniciar_refrao_var"] = tk.BooleanVar(value=False); chk = ttk.Checkbutton(top_frame, text="Iniciar com Refrão", variable=data_dict["iniciar_refrao_var"]); chk.pack(side='left', anchor='w', padx=5, pady=(5,2))
        refrao_frame = ttk.LabelFrame(parent_frame, text="Refrão (Amarelo)", padding=5); refrao_frame.pack(fill='x', expand=False, pady=5)
        self._criar_controles_estilo(refrao_frame, data_dict, "refrao", DEFAULT_TAMANHO_FONTE_MUSICA_REFRAO, default_bold=True)
        data_dict["refrao_widget"] = scrolledtext.ScrolledText(refrao_frame, height=6, width=90, wrap=tk.WORD, font=('Arial', 10)); data_dict["refrao_widget"].pack(fill="x", expand=True, padx=5, pady=(5,0))
        verso_frame = ttk.LabelFrame(parent_frame, text="Versos (Branco)", padding=5); verso_frame.pack(fill='both', expand=True, pady=5)
        self._criar_controles_estilo(verso_frame, data_dict, "verso", DEFAULT_TAMANHO_FONTE_MUSICA_VERSO, default_bold=True)
        data_dict["verso_widget"] = scrolledtext.ScrolledText(verso_frame, height=10, width=90, wrap=tk.WORD, font=('Arial', 10)); data_dict["verso_widget"].pack(fill="both", expand=True, padx=5, pady=(5,0))

    def _criar_widgets_leitura_estilos(self, parent_frame, nome_parte, data_dict):
        # (Função igual à v23, mas chama _criar_controles_estilo)
        data_dict["tipo"] = "leitura"; titulo_amarelo_padrao_list = DEFAULT_TEXTS.get(nome_parte, {}).get("titulo_amarelo", [])
        ref_frame = ttk.LabelFrame(parent_frame, text=f"Título/Referência - {nome_parte} (Amarelo)", padding=5); ref_frame.pack(fill='x', expand=False, pady=5)
        self._criar_controles_estilo(ref_frame, data_dict, "titulo_amarelo", DEFAULT_TAMANHO_FONTE_LEITURA_TITULO_AMARELO, default_bold=True)
        data_dict["titulo_amarelo_widget"] = scrolledtext.ScrolledText(ref_frame, height=4, width=90, wrap=tk.WORD, font=('Arial', 10)); data_dict["titulo_amarelo_widget"].pack(fill="x", expand=True, padx=5, pady=(5,0))
        if titulo_amarelo_padrao_list: data_dict["titulo_amarelo_widget"].insert(tk.END, "\n".join(titulo_amarelo_padrao_list))
        texto_frame = ttk.LabelFrame(parent_frame, text=f"Texto Principal - {nome_parte} (Branco)", padding=5); texto_frame.pack(fill='both', expand=True, pady=5)
        self._criar_controles_estilo(texto_frame, data_dict, "texto_branco", DEFAULT_TAMANHO_FONTE_LEITURA_TEXTO_BRANCO, default_bold=True)
        data_dict["texto_branco_widget"] = scrolledtext.ScrolledText(texto_frame, height=15, width=90, wrap=tk.WORD, font=('Arial', 10)); data_dict["texto_branco_widget"].pack(fill="both", expand=True, padx=5, pady=(5,0))
        default_txt = DEFAULT_TEXTS.get(nome_parte, {}).get("texto_branco", [])
        if default_txt: data_dict["texto_branco_widget"].insert(tk.END, "\n".join(default_txt))

    def _criar_widgets_aclamacao_estilos(self, parent_frame, nome_parte, data_dict):
        # (Função igual à v23, mas chama _criar_controles_estilo)
        data_dict["tipo"] = "aclamacao"; data_dict["titulo_geracao"] = DEFAULT_TEXTS.get(nome_parte, {}).get("titulo", nome_parte.upper())
        aclamacao_frame = ttk.LabelFrame(parent_frame, text="Aclamação (Amarelo - Superior)", padding=5); aclamacao_frame.pack(fill='x', expand=False, pady=5)
        self._criar_controles_estilo(aclamacao_frame, data_dict, "aclamacao", DEFAULT_TAMANHO_FONTE_ACLAMACAO, default_bold=True)
        data_dict["aclamacao_widget"] = scrolledtext.ScrolledText(aclamacao_frame, height=5, width=90, wrap=tk.WORD, font=('Arial', 10)); data_dict["aclamacao_widget"].pack(fill="x", expand=True, padx=5, pady=(5,0))
        default_aclamacao = DEFAULT_TEXTS.get(nome_parte, {}).get("aclamacao_texto", []);
        if default_aclamacao: data_dict["aclamacao_widget"].insert(tk.END, "\n".join(default_aclamacao))
        antifona_frame = ttk.LabelFrame(parent_frame, text="Antífona (Branco - Inferior)", padding=5); antifona_frame.pack(fill='both', expand=True, pady=5)
        self._criar_controles_estilo(antifona_frame, data_dict, "antifona", DEFAULT_TAMANHO_FONTE_ANTIFONA, default_bold=True)
        data_dict["antifona_widget"] = scrolledtext.ScrolledText(antifona_frame, height=12, width=90, wrap=tk.WORD, font=('Arial', 10)); data_dict["antifona_widget"].pack(fill="both", expand=True, padx=5, pady=(5,0))
        default_antifona = DEFAULT_TEXTS.get(nome_parte, {}).get("antifona_texto", []);
        if default_antifona: data_dict["antifona_widget"].insert(tk.END, "\n".join(default_antifona))

    def _criar_widgets_palavra_estilos(self, parent_frame, nome_parte, data_dict):
        # (Função igual à v23, mas chama _criar_controles_estilo)
        data_dict["tipo"] = "palavra"; data_dict["titulo_geracao"] = DEFAULT_TEXTS.get(nome_parte, {}).get("titulo", nome_parte.upper())
        texto_frame = ttk.LabelFrame(parent_frame, text=f"Texto - {nome_parte} (Amarelo)", padding=5); texto_frame.pack(fill='both', expand=True, pady=5)
        self._criar_controles_estilo(texto_frame, data_dict, "texto", DEFAULT_TAMANHO_FONTE_PALAVRA, default_bold=True)
        data_dict["texto_widget"] = scrolledtext.ScrolledText(texto_frame, height=20, width=90, wrap=tk.WORD, font=('Arial', 10)); data_dict["texto_widget"].pack(fill="both", expand=True, padx=5, pady=(5,0))
        default_txt = DEFAULT_TEXTS.get(nome_parte, {}).get("texto", [])
        if default_txt: data_dict["texto_widget"].insert(tk.END, "\n".join(default_txt))


    def gerar_apresentacao(self):
        # (Verificação inicial igual)
        if not DEFAULT_TEXTS or not TEXTO_CREDO or not TEXTO_ORACAO_SANTA_LUZIA: messagebox.showerror("Erro Config.", "Dicionários de texto padrão incompletos."); return

        self.status_label.config(text="Gerando apresentação...")
        self.master.update_idletasks()

        try:
            prs = Presentation()
            prs.slide_width = LARGURA_SLIDE; prs.slide_height = ALTURA_SLIDE
            slide_master = prs.slide_masters[0]
            background = slide_master.background; fill = background.fill
            fill.solid(); fill.fore_color.rgb = COR_FUNDO_PRETO
            layout_slide_branco = next((l for i, l in enumerate(prs.slide_layouts) if "Branco" in l.name or "Blank" in l.name), prs.slide_layouts[5 if len(prs.slide_layouts) > 5 else 0])

            # --- Funções Auxiliares (Atualizadas para usar estilos) ---
            # (Funções _get_font_style_from_gui, adicionar_secao_musical, adicionar_leitura_slide_unico, adicionar_aclamacao_slide_unico, adicionar_secao_fixa, adicionar_secao_palavra - iguais à v23 - OMITIDAS PARA BREVIDADE)
            def _get_font_style_from_gui(data_dict, prefix_key, default_size_pt, default_bold=True, default_italic=False):
                font_size = default_size_pt; font_name = NOME_FONTE_PADRAO; bold_state = default_bold; italic_state = default_italic
                spinbox = data_dict.get(f"{prefix_key}_font_spinbox"); combo = data_dict.get(f"{prefix_key}_font_combo"); bold_var = data_dict.get(f"{prefix_key}_bold_var"); italic_var = data_dict.get(f"{prefix_key}_italic_var")
                if spinbox:
                    try:
                        valor_int = int(spinbox.get())
                        font_size = Pt(valor_int) if 10 <= valor_int <= 100 else default_size_pt
                    except (tk.TclError, ValueError):
                        pass
                if combo: selected_font = combo.get(); font_name = selected_font if selected_font in FONTES_COMUNS_PPT else NOME_FONTE_PADRAO
                if bold_var:
                    try:
                        bold_state = bold_var.get()
                    except tk.TclError:
                        pass
                if italic_var:
                    try:
                        italic_state = italic_var.get()
                    except tk.TclError:
                        pass
                return font_size, font_name, bold_state, italic_state
            # <<< FUNÇÃO MUSICAL CORRIGIDA >>>
            def adicionar_secao_musical(nome_parte_gui):
                conteudo_adicionado_total = False
                widget_existe = nome_parte_gui in self.widgets_gui and self.widgets_gui[nome_parte_gui]["tipo"] == "musica"
                pegar_da_gui = widget_existe

                titulo_parte = DEFAULT_TEXTS.get(nome_parte_gui, {}).get("titulo", nome_parte_gui.upper())
                refrao_gui_str = ""; versos_gui_str = ""
                # Define padrões de estilo primeiro
                tamanho_refrao, nome_refrao, bold_refrao, italic_refrao = DEFAULT_TAMANHO_FONTE_MUSICA_REFRAO, NOME_FONTE_PADRAO, True, False
                tamanho_verso, nome_verso, bold_verso, italic_verso = DEFAULT_TAMANHO_FONTE_MUSICA_VERSO, NOME_FONTE_PADRAO, True, False
                iniciar_com_refrao = False # Padrão

                # Só lê da GUI se os widgets existirem
                if pegar_da_gui:
                    data_gui = self.widgets_gui[nome_parte_gui] # Define data_gui AQUI
                    refrao_gui_str = data_gui["refrao_widget"].get("1.0", tk.END).strip()
                    versos_gui_str = data_gui["verso_widget"].get("1.0", tk.END).strip()
                    tamanho_refrao, nome_refrao, bold_refrao, italic_refrao = _get_font_style_from_gui(data_gui, "refrao", DEFAULT_TAMANHO_FONTE_MUSICA_REFRAO, True, False)
                    tamanho_verso, nome_verso, bold_verso, italic_verso = _get_font_style_from_gui(data_gui, "verso", DEFAULT_TAMANHO_FONTE_MUSICA_VERSO, True, False)
                    # Lê o checkbox SOMENTE se data_gui foi definido
                    if "iniciar_refrao_var" in data_gui:
                        try: # Adiciona try-except para segurança extra
                           iniciar_com_refrao = data_gui["iniciar_refrao_var"].get()
                        except Exception as e_check:
                            print(f"Aviso: Erro ao ler checkbox para {nome_parte_gui}: {e_check}")
                            iniciar_com_refrao = False # Usa False em caso de erro
                # else: # Para Pós-Comunhão e Final, os padrões de estilo já foram definidos acima

                defaults = DEFAULT_TEXTS.get(nome_parte_gui, {})
                refrao_padrao = defaults.get("refrao", [])
                versos_padrao = defaults.get("versos", [])

                refrao_final = [l.strip() for l in refrao_gui_str.split('\n') if l.strip()] if refrao_gui_str else []
                versos_processados = []
                if versos_gui_str:
                    # (Lógica de split igual)
                    linhas_input = versos_gui_str.split('\n'); bloco_verso_atual = []
                    for linha in linhas_input:
                        linha_limpa = linha.strip()
                        if linha_limpa: bloco_verso_atual.append(linha_limpa)
                        elif bloco_verso_atual: versos_processados.append(bloco_verso_atual); bloco_verso_atual = []
                    if bloco_verso_atual: versos_processados.append(bloco_verso_atual)
                else:
                    # <<< ATENÇÃO: Usa padrão SÓ SE O CAMPO REFRÃO TAMBÉM ESTIVER VAZIO >>>
                    #    (Ou seja, não pegamos versos padrão se o usuário digitou um refrão mas não versos)
                    if not refrao_final: # Só pega versos padrão se não houver refrão da GUI
                         versos_processados = versos_padrao
                    # <<< FIM DA ATENÇÃO >>>


                # Adiciona seção APENAS se houver versos processados OU refrão da GUI
                if versos_processados or refrao_final:
                    titulo_adicionado = adiciona_texto_com_divisao(prs, layout_slide_branco, [titulo_parte], COR_TITULO, TAMANHO_TITULO_PARTE, NOME_FONTE_PADRAO, True, False, 5, use_auto_size=False)
                    if titulo_adicionado: conteudo_adicionado_total = True

                    # Adiciona refrão inicial se checkbox marcado E refrão existe
                    if iniciar_com_refrao and refrao_final:
                         refrao_adicionado = adiciona_texto_com_divisao(prs, layout_slide_branco, refrao_final, COR_REFRAO, tamanho_refrao, nome_refrao, bold_refrao, italic_refrao, LINHAS_POR_SLIDE_VERSO, use_auto_size=True)
                         if refrao_adicionado: conteudo_adicionado_total = True

                    # Loop de intercalação
                    for i, estrofe in enumerate(versos_processados):
                         verso_adicionado = adiciona_texto_com_divisao(prs, layout_slide_branco, estrofe, COR_VERSO, tamanho_verso, nome_verso, bold_verso, italic_verso, LINHAS_POR_SLIDE_VERSO, use_auto_size=True)
                         if verso_adicionado: conteudo_adicionado_total = True
                         # Adiciona refrão após estrofe (se existir)
                         if refrao_final:
                             refrao_adicionado = adiciona_texto_com_divisao(prs, layout_slide_branco, refrao_final, COR_REFRAO, tamanho_refrao, nome_refrao, bold_refrao, italic_refrao, LINHAS_POR_SLIDE_VERSO, use_auto_size=True)
                             if refrao_adicionado: conteudo_adicionado_total = True

                    # Caso especial: só refrão (digitado na GUI, sem versos)
                    if not versos_processados and refrao_final:
                        # Se não iniciou com refrão, adiciona aqui
                        if not iniciar_com_refrao:
                            refrao_adicionado = adiciona_texto_com_divisao(prs, layout_slide_branco, refrao_final, COR_REFRAO, tamanho_refrao, nome_refrao, bold_refrao, italic_refrao, LINHAS_POR_SLIDE_VERSO, use_auto_size=True)
                            if refrao_adicionado: conteudo_adicionado_total = True
                return conteudo_adicionado_total
            def adicionar_leitura_slide_unico(nome_parte_gui):
                 conteudo_adicionado = False
                 if nome_parte_gui in self.widgets_gui and self.widgets_gui[nome_parte_gui]["tipo"] == "leitura":
                     data_gui = self.widgets_gui[nome_parte_gui]; titulo_amarelo_gui_str = data_gui["titulo_amarelo_widget"].get("1.0", tk.END).strip(); texto_branco_gui_str = data_gui["texto_branco_widget"].get("1.0", tk.END).strip()
                     tamanho_titulo, nome_titulo, bold_titulo, italic_titulo = _get_font_style_from_gui(data_gui, "titulo_amarelo", DEFAULT_TAMANHO_FONTE_LEITURA_TITULO_AMARELO, True, False); tamanho_texto, nome_texto, bold_texto, italic_texto = _get_font_style_from_gui(data_gui, "texto_branco", DEFAULT_TAMANHO_FONTE_LEITURA_TEXTO_BRANCO, True, False)
                     defaults = DEFAULT_TEXTS.get(nome_parte_gui, {}); titulo_amarelo_padrao = defaults.get("titulo_amarelo", []); texto_branco_padrao = defaults.get("texto_branco", [])
                     titulo_amarelo_final = [l.strip() for l in titulo_amarelo_gui_str.split('\n') if l.strip()] if titulo_amarelo_gui_str else titulo_amarelo_padrao; texto_branco_final = [l.strip() for l in texto_branco_gui_str.split('\n') if l.strip()] if texto_branco_gui_str else texto_branco_padrao
                     if titulo_amarelo_final or texto_branco_final:
                         slide = prs.slides.add_slide(layout_slide_branco); conteudo_adicionado = True; esquerda = MARGEM_TEXTO; topo = MARGEM_TEXTO; largura = LARGURA_SLIDE - (2 * MARGEM_TEXTO); altura = ALTURA_SLIDE - (2 * MARGEM_TEXTO)
                         caixa_texto = slide.shapes.add_textbox(esquerda, topo, largura, altura); frame_texto = caixa_texto.text_frame; frame_texto.clear(); frame_texto.word_wrap = True; frame_texto.vertical_anchor = MSO_ANCHOR.MIDDLE; frame_texto.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
                         if titulo_amarelo_final:
                             p_titulo = frame_texto.add_paragraph(); texto_titulo_continuo = " ".join(titulo_amarelo_final); p_titulo.text = texto_titulo_continuo; p_titulo.alignment = PP_ALIGN.CENTER;
                             p_titulo.font.name = nome_titulo; p_titulo.font.size = tamanho_titulo; p_titulo.font.color.rgb = COR_REFRAO; p_titulo.font.bold = bold_titulo; p_titulo.font.italic = italic_titulo
                         if texto_branco_final:
                             if titulo_amarelo_final: p_espaco = frame_texto.add_paragraph(); p_espaco.text = ""
                             p_texto = frame_texto.add_paragraph(); texto_principal_continuo = " ".join(texto_branco_final); p_texto.text = texto_principal_continuo; p_texto.alignment = PP_ALIGN.CENTER;
                             p_texto.font.name = nome_texto; p_texto.font.size = tamanho_texto; p_texto.font.color.rgb = COR_VERSO; p_texto.font.bold = bold_texto; p_texto.font.italic = italic_texto
                         try: caixa_texto.left = esquerda; caixa_texto.top = topo; caixa_texto.width = largura; caixa_texto.height = altura; frame_texto.margin_bottom = Inches(0.05); frame_texto.margin_top = Inches(0.05); frame_texto.margin_left = Inches(0.1); frame_texto.margin_right = Inches(0.1)
                         except Exception as e_resize: print(f"Aviso: resize caixa leitura: {e_resize}")
                 return conteudo_adicionado
            def adicionar_aclamacao_slide_unico(nome_parte_gui):
                conteudo_adicionado = False
                if nome_parte_gui in self.widgets_gui and self.widgets_gui[nome_parte_gui]["tipo"] == "aclamacao":
                    data_gui = self.widgets_gui[nome_parte_gui]; titulo_secao = data_gui["titulo_geracao"]
                    aclamacao_gui_str = data_gui["aclamacao_widget"].get("1.0", tk.END).strip(); antifona_gui_str = data_gui["antifona_widget"].get("1.0", tk.END).strip()
                    tamanho_ac, nome_ac, bold_ac, italic_ac = _get_font_style_from_gui(data_gui, "aclamacao", DEFAULT_TAMANHO_FONTE_ACLAMACAO, True, False); tamanho_an, nome_an, bold_an, italic_an = _get_font_style_from_gui(data_gui, "antifona", DEFAULT_TAMANHO_FONTE_ANTIFONA, True, False)
                    defaults = DEFAULT_TEXTS.get(nome_parte_gui, {}); aclamacao_padrao = defaults.get("aclamacao_texto", []); antifona_padrao = defaults.get("antifona_texto", [])
                    aclamacao_final = [l.strip() for l in aclamacao_gui_str.split('\n') if l.strip()] if aclamacao_gui_str else aclamacao_padrao; antifona_final = [l.strip() for l in antifona_gui_str.split('\n') if l.strip()] if antifona_gui_str else antifona_padrao
                    titulo_adicionado = adiciona_texto_com_divisao(prs, layout_slide_branco, [titulo_secao], COR_TITULO, TAMANHO_TITULO_PARTE, NOME_FONTE_PADRAO, True, False, 5, use_auto_size=False)
                    if titulo_adicionado: conteudo_adicionado = True
                    ac_adicionado = False
                    if aclamacao_final: ac_adicionado = adiciona_texto_com_divisao(prs, layout_slide_branco, aclamacao_final, COR_REFRAO, tamanho_ac, nome_ac, bold_ac, italic_ac, LINHAS_POR_SLIDE_ACLAMACAO_TXT, use_auto_size=True);
                    if ac_adicionado: conteudo_adicionado = True
                    an_adicionado = False
                    if antifona_final: an_adicionado = adiciona_texto_com_divisao(prs, layout_slide_branco, antifona_final, COR_VERSO, tamanho_an, nome_an, bold_an, italic_an, LINHAS_POR_SLIDE_ANTIFONA_TXT, use_auto_size=True)
                    if an_adicionado: conteudo_adicionado = True
                return conteudo_adicionado
           # Função adicionar_secao_fixa (CORRIGIDA)
           # <<< FUNÇÃO ADICIONAR_SECAO_FIXA CORRIGIDA >>>
            def adicionar_secao_fixa(titulo_secao, texto_linhas, tamanho_fonte, linhas_por_slide, cor=COR_VERSO, add_separador=True, bold_content=True, use_auto_size_content=False):
                 titulo_adicionado = False
                 conteudo_adicionado = False

                 # 1. Adiciona o TÍTULO primeiro (se existir)
                 if titulo_secao and titulo_secao.strip():
                     titulo_adicionado = adiciona_texto_com_divisao(
                         prs, layout_slide_branco,
                         [titulo_secao], # Título como lista
                         COR_TITULO, # Cor do Título (Amarelo)
                         TAMANHO_TITULO_PARTE, # Tamanho Fixo do Título da Parte
                         NOME_FONTE_PADRAO, # Fonte Padrão
                         True, # Negrito=True para Título
                         False, # Itálico=False para Título
                         5, # Max linhas para Título
                         use_auto_size=False # Não autoajustar tamanho do título
                     )

                 # 2. Adiciona o CONTEÚDO depois (se existir)
                 if texto_linhas:
                     conteudo_adicionado = adiciona_texto_com_divisao(
                         prs, layout_slide_branco,
                         texto_linhas, # Conteúdo (lista de strings)
                         cor, # Cor do conteúdo
                         tamanho_fonte, # Tamanho da fonte para o conteúdo
                         NOME_FONTE_PADRAO, # Fonte Padrão para conteúdo fixo
                         bold_content, # Negrito para conteúdo
                         False, # Itálico=False para conteúdo fixo
                         linhas_por_slide, # Max linhas para conteúdo
                         use_auto_size=use_auto_size_content # Autoajuste para conteúdo
                     )

                 # 3. Adiciona o SEPARADOR se algo foi adicionado e add_separador=True
                 if (titulo_adicionado or conteudo_adicionado) and add_separador:
                     prs.slides.add_slide(layout_slide_branco)
                     print(f"Separador adicionado após {titulo_secao}")

                 return titulo_adicionado or conteudo_adicionado
            def adicionar_secao_palavra(nome_parte_gui):
                conteudo_adicionado_total = False
                if nome_parte_gui in self.widgets_gui and self.widgets_gui[nome_parte_gui]["tipo"] == "palavra":
                    data_gui = self.widgets_gui[nome_parte_gui]; titulo_secao = data_gui["titulo_geracao"]
                    texto_gui_str = data_gui["texto_widget"].get("1.0", tk.END).strip()
                    tamanho_fonte, nome_fonte, bold_state, italic_state = _get_font_style_from_gui(data_gui, "texto", DEFAULT_TAMANHO_FONTE_PALAVRA, True, False)
                    defaults = DEFAULT_TEXTS.get(nome_parte_gui, {}); texto_padrao = defaults.get("texto", [])
                    texto_final = [l.strip() for l in texto_gui_str.split('\n') if l.strip()] if texto_gui_str else texto_padrao
                    if texto_final:
                        titulo_adicionado = adiciona_texto_com_divisao(prs, layout_slide_branco, [titulo_secao], COR_TITULO, TAMANHO_TITULO_PARTE, NOME_FONTE_PADRAO, True, False, 5, use_auto_size=False)
                        if titulo_adicionado: conteudo_adicionado_total = True
                        texto_adicionado = adiciona_texto_com_divisao(prs, layout_slide_branco, texto_final, COR_TITULO, tamanho_fonte, nome_fonte, bold_state, italic_state, LINHAS_POR_SLIDE_PALAVRA, use_auto_size=True)
                        if texto_adicionado: conteudo_adicionado_total = True
                return conteudo_adicionado_total


            # --- Montagem da Apresentação ---
            # (Lógica de montagem igual à v19)
            ordem_final_geracao = [ "Entrada", "Ato Penitencial", "Palavra", "1ª Leitura", "Salmo", "2ª Leitura", "Aclamação", "CREDO", "PRECES", "Oferendas", "SANTO_TITULO", "ORACAO_EUCARISTICA", "CORDEIRO_TITULO", "Comunhão", "Pós-Comunhão", "SANTA_LUZIA", "AVISOS", "Final" ]
            initial_title_str = self.initial_title_widget.get("1.0", tk.END).strip(); initial_title_lines = [l.strip() for l in initial_title_str.split('\n') if l.strip()]
            if initial_title_lines:
                 titulo_inicial_adicionado = adiciona_texto_com_divisao(prs, layout_slide_branco, initial_title_lines, COR_TITULO, TAMANHO_FONTE_TITULO_INICIAL, NOME_FONTE_PADRAO, True, False, 5, use_auto_size=True)
                 if titulo_inicial_adicionado: prs.slides.add_slide(layout_slide_branco)
            for nome_parte in ordem_final_geracao:
                separador_necessario = False
                if nome_parte == "PALAVRA_INTRO": separador_necessario = adicionar_secao_fixa("PALAVRA", "Sem Conteúdo", Pt(80), 6, cor=COR_TITULO, add_separador=False,use_auto_size_content=True)
                elif nome_parte == "CREDO": separador_necessario = adicionar_secao_fixa("ORAÇÃO DO CREDO", TEXTO_CREDO, Pt(83), 3,use_auto_size_content=True,add_separador=False)
                elif nome_parte == "PRECES": separador_necessario = adicionar_secao_fixa("PRECES", [], TAMANHO_TITULO_PARTE, 1,),
                elif nome_parte == "ORACAO_EUCARISTICA": separador_necessario = adicionar_secao_fixa("ORAÇÃO EUCARÍSTICA", [], TAMANHO_TITULO_PARTE, 2)
                elif nome_parte == "SANTA_LUZIA": separador_necessario = adicionar_secao_fixa("ORAÇÃO A SANTA LUZIA", TEXTO_ORACAO_SANTA_LUZIA, Pt(80), LINHAS_POR_SLIDE_ORACAO,use_auto_size_content=True)
                elif nome_parte == "AVISOS": separador_necessario = adicionar_secao_fixa("AVISOS", TEXTO_AVISOS, Pt(90), 4, add_separador=False, bold_content=False, use_auto_size_content=True)
                elif nome_parte in ["1ª Leitura", "Salmo", "2ª Leitura"]: separador_necessario = adicionar_leitura_slide_unico(nome_parte)
                elif nome_parte == "Palavra": separador_necessario = adicionar_secao_palavra(nome_parte)
                elif nome_parte == "Aclamação": separador_necessario = adicionar_aclamacao_slide_unico(nome_parte)
                elif nome_parte in self.widgets_gui or nome_parte in ["Pós-Comunhão", "Final"]: separador_necessario = adicionar_secao_musical(nome_parte)
                if separador_necessario and nome_parte != ordem_final_geracao[-1]:
                     prs.slides.add_slide(layout_slide_branco); print(f"Separador após {nome_parte}")


            # --- Salvar e Abrir ---
            # (Lógica de salvar e abrir igual à v21)
            filepath = filedialog.asksaveasfilename( defaultextension=".pptx", filetypes=[("PowerPoint Presentations", "*.pptx"), ("All Files", "*.*")], title="Salvar Apresentação Como...", initialfile="Missa_Gerada_v23.pptx" )
            if not filepath: self.status_label.config(text="Geração cancelada."); return
            prs.save(filepath)
            self.status_label.config(text=f"Salvo: {os.path.basename(filepath)}")
            try:
                if platform.system() == 'Darwin': subprocess.call(('open', filepath))
                elif platform.system() == 'Windows': os.startfile(filepath)
                else: subprocess.call(('xdg-open', filepath))
                print(f"Tentando abrir {filepath}")
            except Exception as e_open: print(f"Não foi possível abrir o arquivo automaticamente: {e_open}")
            messagebox.showinfo("Sucesso", f"Apresentação '{os.path.basename(filepath)}' gerada e salva!")


        except Exception as e:
            # (Lógica de erro igual)
            self.status_label.config(text="Erro durante a geração!")
            print(f"Erro detalhado: {e}"); import traceback; traceback.print_exc()
            messagebox.showerror("Erro", f"Ocorreu um erro:\n{e}\nVerifique o console.")
        finally:
            self.master.update_idletasks()

# --- Iniciar a Aplicação ---
if __name__ == "__main__":
    # (Colar dicionários completos aqui)
    if 'Entrada' not in DEFAULT_TEXTS or 'Palavra' not in DEFAULT_TEXTS or '1ª Leitura' not in DEFAULT_TEXTS or 'Aclamação' not in DEFAULT_TEXTS or not TEXTO_CREDO or not TEXTO_ORACAO_SANTA_LUZIA: print("ERRO CRÍTICO: Dicionários de texto padrão não estão completos!"); exit()
    root = tk.Tk()
    # Tenta listar fontes do sistema (opcional)
    # try: FONTES_COMUNS_PPT = sorted(list(tkFont.families())); print(f"Fontes do sistema: {len(FONTES_COMUNS_PPT)}")
    # except Exception as e_font: print(f"Erro ao listar fontes: {e_font}")
    app = MassSlideGeneratorApp(root)
    root.mainloop()


   