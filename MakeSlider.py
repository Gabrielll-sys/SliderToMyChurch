import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN, MSO_AUTO_SIZE

# --- Constantes de Configuração ---
# (Mantendo valores da v17/v18, mas ajustando nomes para clareza)
LARGURA_SLIDE = Inches(16)
ALTURA_SLIDE = Inches(9)
MARGEM_TEXTO = Inches(0.1)
DEFAULT_TAMANHO_FONTE_MUSICA_REFRAO = Pt(80)
DEFAULT_TAMANHO_FONTE_MUSICA_VERSO = Pt(80)
DEFAULT_TAMANHO_FONTE_ACLAMACAO = Pt(54)
DEFAULT_TAMANHO_FONTE_ANTIFONA = Pt(48)
# Tamanhos para Leitura/Salmo
DEFAULT_TAMANHO_FONTE_LEITURA_TITULO_AMARELO = Pt(80) # Tamanho para a Ref/Título (Amarelo)
DEFAULT_TAMANHO_FONTE_LEITURA_TEXTO_BRANCO = Pt(54)  # Tamanho para Texto Principal (Branco)
TAMANHO_TITULO_PARTE = Pt(60) # Para títulos de seção NÃO editáveis (PALAVRA, CREDO, etc.)
TAMANHO_FONTE_TITULO_INICIAL = Pt(90)
TAMANHO_FONTE_ORACAO = Pt(36)
LINHAS_POR_SLIDE_VERSO = 4
LINHAS_POR_SLIDE_ORACAO = 5
LINHAS_POR_SLIDE_LEITURA = 5 # Limite para o texto branco da leitura
NOME_FONTE = 'Arial'
COR_REFRAO = RGBColor(255, 192, 0) # Amarelo
COR_VERSO = RGBColor(255, 255, 255) # Branco
COR_TITULO = RGBColor(255, 192, 0) # Amarelo
COR_FUNDO_PRETO = RGBColor(0, 0, 0)

# --- Textos Padrão (Fallback da GUI) ---
DEFAULT_TEXTS = {
    # ===========================================================
    # COLE O DICIONÁRIO DEFAULT_TEXTS COMPLETO AQUI
     "Entrada": {"titulo": "CANTO DE ENTRADA", "refrao": ["SENHOR, EIS AQUI O TEU","POVO QUE VEM IMPLORAR","TEU PERDÃO","É GRANDE O NOSSO","PECADO, PORÉM É MAIOR O","TEU CORAÇÃO"], "versos": [["SABENDO QUE","ACOLHESTE ZAQUEU, O","COBRADOR E ASSIM LHE","DEVOLVESTE TUA PAZ E","TEU AMOR TAMBÉM"],["NOS COLOCAMOS AO","LADO DOS QUE VÃO","BUSCAR NO TEU ALTAR A","GRAÇA DO PERDÃO"],["REVENDO EM MADALENA","A NOSSA PRÓPRIA FÉ","CHORANDO NOSSAS","PENAS DIANTE DOS TEUS","PÉS TAMBÉM"],["NÓS DESEJAMOS O","NOSSO AMOR TE DAR","PORQUE SÓ MUITO","AMOR NOS PODE","LIBERTAR"],["MOTIVOS TEMOS NÓS","DE SEMPRE CONFIAR,","DE ERGUER A NOSSA VOZ,","DE NÃO DESESPERAR,","OLHANDO AQUELE GESTO"],["QUE O BOM LADRÃO","SALVOU,","NÃO FOI, TAMBÉM, POR","NÓS,","TEU SANGUE QUE JORROU?"]]},
     "Ato Penitencial": {"titulo": "ATO PENITENCIAL", "refrao": [], "versos": []},
     # <<< LEITURAS: 'titulo_amarelo' e 'texto_branco' >>>
     "1ª Leitura": {"titulo_amarelo": ["PRIMEIRA LEITURA", "Josue 5,9a.10-12"], "texto_branco": [] }, # Juntando título e ref padrão
     "Salmo": {"titulo_amarelo": ["SALMO 33 (34)", "Salmo(33 e 34)"], "texto_branco": ["- Louvo a Vós Senhor"] },
     "2ª Leitura": {"titulo_amarelo": ["SEGUNDA LEITURA", "2Corintíos 5,17-21"], "texto_branco": [] },
     "Aclamação": {"titulo": "ACLAMAÇÃO DO EVANGELHO", "aclamacao_texto": ["Louvor e honra a vós, Senhor Jesus."], "antifona_texto": ["Vou levantar-me e vou a meu pai","e lhe direi: Meu pai, eu pequei","contra o céu e contra ti.","(Lc 15,1-3.11-32)"]},
     "Oferendas": {"titulo": "PREPARAÇÃO DAS OFERENDAS", "refrao": ["CONFIEI NO TEU AMOR E","VOLTEI, SIM, AQUI É MEU","LUGAR. EU GASTEI TEUS","BENS, Ó PAI, E TE DOU","ESTE PRANTO EM MINHAS","MÃOS"], "versos": [["MUITO ALEGRE EU TE","PEDI O QUE ERA MEU","PARTI, UM SONHO TÃO","NORMAL"],["DISSIPEI MEUS BENS E","O CORAÇÃO TAMBÉM","NO FIM, MEU MUNDO","ERA IRREAL"],["MIL AMIGOS CONHECI,","DISSERAM ADEUS","CAIU A SOLIDÃO EM","MIM"],["UM PATRÃO CRUEL","LEVOU-ME A REFLETIR","MEU PAI NÃO TRATA","UM SERVO ASSIM"],["NEM DEIXASTE-ME","FALAR DA INGRATIDÃO","MORREU NO ABRAÇO","O MAL QUE EU FIZ"],["FESTA, ROUPA NOVA,","ANEL, SANDÁLIA AOS","PÉS","VOLTEI À VIDA, SOU","FELIZ"]]},
     "Comunhão": {"titulo": "COMUNHÃO", "refrao": ["PROVAI E VEDE COMO DEUS É","BOM FELIZ DE QUEM NO SEU","AMOR CONFIA EM JESUS","CRISTO, SE FAZ GRAÇA E DOM","SE FAZ PALAVRA E PÃO ΝΑ","EUCARISTIA"], "versos": [["Ó PAI, TEU POVO BUSCA VIDA","NOVA NA DIREÇÃO DA PÁSCOA","DE JESUS EM NOSSA FRONTE, O","SINAL DAS CINZAS NA","CAMINHADA,","VEM SER FORÇA E LUZ"],["QUANDO, NA VIDA, ANDAMOS","NO DESERTO E A TENTAÇÃO","VEM NOS TIRAR A PAZ A","FORTALEZA E A PALAVRA","CERTA EM TI BUSCAMOS, DEUS","DE NOSSOS PAIS"],["PEREGRINAMOS ENTRE LUZ E","SOMBRAS A CRUZ NOS PESA, O","MAL NOS DESFIGURA MAS NA","ORAÇÃO E NA PALAVRA","ACHAMOS A TUA GRAÇA, QUE","NOS TRANSFIGURA"],["Ó DEUS, CONHECES NOSSO","SOFRIMENTO HÁ MUITA DOR, É","GRANDE A AFLIÇÃO","TRANSFORMA EM FESTA NOSSA","DOR-LAMENTO ACOLHE OS","FRUTOS BONS DA CONVERSÃO"],["QUANDO O PECADO NOS","CONSOME E FERE E EM TI","BUSCAMOS A PAZ DO PERDÃO","O NOSSO RIO DE AFLIÇÃO SE","PERDE NO MAR PROFUNDO DO","TEU CORAÇÃO"],["POR QUE FICAR EM COISAS JÁ","PASSADAS? O TEU PERDÃO","LIBERTA E NOS RENOVA O TEU","AMOR NOS ABRE NOVA","ESTRADA TRAZ ALEGRIA E PAZ,","NOS REVIGORA"]]},
     "Pós-Comunhão": {"titulo": "CANTO PÓS-COMUNHÃO", "refrao": [], "versos": []},
     "Final": {"titulo": "CANTO FINAL", "refrao": [], "versos": []},
    # ===========================================================
}


# --- Textos Fixos ---
# (Cole os textos fixos completos)
TEXTO_PALAVRA_INTRO = ["DESÇA COMO A CHUVA A TUA","PALAVRA. QUE SE ESPALHE","COMO ORVALHO. COMO","CHUVISCO NA RELVA. COMO","AGUACEIRO NA GRAMA.","AMÉM!"]
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


# --- Função Auxiliar (adiciona_texto_com_divisao - igual v14) ---
def adiciona_texto_com_divisao(prs, layout, linhas_originais, cor, tamanho_fonte, max_linhas, bold=True, use_auto_size=True):
    # (Função igual à v14 - OMITIDA PARA BREVIDADE)
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
        p = frame_texto.add_paragraph(); p.text = texto_bloco_continuo; p.alignment = PP_ALIGN.CENTER; p.font.name = NOME_FONTE; p.font.size = tamanho_fonte; p.font.color.rgb = cor; p.font.bold = bold
        try:
            caixa_texto.left = esquerda; caixa_texto.top = topo; caixa_texto.width = largura; caixa_texto.height = altura
            frame_texto.margin_bottom = Inches(0.05); frame_texto.margin_top = Inches(0.05)
            frame_texto.margin_left = Inches(0.1); frame_texto.margin_right = Inches(0.1)
        except Exception as e_resize: print(f"Aviso: resize caixa: {e_resize}")
    return slides_criados > 0


# --- Classe Principal da Aplicação GUI ---
class MassSlideGeneratorApp:
    def __init__(self, master):
        self.master = master
        master.title("Gerador de Slides v19 (Leitura Título Replace)")
        master.geometry("850x800")
        title_frame = ttk.Frame(master, padding="10"); title_frame.pack(fill="x", padx=10, pady=(5, 0))
        ttk.Label(title_frame, text="Título Inicial da Apresentação:", font=('Arial', 11, 'bold')).pack(anchor='w')
        self.initial_title_widget = scrolledtext.ScrolledText(title_frame, height=3, width=90, wrap=tk.WORD, font=('Arial', 10)); self.initial_title_widget.pack(fill="x", expand=True, pady=(2, 5))
        self.initial_title_widget.insert(tk.END, "4º DOMINGO DA\nQUARESMA")
        self.notebook = ttk.Notebook(master)
        self.ordem_gui = [ "Entrada", "Ato Penitencial", "1ª Leitura", "Salmo", "2ª Leitura", "Aclamação", "Oferendas", "Comunhão", ]
        self.widgets_gui = {}

        for nome_parte in self.ordem_gui:
            frame = ttk.Frame(self.notebook, padding="10")
            self.notebook.add(frame, text=nome_parte)
            self.widgets_gui[nome_parte] = {}
            # <<< Pega o título amarelo PADRÃO para exibir na GUI >>>
            titulo_amarelo_padrao = DEFAULT_TEXTS.get(nome_parte, {}).get("titulo_amarelo", [nome_parte.upper()])

            if nome_parte in ["1ª Leitura", "Salmo", "2ª Leitura"]:
                 # Passa o título amarelo padrão para a função de widget
                 self._criar_widgets_leitura_com_fonte(frame, nome_parte, self.widgets_gui[nome_parte], titulo_amarelo_padrao)
            elif nome_parte == "Aclamação":
                 titulo_sugerido = DEFAULT_TEXTS.get(nome_parte, {}).get("titulo", nome_parte.upper())
                 self._criar_widgets_aclamacao_com_fonte(frame, nome_parte, self.widgets_gui[nome_parte], titulo_sugerido)
            else: # Assume musical padrão
                 titulo_sugerido = DEFAULT_TEXTS.get(nome_parte, {}).get("titulo", nome_parte.upper())
                 self._criar_widgets_musica_com_fonte(frame, nome_parte, self.widgets_gui[nome_parte], titulo_sugerido)

        self.notebook.pack(expand=True, fill="both", padx=10, pady=5)
        bottom_frame = ttk.Frame(master, padding="5"); bottom_frame.pack(fill="x", side="bottom", pady=(0, 5))
        self.status_label = ttk.Label(bottom_frame, text="Pronto."); self.status_label.pack(side="left", padx=10)
        self.generate_button = ttk.Button(bottom_frame, text="Gerar PowerPoint", command=self.gerar_apresentacao); self.generate_button.pack(side="right", padx=10)

    # --- Funções para Criar Widgets ---
    def _criar_spinbox_fonte(self, parent, label_text, default_value_pt, data_dict, key):
        # (Função igual à v16)
        font_frame = ttk.Frame(parent); font_frame.pack(fill='x', pady=(2,5))
        ttk.Label(font_frame, text=label_text, font=('Arial', 9)).pack(side='left', padx=(0, 5))
        spinbox = ttk.Spinbox(font_frame, from_=10, to=100, increment=1, width=5, justify='right', wrap=True)
        spinbox.set(int(default_value_pt.pt)); spinbox.pack(side='left'); data_dict[key] = spinbox

    def _criar_widgets_musica_com_fonte(self, parent_frame, nome_parte, data_dict, titulo_sugerido):
        # (Função igual à v16 - OMITIDA)
        data_dict["tipo"] = "musica"; data_dict["titulo_geracao"] = titulo_sugerido
        refrao_frame = ttk.Frame(parent_frame); refrao_frame.pack(fill='x', expand=True)
        ttk.Label(refrao_frame, text="Refrão (Amarelo):", font=('Arial', 10, 'bold')).pack(pady=(5,2), anchor='w')
        self._criar_spinbox_fonte(refrao_frame, "Fonte:", DEFAULT_TAMANHO_FONTE_MUSICA_REFRAO, data_dict, "refrao_font_spinbox")
        data_dict["refrao_widget"] = scrolledtext.ScrolledText(refrao_frame, height=8, width=90, wrap=tk.WORD, font=('Arial', 10)); data_dict["refrao_widget"].pack(fill="x", expand=True, padx=5, pady=(0,5))
        verso_frame = ttk.Frame(parent_frame); verso_frame.pack(fill='x', expand=True)
        ttk.Label(verso_frame, text="Versos (Branco):", font=('Arial', 10, 'bold')).pack(pady=(10,2), anchor='w')
        self._criar_spinbox_fonte(verso_frame, "Fonte:", DEFAULT_TAMANHO_FONTE_MUSICA_VERSO, data_dict, "verso_font_spinbox")
        data_dict["verso_widget"] = scrolledtext.ScrolledText(verso_frame, height=12, width=90, wrap=tk.WORD, font=('Arial', 10)); data_dict["verso_widget"].pack(fill="x", expand=True, padx=5, pady=(0,5))

    # <<< WIDGET LEITURA CORRIGIDO >>>
    def _criar_widgets_leitura_com_fonte(self, parent_frame, nome_parte, data_dict, titulo_amarelo_padrao_list):
        data_dict["tipo"] = "leitura"
        # Não guarda mais o título fixo aqui, ele vem do DEFAULT_TEXTS na geração

        # Frame para Título/Referência Amarelo + Fonte
        ref_frame = ttk.Frame(parent_frame); ref_frame.pack(fill='x', expand=True)
        ttk.Label(ref_frame, text=f"Título/Referência - {nome_parte} (Amarelo): [Vazio = Padrão]", font=('Arial', 10, 'bold')).pack(pady=(5,2), anchor='w')
        self._criar_spinbox_fonte(ref_frame, "Fonte:", DEFAULT_TAMANHO_FONTE_LEITURA_TITULO_AMARELO, data_dict, "titulo_amarelo_font_spinbox") # Chave do spinbox
        data_dict["titulo_amarelo_widget"] = scrolledtext.ScrolledText(ref_frame, height=4, width=90, wrap=tk.WORD, font=('Arial', 10)); data_dict["titulo_amarelo_widget"].pack(fill="x", expand=True, padx=5, pady=(0,5))
        # Preenche widget com o padrão
        if titulo_amarelo_padrao_list: data_dict["titulo_amarelo_widget"].insert(tk.END, "\n".join(titulo_amarelo_padrao_list))

        # Frame para Texto Branco + Fonte
        texto_frame = ttk.Frame(parent_frame); texto_frame.pack(fill='both', expand=True)
        ttk.Label(texto_frame, text=f"Texto Principal - {nome_parte} (Branco): [Vazio = Padrão]", font=('Arial', 10, 'bold')).pack(pady=(10,2), anchor='w')
        self._criar_spinbox_fonte(texto_frame, "Fonte:", DEFAULT_TAMANHO_FONTE_LEITURA_TEXTO_BRANCO, data_dict, "texto_branco_font_spinbox") # Chave do spinbox
        data_dict["texto_branco_widget"] = scrolledtext.ScrolledText(texto_frame, height=18, width=90, wrap=tk.WORD, font=('Arial', 10)); data_dict["texto_branco_widget"].pack(fill="both", expand=True, padx=5, pady=(0,5))
        default_txt = DEFAULT_TEXTS.get(nome_parte, {}).get("texto_branco", []) # Usa chave correta
        if default_txt: data_dict["texto_branco_widget"].insert(tk.END, "\n".join(default_txt))

    def _criar_widgets_aclamacao_com_fonte(self, parent_frame, nome_parte, data_dict, titulo_sugerido):
        # (Função igual à v16 - OMITIDA)
        data_dict["tipo"] = "aclamacao"; data_dict["titulo_geracao"] = titulo_sugerido
        aclamacao_frame = ttk.Frame(parent_frame); aclamacao_frame.pack(fill='x', expand=True)
        ttk.Label(aclamacao_frame, text="Aclamação (Amarelo - Superior):", font=('Arial', 10, 'bold')).pack(pady=(5,2), anchor='w')
        self._criar_spinbox_fonte(aclamacao_frame, "Fonte:", DEFAULT_TAMANHO_FONTE_ACLAMACAO, data_dict, "aclamacao_font_spinbox")
        data_dict["aclamacao_widget"] = scrolledtext.ScrolledText(aclamacao_frame, height=6, width=90, wrap=tk.WORD, font=('Arial', 10)); data_dict["aclamacao_widget"].pack(fill="x", expand=True, padx=5, pady=(0,5))
        default_aclamacao = DEFAULT_TEXTS.get(nome_parte, {}).get("aclamacao_texto", []);
        if default_aclamacao: data_dict["aclamacao_widget"].insert(tk.END, "\n".join(default_aclamacao))
        antifona_frame = ttk.Frame(parent_frame); antifona_frame.pack(fill='x', expand=True)
        ttk.Label(antifona_frame, text="Antífona (Branco - Inferior):", font=('Arial', 10, 'bold')).pack(pady=(10,2), anchor='w')
        self._criar_spinbox_fonte(antifona_frame, "Fonte:", DEFAULT_TAMANHO_FONTE_ANTIFONA, data_dict, "antifona_font_spinbox")
        data_dict["antifona_widget"] = scrolledtext.ScrolledText(antifona_frame, height=15, width=90, wrap=tk.WORD, font=('Arial', 10)); data_dict["antifona_widget"].pack(fill="x", expand=True, padx=5, pady=(0,5))
        default_antifona = DEFAULT_TEXTS.get(nome_parte, {}).get("antifona_texto", []);
        if default_antifona: data_dict["antifona_widget"].insert(tk.END, "\n".join(default_antifona))


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

            # --- Funções Auxiliares ---
            def _get_font_size_from_spinbox(widget, default_pt_value):
                # Função auxiliar para obter tamanho de fonte do spinbox
                try:
                    valor_str = widget.get()
                    valor_int = int(valor_str)
                    if 10 <= valor_int <= 100:
                        return Pt(valor_int)
                    else:
                        print(f"Aviso: Fonte '{valor_str}' fora (10-100). Usando {default_pt_value.pt}pt.")
                        return default_pt_value
                except (tk.TclError, ValueError, AttributeError) as e:
                    print(f"Erro fonte: {e}. Usando {default_pt_value.pt}pt.")
                    return default_pt_value

            # Função adicionar_secao_musical (igual v16 - OMITIDA)
            def adicionar_secao_musical(nome_parte_gui):
                conteudo_adicionado_total = False; widget_existe = nome_parte_gui in self.widgets_gui and self.widgets_gui[nome_parte_gui]["tipo"] == "musica"; pegar_da_gui = widget_existe
                titulo_parte = DEFAULT_TEXTS.get(nome_parte_gui, {}).get("titulo", nome_parte_gui.upper()); refrao_gui_str = ""; versos_gui_str = ""
                tamanho_fonte_refrao = DEFAULT_TAMANHO_FONTE_MUSICA_REFRAO; tamanho_fonte_verso = DEFAULT_TAMANHO_FONTE_MUSICA_VERSO
                if pegar_da_gui: data_gui = self.widgets_gui[nome_parte_gui]; refrao_gui_str = data_gui["refrao_widget"].get("1.0", tk.END).strip(); versos_gui_str = data_gui["verso_widget"].get("1.0", tk.END).strip(); tamanho_fonte_refrao = _get_font_size_from_spinbox(data_gui.get("refrao_font_spinbox"), DEFAULT_TAMANHO_FONTE_MUSICA_REFRAO); tamanho_fonte_verso = _get_font_size_from_spinbox(data_gui.get("verso_font_spinbox"), DEFAULT_TAMANHO_FONTE_MUSICA_VERSO)
                else: tamanho_fonte_refrao = DEFAULT_TAMANHO_FONTE_MUSICA_REFRAO; tamanho_fonte_verso = DEFAULT_TAMANHO_FONTE_MUSICA_VERSO
                defaults = DEFAULT_TEXTS.get(nome_parte_gui, {}); refrao_padrao = defaults.get("refrao", []); versos_padrao = defaults.get("versos", [])
                refrao_final = [l.strip() for l in refrao_gui_str.split('\n') if l.strip()] if refrao_gui_str else refrao_padrao; versos_processados = []
                if versos_gui_str:
                    linhas_input = versos_gui_str.split('\n'); bloco_verso_atual = []
                    for linha in linhas_input:
                        linha_limpa = linha.strip()
                        if linha_limpa: bloco_verso_atual.append(linha_limpa)
                        elif bloco_verso_atual: versos_processados.append(bloco_verso_atual); bloco_verso_atual = []
                    if bloco_verso_atual: versos_processados.append(bloco_verso_atual)
                else: versos_processados = versos_padrao
                if versos_processados or refrao_final:
                    titulo_adicionado = adiciona_texto_com_divisao(prs, layout_slide_branco, [titulo_parte], COR_TITULO, TAMANHO_TITULO_PARTE, 5, use_auto_size=False);
                    if titulo_adicionado: conteudo_adicionado_total = True
                    for i, estrofe in enumerate(versos_processados):
                         verso_adicionado = adiciona_texto_com_divisao(prs, layout_slide_branco, estrofe, COR_VERSO, tamanho_fonte_verso, LINHAS_POR_SLIDE_VERSO, use_auto_size=True)
                         if verso_adicionado: conteudo_adicionado_total = True
                         if refrao_final:
                             refrao_adicionado = adiciona_texto_com_divisao(prs, layout_slide_branco, refrao_final, COR_REFRAO, tamanho_fonte_refrao, LINHAS_POR_SLIDE_VERSO, use_auto_size=True)
                             if refrao_adicionado: conteudo_adicionado_total = True
                    if not versos_processados and refrao_final:
                        refrao_adicionado = adiciona_texto_com_divisao(prs, layout_slide_branco, refrao_final, COR_REFRAO, tamanho_fonte_refrao, LINHAS_POR_SLIDE_VERSO, use_auto_size=True)
                        if refrao_adicionado: conteudo_adicionado_total = True
                return conteudo_adicionado_total


            # <<< FUNÇÃO LEITURA CORRIGIDA >>>
            def adicionar_leitura_slide_unico(nome_parte_gui):
                 conteudo_adicionado = False
                 if nome_parte_gui in self.widgets_gui and self.widgets_gui[nome_parte_gui]["tipo"] == "leitura": # Verifica o tipo correto
                     data_gui = self.widgets_gui[nome_parte_gui]
                     # Pega textos da GUI
                     titulo_amarelo_gui_str = data_gui["titulo_amarelo_widget"].get("1.0", tk.END).strip() # Nome correto do widget
                     texto_branco_gui_str = data_gui["texto_branco_widget"].get("1.0", tk.END).strip() # Nome correto do widget
                     # Pega tamanhos das fontes
                     tamanho_fonte_titulo_amarelo = _get_font_size_from_spinbox(data_gui.get("titulo_amarelo_font_spinbox"), DEFAULT_TAMANHO_FONTE_LEITURA_TITULO_AMARELO) # Chave correta
                     tamanho_fonte_texto_branco = _get_font_size_from_spinbox(data_gui.get("texto_branco_font_spinbox"), DEFAULT_TAMANHO_FONTE_LEITURA_TEXTO_BRANCO) # Chave correta

                     defaults = DEFAULT_TEXTS.get(nome_parte_gui, {})
                     titulo_amarelo_padrao = defaults.get("titulo_amarelo", []) # Chave correta
                     texto_branco_padrao = defaults.get("texto_branco", []) # Chave correta
                     # Decide finais
                     titulo_amarelo_final = [l.strip() for l in titulo_amarelo_gui_str.split('\n') if l.strip()] if titulo_amarelo_gui_str else titulo_amarelo_padrao
                     texto_branco_final = [l.strip() for l in texto_branco_gui_str.split('\n') if l.strip()] if texto_branco_gui_str else texto_branco_padrao

                     # Cria slide único para Título Amarelo + Texto Branco (se houver um dos dois)
                     if titulo_amarelo_final or texto_branco_final:
                         slide = prs.slides.add_slide(layout_slide_branco)
                         conteudo_adicionado = True

                         esquerda = MARGEM_TEXTO; topo = MARGEM_TEXTO; largura = LARGURA_SLIDE - (2 * MARGEM_TEXTO); altura = ALTURA_SLIDE - (2 * MARGEM_TEXTO)
                         caixa_texto = slide.shapes.add_textbox(esquerda, topo, largura, altura)
                         frame_texto = caixa_texto.text_frame; frame_texto.clear(); frame_texto.word_wrap = True
                         frame_texto.vertical_anchor = MSO_ANCHOR.MIDDLE; frame_texto.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

                         # 1. Adiciona Título/Referência (Amarelo)
                         if titulo_amarelo_final:
                             p_titulo = frame_texto.add_paragraph()
                             texto_titulo_continuo = " ".join(titulo_amarelo_final)
                             p_titulo.text = texto_titulo_continuo; p_titulo.alignment = PP_ALIGN.CENTER; p_titulo.font.name = NOME_FONTE
                             p_titulo.font.size = tamanho_fonte_titulo_amarelo # Usa tamanho da GUI/Default
                             p_titulo.font.color.rgb = COR_REFRAO; p_titulo.font.bold = True # Amarelo

                         # 2. Adiciona Texto Principal (Branco)
                         if texto_branco_final:
                              # Adiciona espaço se ambos existirem
                             if titulo_amarelo_final: p_espaco = frame_texto.add_paragraph(); p_espaco.text = ""
                             p_texto = frame_texto.add_paragraph()
                             texto_principal_continuo = " ".join(texto_branco_final)
                             p_texto.text = texto_principal_continuo; p_texto.alignment = PP_ALIGN.CENTER; p_texto.font.name = NOME_FONTE
                             p_texto.font.size = tamanho_fonte_texto_branco # Usa tamanho da GUI/Default
                             p_texto.font.color.rgb = COR_VERSO; p_texto.font.bold = True # Branco

                         try: caixa_texto.left = esquerda; caixa_texto.top = topo; caixa_texto.width = largura; caixa_texto.height = altura; frame_texto.margin_bottom = Inches(0.05); frame_texto.margin_top = Inches(0.05); frame_texto.margin_left = Inches(0.1); frame_texto.margin_right = Inches(0.1)
                         except Exception as e_resize: print(f"Aviso: resize caixa leitura: {e_resize}")

                 # Não adiciona slide de título de seção separado
                 return conteudo_adicionado


            # Função adicionar_aclamacao_slide_unico (igual v15 - OMITIDA)
            def adicionar_aclamacao_slide_unico(nome_parte_gui):
                conteudo_adicionado = False
                if nome_parte_gui in self.widgets_gui and self.widgets_gui[nome_parte_gui]["tipo"] == "aclamacao":
                    data_gui = self.widgets_gui[nome_parte_gui]; titulo_secao = data_gui["titulo_geracao"]
                    aclamacao_gui_str = data_gui["aclamacao_widget"].get("1.0", tk.END).strip(); antifona_gui_str = data_gui["antifona_widget"].get("1.0", tk.END).strip()
                    tamanho_fonte_aclamacao = _get_font_size_from_spinbox(data_gui.get("aclamacao_font_spinbox"), DEFAULT_TAMANHO_FONTE_ACLAMACAO); tamanho_fonte_antifona = _get_font_size_from_spinbox(data_gui.get("antifona_font_spinbox"), DEFAULT_TAMANHO_FONTE_ANTIFONA)
                    defaults = DEFAULT_TEXTS.get(nome_parte_gui, {}); aclamacao_padrao = defaults.get("aclamacao_texto", []); antifona_padrao = defaults.get("antifona_texto", [])
                    aclamacao_final = [l.strip() for l in aclamacao_gui_str.split('\n') if l.strip()] if aclamacao_gui_str else aclamacao_padrao; antifona_final = [l.strip() for l in antifona_gui_str.split('\n') if l.strip()] if antifona_gui_str else antifona_padrao
                    titulo_adicionado = adiciona_texto_com_divisao(prs, layout_slide_branco, [titulo_secao], COR_TITULO, TAMANHO_TITULO_PARTE, 5, use_auto_size=False)
                    if titulo_adicionado: conteudo_adicionado = True
                    if aclamacao_final or antifona_final:
                        slide = prs.slides.add_slide(layout_slide_branco); conteudo_adicionado = True
                        esquerda = MARGEM_TEXTO; topo = MARGEM_TEXTO; largura = LARGURA_SLIDE - (2 * MARGEM_TEXTO); altura = ALTURA_SLIDE - (2 * MARGEM_TEXTO)
                        caixa_texto = slide.shapes.add_textbox(esquerda, topo, largura, altura); frame_texto = caixa_texto.text_frame; frame_texto.clear(); frame_texto.word_wrap = True
                        frame_texto.vertical_anchor = MSO_ANCHOR.MIDDLE; frame_texto.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
                        if aclamacao_final:
                            p_aclamacao = frame_texto.add_paragraph(); texto_aclamacao_continuo = " ".join(aclamacao_final); p_aclamacao.text = texto_aclamacao_continuo; p_aclamacao.alignment = PP_ALIGN.CENTER; p_aclamacao.font.name = NOME_FONTE
                            p_aclamacao.font.size = tamanho_fonte_aclamacao; p_aclamacao.font.color.rgb = COR_REFRAO; p_aclamacao.font.bold = True
                        if antifona_final:
                            if aclamacao_final: p_espaco_ac = frame_texto.add_paragraph(); p_espaco_ac.text = ""
                            p_antifona = frame_texto.add_paragraph(); texto_antifona_continuo = " ".join(antifona_final); p_antifona.text = texto_antifona_continuo; p_antifona.alignment = PP_ALIGN.CENTER; p_antifona.font.name = NOME_FONTE
                            p_antifona.font.size = tamanho_fonte_antifona; p_antifona.font.color.rgb = COR_VERSO; p_antifona.font.bold = True
                        try: caixa_texto.left = esquerda; caixa_texto.top = topo; caixa_texto.width = largura; caixa_texto.height = altura; frame_texto.margin_bottom = Inches(0.05); frame_texto.margin_top = Inches(0.05); frame_texto.margin_left = Inches(0.1); frame_texto.margin_right = Inches(0.1)
                        except Exception as e_resize: print(f"Aviso: resize caixa aclamacao: {e_resize}")
                return conteudo_adicionado

            # Função adicionar_secao_fixa (igual v13_fix - OMITIDA)
            def adicionar_secao_fixa(titulo_secao, texto_linhas, tamanho_fonte, linhas_por_slide, cor=COR_VERSO, add_separador=True, bold_content=True, use_auto_size_content=False):
                 titulo_adicionado = False; conteudo_adicionado = False
                 if titulo_secao and titulo_secao.strip(): titulo_adicionado = adiciona_texto_com_divisao(prs, layout_slide_branco, [titulo_secao], COR_TITULO, TAMANHO_TITULO_PARTE, 5, bold=True, use_auto_size=False)
                 if texto_linhas: conteudo_adicionado = adiciona_texto_com_divisao(prs, layout_slide_branco, texto_linhas, cor, tamanho_fonte, linhas_por_slide, bold=bold_content, use_auto_size=use_auto_size_content)
                 if (titulo_adicionado or conteudo_adicionado) and add_separador: prs.slides.add_slide(layout_slide_branco); print(f"Separador após {titulo_secao}")
                 return titulo_adicionado or conteudo_adicionado


            # --- Montagem da Apresentação ---
            # (Lógica de montagem igual à v15)
            ordem_final_geracao = [ "Entrada", "Ato Penitencial", "PALAVRA_INTRO", "1ª Leitura", "Salmo", "2ª Leitura", "Aclamação", "CREDO", "PRECES", "Oferendas", "SANTO_TITULO", "ORACAO_EUCARISTICA", "CORDEIRO_TITULO", "Comunhão", "Pós-Comunhão", "SANTA_LUZIA", "AVISOS", "Final" ]
            initial_title_str = self.initial_title_widget.get("1.0", tk.END).strip(); initial_title_lines = [l.strip() for l in initial_title_str.split('\n') if l.strip()]
            if initial_title_lines:
                 titulo_inicial_adicionado = adiciona_texto_com_divisao(prs, layout_slide_branco, initial_title_lines, COR_TITULO, TAMANHO_FONTE_TITULO_INICIAL, 5, use_auto_size=True)
                 if titulo_inicial_adicionado: prs.slides.add_slide(layout_slide_branco)
            for nome_parte in ordem_final_geracao:
                separador_necessario = False
                if nome_parte == "PALAVRA_INTRO": separador_necessario = adicionar_secao_fixa("PALAVRA", TEXTO_PALAVRA_INTRO, Pt(90), 6, cor=COR_TITULO, add_separador=False)
                elif nome_parte == "CREDO": separador_necessario = adicionar_secao_fixa("ORAÇÃO DO CREDO", TEXTO_CREDO, Pt(83), 4,use_auto_size_content=True)
                elif nome_parte == "PRECES": separador_necessario = adicionar_secao_fixa("PRECES", [], TAMANHO_TITULO_PARTE, 1,),
                elif nome_parte == "ORACAO_EUCARISTICA": separador_necessario = adicionar_secao_fixa("ORAÇÃO EUCARÍSTICA", [], TAMANHO_TITULO_PARTE, 2)
                elif nome_parte == "SANTA_LUZIA": separador_necessario = adicionar_secao_fixa("ORAÇÃO A SANTA LUZIA", TEXTO_ORACAO_SANTA_LUZIA, Pt(80), LINHAS_POR_SLIDE_ORACAO,use_auto_size_content=True)
                elif nome_parte == "AVISOS": separador_necessario = adicionar_secao_fixa("AVISOS", TEXTO_AVISOS, Pt(90), 4, add_separador=False, bold_content=False, use_auto_size_content=True)
                elif nome_parte in ["1ª Leitura", "Salmo", "2ª Leitura"]:
                    separador_necessario = adicionar_leitura_slide_unico(nome_parte)
                    # <<< Lógica de separador APÓS CADA LEITURA >>>
                    if separador_necessario:
                       prs.slides.add_slide(layout_slide_branco)
                       print(f"Separador adicionado após {nome_parte}")
                elif nome_parte == "Aclamação":
                    separador_necessario = adicionar_aclamacao_slide_unico(nome_parte)
                    if separador_necessario: prs.slides.add_slide(layout_slide_branco)
                elif nome_parte in self.widgets_gui or nome_parte in ["Pós-Comunhão", "Final"]:
                    separador_necessario = adicionar_secao_musical(nome_parte)
                    if separador_necessario and nome_parte != ordem_final_geracao[-1]: prs.slides.add_slide(layout_slide_branco)

            # --- Salvar ---
            # (Lógica de salvar igual)
            filepath = filedialog.asksaveasfilename( defaultextension=".pptx", filetypes=[("PowerPoint Presentations", "*.pptx"), ("All Files", "*.*")], title="Salvar Apresentação Como...", initialfile="Missa_Gerada_v19.pptx" )
            if not filepath: self.status_label.config(text="Geração cancelada."); return
            prs.save(filepath)
            self.status_label.config(text=f"Salvo: {os.path.basename(filepath)}")
            messagebox.showinfo("Sucesso", f"Apresentação '{os.path.basename(filepath)}' gerada com sucesso!")

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
    if 'Entrada' not in DEFAULT_TEXTS or '1ª Leitura' not in DEFAULT_TEXTS or 'Aclamação' not in DEFAULT_TEXTS or not TEXTO_CREDO or not TEXTO_ORACAO_SANTA_LUZIA: print("ERRO CRÍTICO: Dicionários de texto padrão não estão completos!"); exit()
    root = tk.Tk()
    app = MassSlideGeneratorApp(root)
    root.mainloop()


 