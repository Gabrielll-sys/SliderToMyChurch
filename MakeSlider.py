import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
# Importa MSO_AUTO_SIZE para ajuste automático
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN, MSO_AUTO_SIZE

# --- Constantes de Configuração ---
LARGURA_SLIDE = Inches(16)
ALTURA_SLIDE = Inches(9)
MARGEM_TEXTO = Inches(0) # Margens laterais e superior/inferior
TAMANHO_FONTE_MUSICA = Pt(72) # Tamanho inicial GRANDE
TAMANHO_TITULO_PARTE = Pt(60)
TAMANHO_FONTE_PADRAO = Pt(44)
TAMANHO_FONTE_LEITURA = Pt(40)
TAMANHO_FONTE_ORACAO = Pt(76)
LINHAS_POR_SLIDE_VERSO = 7 # Limite de linhas *originais* antes de forçar divisão
LINHAS_POR_SLIDE_ORACAO = 7
NOME_FONTE = 'Arial'
COR_REFRAO = RGBColor(255, 192, 0)
COR_VERSO = RGBColor(255, 255, 255)
COR_TITULO = RGBColor(255, 192, 0)
COR_FUNDO_PRETO = RGBColor(0, 0, 0)

# --- Textos Padrão para Músicas (Fallback da GUI) ---
DEFAULT_TEXTS = {
    # ===========================================================
    # COLE AQUI O DICIONÁRIO 'DEFAULT_TEXTS' COMPLETO
    # DA RESPOSTA ANTERIOR.
     "Entrada": {
        "titulo": "CANTO DE ENTRADA",
        "refrao": ["SENHOR, EIS AQUI O TEU","POVO QUE VEM IMPLORAR","TEU PERDÃO","É GRANDE O NOSSO","PECADO, PORÉM É MAIOR O","TEU CORAÇÃO"],
        "versos": [
            ["SABENDO QUE","ACOLHESTE ZAQUEU, O","COBRADOR E ASSIM LHE","DEVOLVESTE TUA PAZ E","TEU AMOR TAMBÉM"],
            ["NOS COLOCAMOS AO","LADO DOS QUE VÃO","BUSCAR NO TEU ALTAR A","GRAÇA DO PERDÃO"],
            ["REVENDO EM MADALENA","A NOSSA PRÓPRIA FÉ","CHORANDO NOSSAS","PENAS DIANTE DOS TEUS","PÉS TAMBÉM"],
            ["NÓS DESEJAMOS O","NOSSO AMOR TE DAR","PORQUE SÓ MUITO","AMOR NOS PODE","LIBERTAR"],
            ["MOTIVOS TEMOS NÓS","DE SEMPRE CONFIAR,","DE ERGUER A NOSSA VOZ,","DE NÃO DESESPERAR,","OLHANDO AQUELE GESTO"],
            ["QUE O BOM LADRÃO","SALVOU,","NÃO FOI, TAMBÉM, POR","NÓS,","TEU SANGUE QUE JORROU?"]
        ]
     },
     "Ato Penitencial": {"titulo": "ATO PENITENCIAL", "refrao": [], "versos": []},
     "Aclamação": {"titulo": "ACLAMAÇÃO DO EVANGELHO", "refrao": ["Glória a Vós o Cristo, verbo de Deus! (bis)"], "versos": [["Vou levantar-me e vou a meu pai","e lhe direi: Meu pai, eu pequei","contra o céu e contra ti.","(Lc 15,1-3.11-32)"]]},
     "Oferendas": {"titulo": "PREPARAÇÃO DAS OFERENDAS", "refrao": ["CONFIEI NO TEU AMOR E","VOLTEI, SIM, AQUI É MEU","LUGAR. EU GASTEI TEUS","BENS, Ó PAI, E TE DOU","ESTE PRANTO EM MINHAS","MÃOS"], "versos": [["MUITO ALEGRE EU TE","PEDI O QUE ERA MEU","PARTI, UM SONHO TÃO","NORMAL"],["DISSIPEI MEUS BENS E","O CORAÇÃO TAMBÉM","NO FIM, MEU MUNDO","ERA IRREAL"],["MIL AMIGOS CONHECI,","DISSERAM ADEUS","CAIU A SOLIDÃO EM","MIM"],["UM PATRÃO CRUEL","LEVOU-ME A REFLETIR","MEU PAI NÃO TRATA","UM SERVO ASSIM"],["NEM DEIXASTE-ME","FALAR DA INGRATIDÃO","MORREU NO ABRAÇO","O MAL QUE EU FIZ"],["FESTA, ROUPA NOVA,","ANEL, SANDÁLIA AOS","PÉS","VOLTEI À VIDA, SOU","FELIZ"]]},
     "Santo": {"titulo": "SANTO", "refrao": [], "versos": []},
     "Cordeiro": {"titulo": "CORDEIRO", "refrao": [], "versos": []},
     "Comunhão": {"titulo": "COMUNHÃO", "refrao": ["PROVAI E VEDE COMO DEUS É","BOM FELIZ DE QUEM NO SEU","AMOR CONFIA EM JESUS","CRISTO, SE FAZ GRAÇA E DOM","SE FAZ PALAVRA E PÃO ΝΑ","EUCARISTIA"], "versos": [["Ó PAI, TEU POVO BUSCA VIDA","NOVA NA DIREÇÃO DA PÁSCOA","DE JESUS EM NOSSA FRONTE, O","SINAL DAS CINZAS NA","CAMINHADA,","VEM SER FORÇA E LUZ"],["QUANDO, NA VIDA, ANDAMOS","NO DESERTO E A TENTAÇÃO","VEM NOS TIRAR A PAZ A","FORTALEZA E A PALAVRA","CERTA EM TI BUSCAMOS, DEUS","DE NOSSOS PAIS"],["PEREGRINAMOS ENTRE LUZ E","SOMBRAS A CRUZ NOS PESA, O","MAL NOS DESFIGURA MAS NA","ORAÇÃO E NA PALAVRA","ACHAMOS A TUA GRAÇA, QUE","NOS TRANSFIGURA"],["Ó DEUS, CONHECES NOSSO","SOFRIMENTO HÁ MUITA DOR, É","GRANDE A AFLIÇÃO","TRANSFORMA EM FESTA NOSSA","DOR-LAMENTO ACOLHE OS","FRUTOS BONS DA CONVERSÃO"],["QUANDO O PECADO NOS","CONSOME E FERE E EM TI","BUSCAMOS A PAZ DO PERDÃO","O NOSSO RIO DE AFLIÇÃO SE","PERDE NO MAR PROFUNDO DO","TEU CORAÇÃO"],["POR QUE FICAR EM COISAS JÁ","PASSADAS? O TEU PERDÃO","LIBERTA E NOS RENOVA O TEU","AMOR NOS ABRE NOVA","ESTRADA TRAZ ALEGRIA E PAZ,","NOS REVIGORA"]]},
     "Pós-Comunhão": {"titulo": "CANTO PÓS-COMUNHÃO", "refrao": [], "versos": []},
     "Final": {"titulo": "CANTO FINAL", "refrao": [], "versos": []},
    # ===========================================================
}

# --- Textos Fixos ---
# (Cole os textos fixos completos: TITULO_INICIAL, PALAVRA_INTRO, LEITURAS, CREDO, SANTA_LUZIA, AVISOS)
TEXTO_TITULO_INICIAL = ["4º DOMINGO DA", "QUARESMA"]
TEXTO_PALAVRA_INTRO = ["DESÇA COMO A CHUVA A TUA","PALAVRA. QUE SE ESPALHE","COMO ORVALHO. COMO","CHUVISCO NA RELVA. COMO","AGUACEIRO NA GRAMA.","AMÉM!"]
TEXTO_LEITURAS = [ ("PRIMEIRA LEITURA", ["Josue 5,9 a.10-12"]), ("SALMO 33 (34)", ["Provai e vede quão suave", "é o Senhor!"]), ("SEGUNDA LEITURA", ["2Corintíos 5,17-21"]) ]
TEXTO_CREDO = [ "CREIO EM DEUS PAI TODO PODEROSO,", "CRIADOR DO CÉU E DA TERRA.",
               "E EM JESUS CRISTO, SEU ÚNICO FILHO,", "NOSSO SENHOR,",
               "QUE FOI CONCEBIDO PELO PODER DO ESPÍRITO SANTO;", "NASCEU DA VIRGEM MARIA;",
               "PADECEU SOB PÔNCIO PILATOS,", "FOI CRUCIFICADO, MORTO E SEPULTADO.",
               "DESCEU À MANSÃO DOS MORTOS;", "RESSUSCITOU AO TERCEIRO DIA;",
               "SUBIU AOS CÉUS, ESTÁ SENTADO À DIREITA", "DE DEUS PAI TODO-PODEROSO,",
               "DONDE HÁ DE VIR A JULGAR OS VIVOS E OS MORTOS.",
               "CREIO NO ESPÍRITO SANTO,", "NA SANTA IGREJA CATÓLICA,",
               "NA COMUNHÃO DOS SANTOS,", "NA REMISSÃO DOS PECADOS,",
               "NA RESSURREIÇÃO DA CARNE,", "NA VIDA ETERNA.", "AMÉM." ]
TEXTO_ORACAO_SANTA_LUZIA = [ "Ó VIRGEM ADMIRÁVEL.", "CHEIA DE FIRMEZA E DE", "CONSTÂNCIA, QUE NEM",
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

# --- Função Auxiliar Atualizada ---
def adiciona_texto_com_divisao(prs, layout, linhas_originais, cor, tamanho_fonte, max_linhas, bold=True, use_auto_size=True):
    """
    Adiciona texto, dividindo baseado em max_linhas originais,
    mas permite que o PowerPoint faça o wrap e use auto_size.
    """
    if not linhas_originais or all(not s or s.isspace() for s in linhas_originais):
        return False
    linhas_validas = [linha for linha in linhas_originais if linha and not linha.isspace()]
    if not linhas_validas:
        return False

    linhas_restantes = linhas_validas[:]
    slides_criados = 0

    while linhas_restantes:
        # Pega o bloco de linhas originais para este slide
        linhas_para_este_bloco = linhas_restantes[:max_linhas]
        linhas_restantes = linhas_restantes[max_linhas:]

        # Junta as linhas do bloco atual em uma única string com espaços
        texto_bloco_continuo = " ".join(linhas_para_este_bloco)

        if not texto_bloco_continuo.strip(): # Pula se o bloco ficou vazio
            continue

        slide = prs.slides.add_slide(layout)
        slides_criados += 1

        esquerda = MARGEM_TEXTO
        topo = MARGEM_TEXTO
        largura = LARGURA_SLIDE - (2 * MARGEM_TEXTO)
        altura = ALTURA_SLIDE - (2 * MARGEM_TEXTO)

        caixa_texto = slide.shapes.add_textbox(esquerda, topo, largura, altura)
        frame_texto = caixa_texto.text_frame
        frame_texto.clear()
        frame_texto.word_wrap = True # Essencial para o wrap automático
        frame_texto.vertical_anchor = MSO_ANCHOR.MIDDLE

        # Configura auto_size ANTES de adicionar texto
        if use_auto_size:
             # Tenta ajustar o texto à forma (pode reduzir fonte)
             frame_texto.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        else:
            # Se não usar auto_size, o texto pode estourar a caixa
            frame_texto.auto_size = MSO_AUTO_SIZE.NONE


        p = frame_texto.add_paragraph()
        # Adiciona o texto contínuo do bloco
        p.text = texto_bloco_continuo
        p.alignment = PP_ALIGN.CENTER # Centraliza
        p.font.name = NOME_FONTE
        # Define o tamanho inicial da fonte (auto_size pode reduzir)
        p.font.size = tamanho_fonte
        p.font.color.rgb = cor
        p.font.bold = bold

    return slides_criados > 0

# --- Classe Principal da Aplicação GUI (sem alterações na estrutura) ---
class MassSlideGeneratorApp:
    def __init__(self, master):
        # (Código __init__ e _criar_widgets_parte igual à v4/v5 - OMITIDO)
        self.master = master
        master.title("Gerador de Slides para Missa v6 (Wrap + AutoSize)")
        master.geometry("800x650")
        self.notebook = ttk.Notebook(master)

        self.partes_musicais_gui = { "Entrada": {}, "Aclamação": {}, "Oferendas": {}, "Santo": {}, "Cordeiro": {}, "Comunhão": {}, "Pós-Comunhão": {}, "Final": {} }
        for nome_parte in self.partes_musicais_gui.keys():
            frame = ttk.Frame(self.notebook, padding="10")
            self.notebook.add(frame, text=nome_parte)
            titulo_default = DEFAULT_TEXTS.get(nome_parte, {}).get("titulo", nome_parte.upper())
            self._criar_widgets_parte(frame, nome_parte, self.partes_musicais_gui[nome_parte], titulo_default)

        self.notebook.pack(expand=True, fill="both", padx=10, pady=5)
        bottom_frame = ttk.Frame(master, padding="5")
        bottom_frame.pack(fill="x", side="bottom", pady=(0, 5))
        self.status_label = ttk.Label(bottom_frame, text="Pronto. Fundo preto. Tentará auto-ajuste.")
        self.status_label.pack(side="left", padx=10)
        self.generate_button = ttk.Button(bottom_frame, text="Gerar PowerPoint", command=self.gerar_apresentacao)
        self.generate_button.pack(side="right", padx=10)

    def _criar_widgets_parte(self, parent_frame, nome_parte, data_dict, titulo_sugerido):
        # (Função igual à v4/v5)
        data_dict["titulo"] = titulo_sugerido
        ttk.Label(parent_frame, text=f"{nome_parte} - Versos (Branco): [Vazio = Padrão]", font=('Arial', 10, 'bold')).pack(pady=(5,2), anchor='w')
        data_dict["verso_widget"] = scrolledtext.ScrolledText(parent_frame, height=6, width=80, wrap=tk.WORD, font=('Arial', 10))
        data_dict["verso_widget"].pack(fill="x", expand=True)
        ttk.Label(parent_frame, text=f"{nome_parte} - Refrão (Amarelo): [Vazio = Padrão]", font=('Arial', 10, 'bold')).pack(pady=(10,2), anchor='w')
        data_dict["refrao_widget"] = scrolledtext.ScrolledText(parent_frame, height=4, width=80, wrap=tk.WORD, font=('Arial', 10))
        data_dict["refrao_widget"].pack(fill="x", expand=True)


    def gerar_apresentacao(self):
        # (Verificação inicial dos dicionários - igual v5)
        if not DEFAULT_TEXTS or not TEXTO_CREDO or not TEXTO_ORACAO_SANTA_LUZIA:
             messagebox.showerror("Erro de Configuração", "Os dicionários de texto padrão não foram preenchidos no código.")
             return

        self.status_label.config(text="Gerando apresentação...")
        self.master.update_idletasks()

        try:
            prs = Presentation()
            prs.slide_width = LARGURA_SLIDE
            prs.slide_height = ALTURA_SLIDE
            slide_master = prs.slide_masters[0]
            background = slide_master.background
            fill = background.fill
            fill.solid()
            fill.fore_color.rgb = COR_FUNDO_PRETO
            layout_slide_branco = next((layout for i, layout in enumerate(prs.slide_layouts) if "Branco" in layout.name or "Blank" in layout.name), prs.slide_layouts[5 if len(prs.slide_layouts) > 5 else 0])

            # --- Funções Auxiliares ATUALIZADAS ---
            def adicionar_secao_musical(nome_parte_gui):
                conteudo_adicionado_total = False
                # (Lógica para pegar texto da GUI ou Padrão - igual v5)
                if nome_parte_gui in self.partes_musicais_gui:
                    data_gui = self.partes_musicais_gui[nome_parte_gui]
                    titulo_parte = data_gui["titulo"]
                    versos_gui_str = data_gui["verso_widget"].get("1.0", tk.END).strip()
                    refrao_gui_str = data_gui["refrao_widget"].get("1.0", tk.END).strip()
                    defaults = DEFAULT_TEXTS.get(nome_parte_gui, {})
                    versos_padrao = defaults.get("versos", [])
                    refrao_padrao = defaults.get("refrao", [])
                    refrao_final = [linha.strip() for linha in refrao_gui_str.split('\n') if linha.strip()] if refrao_gui_str else refrao_padrao
                    versos_para_slides = []
                    if versos_gui_str:
                        versos_para_slides = [[linha.strip() for linha in versos_gui_str.split('\n') if linha.strip()]]
                    else:
                        versos_para_slides = versos_padrao

                    if versos_para_slides or refrao_final:
                        # Adiciona Título (sem auto_size geralmente)
                        titulo_adicionado = adiciona_texto_com_divisao(prs, layout_slide_branco, [titulo_parte], COR_TITULO, TAMANHO_TITULO_PARTE, 5, use_auto_size=False)
                        if titulo_adicionado: conteudo_adicionado_total = True

                        # Adiciona Versos (usando auto_size)
                        for bloco_verso in versos_para_slides:
                             verso_adicionado = adiciona_texto_com_divisao(prs, layout_slide_branco, bloco_verso, COR_VERSO, TAMANHO_FONTE_MUSICA, LINHAS_POR_SLIDE_VERSO, use_auto_size=True)
                             if verso_adicionado: conteudo_adicionado_total = True
                             # Adiciona Refrão após cada bloco (usando auto_size)
                             if refrao_final:
                                 refrao_adicionado = adiciona_texto_com_divisao(prs, layout_slide_branco, refrao_final, COR_REFRAO, TAMANHO_FONTE_MUSICA, LINHAS_POR_SLIDE_VERSO, use_auto_size=True)
                                 if refrao_adicionado: conteudo_adicionado_total = True

                        if not versos_para_slides and refrao_final: # Só refrão
                             refrao_adicionado = adiciona_texto_com_divisao(prs, layout_slide_branco, refrao_final, COR_REFRAO, TAMANHO_FONTE_MUSICA, LINHAS_POR_SLIDE_VERSO, use_auto_size=True)
                             if refrao_adicionado: conteudo_adicionado_total = True

                return conteudo_adicionado_total

            def adicionar_secao_fixa(titulo_secao, texto_linhas, tamanho_fonte, linhas_por_slide, cor=COR_VERSO, add_separador=True, bold_content=True, use_auto_size_content=False):
                conteudo_adicionado_total = False
                # Adiciona Título (sem auto_size)
                titulo_adicionado = adiciona_texto_com_divisao(prs, layout_slide_branco, [titulo_secao], COR_TITULO, TAMANHO_TITULO_PARTE, 5, use_auto_size=False)
                if titulo_adicionado: conteudo_adicionado_total = True

                # Adiciona Conteúdo (com opção de auto_size)
                conteudo_adicionado = adiciona_texto_com_divisao(prs, layout_slide_branco, texto_linhas, cor, tamanho_fonte, linhas_por_slide, bold=bold_content, use_auto_size=use_auto_size_content)
                if conteudo_adicionado: conteudo_adicionado_total = True

                if conteudo_adicionado_total and add_separador:
                    prs.slides.add_slide(layout_slide_branco)
                return conteudo_adicionado_total

            # --- Montagem da Apresentação ---
            # (Ordem e chamadas igual à v5, mas ajustando `use_auto_size_content` onde necessário)

            adicionar_secao_fixa("TITULO_INICIAL_DUMMY", TEXTO_TITULO_INICIAL, TAMANHO_TITULO_PARTE, 5, cor=COR_TITULO) # Título não precisa auto_size

            if adicionar_secao_musical("Entrada"): prs.slides.add_slide(layout_slide_branco)
            if adicionar_secao_musical("Ato Penitencial"): prs.slides.add_slide(layout_slide_branco)
            # Glória
            # Palavra
            palavra_adicionada = adicionar_secao_fixa("PALAVRA", TEXTO_PALAVRA_INTRO, TAMANHO_FONTE_PADRAO, 6, cor=COR_TITULO, add_separador=False, use_auto_size_content=True) # Permitir auto-size aqui
            leitura_adicionada = False
            for titulo_l, texto_l in TEXTO_LEITURAS:
                 slide_l_titulo = prs.slides.add_slide(layout_slide_branco)
                 adiciona_texto_com_divisao(prs, layout_slide_branco, [titulo_l], COR_TITULO, TAMANHO_FONTE_LEITURA + Pt(4), 3, use_auto_size=False)
                 adiciona_texto_com_divisao(prs, layout_slide_branco, texto_l, COR_VERSO, TAMANHO_FONTE_LEITURA, 4, bold=False, use_auto_size=True) # Permitir auto-size
                 leitura_adicionada = True
            if palavra_adicionada or leitura_adicionada: prs.slides.add_slide(layout_slide_branco)

            if adicionar_secao_musical("Aclamação"): prs.slides.add_slide(layout_slide_branco)
            adicionar_secao_fixa("ORAÇÃO DO CREDO", TEXTO_CREDO, TAMANHO_FONTE_ORACAO, LINHAS_POR_SLIDE_ORACAO, use_auto_size_content=True) # Permitir auto-size
            # Preces
            if adicionar_secao_fixa("PRECES", [], TAMANHO_TITULO_PARTE, 1): prs.slides.add_slide(layout_slide_branco)

            if adicionar_secao_musical("Oferendas"): prs.slides.add_slide(layout_slide_branco)
            if adicionar_secao_musical("Santo"): prs.slides.add_slide(layout_slide_branco)
            # Oração Eucarística
            if adicionar_secao_fixa("ORAÇÃO EUCARÍSTICA", [], TAMANHO_TITULO_PARTE, 2): prs.slides.add_slide(layout_slide_branco)

            if adicionar_secao_musical("Cordeiro"): prs.slides.add_slide(layout_slide_branco)
            if adicionar_secao_musical("Comunhão"): prs.slides.add_slide(layout_slide_branco)
            if adicionar_secao_musical("Pós-Comunhão"): prs.slides.add_slide(layout_slide_branco)
            adicionar_secao_fixa("ORAÇÃO A SANTA LUZIA", TEXTO_ORACAO_SANTA_LUZIA, TAMANHO_FONTE_ORACAO, LINHAS_POR_SLIDE_ORACAO, use_auto_size_content=True) # Permitir auto-size
            adicionar_secao_fixa("AVISOS", TEXTO_AVISOS, TAMANHO_FONTE_PADRAO, 5, add_separador=False, bold_content=False, use_auto_size_content=True) # Permitir auto-size

            # --- Salvar ---
            # (Lógica de salvar igual à v5)
            filepath = filedialog.asksaveasfilename( defaultextension=".pptx", filetypes=[("PowerPoint Presentations", "*.pptx"), ("All Files", "*.*")], title="Salvar Apresentação Como...", initialfile="Missa_Gerada_v6.pptx" )
            if not filepath: self.status_label.config(text="Geração cancelada."); return
            prs.save(filepath)
            self.status_label.config(text=f"Salvo: {os.path.basename(filepath)}")
            messagebox.showinfo("Sucesso", f"Apresentação '{os.path.basename(filepath)}' gerada com sucesso!")

        except Exception as e:
            # (Lógica de erro igual à v5)
            self.status_label.config(text="Erro durante a geração!")
            print(f"Erro detalhado: {e}")
            import traceback; traceback.print_exc()
            messagebox.showerror("Erro", f"Ocorreu um erro:\n{e}\nVerifique o console.")
        finally:
            self.master.update_idletasks()


# --- Iniciar a Aplicação ---
if __name__ == "__main__":
    # (Colar dicionários completos aqui)
    if 'Entrada' not in DEFAULT_TEXTS or not TEXTO_CREDO or not TEXTO_ORACAO_SANTA_LUZIA: print("ERRO CRÍTICO: Dicionários de texto padrão não estão completos!"); exit()
    root = tk.Tk()
    app = MassSlideGeneratorApp(root)
    root.mainloop()