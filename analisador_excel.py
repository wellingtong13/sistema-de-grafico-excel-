"""
ANALISADOR DE DADOS EXCEL — DataVista
Interface com tema claro/escuro, navegação funcional, gráficos e pré-visualização tabular.

Dependências:
py -3.14 -m pip install customtkinter pillow pandas matplotlib openpyxl xlrd

Execute:
py -3.14 analisador_excel.py

Marca:
Logotipo oficial SENAI em assets/senai_logo.png, obtido no portal oficial:
https://www.senai.portaldaindustria.com.br/para-voce
"""

import os
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk
import pandas as pd
import matplotlib.pyplot as plt
from PIL import Image
from matplotlib.backends.backend_tkagg import (
    FigureCanvasTkAgg,
    NavigationToolbar2Tk
)


# ============================================================
# CONFIGURAÇÕES VISUAIS
# ============================================================

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

COR_PRIMARIA = "#E1262F"
COR_PRIMARIA_HOVER = "#BF1E27"
COR_FUNDO_CLARO = "#F3F4F6"
COR_TEXTO = "#20242D"
COR_SECUNDARIA = "#687282"
COR_SUCESSO = "#158052"
COR_ERRO = "#DC2626"
COR_GRAFITE = "#171B25"
COR_GRAFITE_CLARO = "#222A38"
COR_AZUL_ELETRICO = "#237BFA"
COR_CIANO = "#5BD6E8"
COR_ROXO = "#9B80F5"
COR_FUNDO_LOGO = "#FCFCFC"

CORES_TEMA_ESCURO = {
    COR_FUNDO_CLARO: "#0D1017",
    "#FFFFFF": "#151922",
    "#E5EAF2": "#2B303B",
    COR_TEXTO: "#F1F3F6",
    COR_SECUNDARIA: "#A6ADBA",
    "#111C35": "#101319",
    "#91A4C7": "#8793A7",
    "#7183A6": "#66738A",
    "#1B2A4A": "#1C222D",
    "#A9B8D2": "#D6DFEA",
    "#263B65": "#303B4D",
    "#94A3B8": "#8793A7",
    "#EEF3FB": "#232A35",
    "#E9EFF8": "#242B37",
    "#DCE6F4": "#333C4A",
    "#334155": "#D5DEEA",
    "#EAF1FF": "#28283B",
    "#E8EDF4": "#303742",
    "#475569": "#ADB4C1",
    "#93C5FD": "#5CCFEC",
    "#7C3AED": "#B39AFB",
    "#CBD5E1": "#424A58",
    COR_PRIMARIA: "#F14952",
    COR_PRIMARIA_HOVER: "#CE3039",
    COR_AZUL_ELETRICO: "#4998FF",
    COR_CIANO: "#72E5F0",
    COR_ROXO: "#B39AFB",
    COR_GRAFITE: "#171B25",
    COR_GRAFITE_CLARO: "#222A38",
    "#F1F5F9": "#F1F3F6",
    "#E2E5EA": "#262D39",
    "#D8DDE5": "#262D39",
    "#202735": "#222A38",
    "#D8E0EB": "#D8E0EB",
}

CORES_TEMA_CLARO = {
    cor_escura: cor_clara
    for cor_clara, cor_escura in CORES_TEMA_ESCURO.items()
}

CORES_TEXTO_TEMA_ESCURO = {
    "#FFFFFF": "#FFFFFF",
    COR_TEXTO: "#F1F3F6",
    COR_SECUNDARIA: "#A6ADBA",
    "#91A4C7": "#91A4C7",
    "#7183A6": "#74839A",
    "#A9B8D2": "#D6DFEA",
    "#334155": "#D5DEEA",
    "#475569": "#ADB4C1",
    "#94A3B8": "#7F91A9",
    "#CBD5E1": "#CBD5E1",
    COR_PRIMARIA: "#F14952",
    COR_AZUL_ELETRICO: "#4998FF",
    COR_CIANO: "#72E5F0",
    COR_ROXO: "#B39AFB",
}

CORES_TEXTO_TEMA_CLARO = {
    cor_escura: cor_clara
    for cor_clara, cor_escura in CORES_TEXTO_TEMA_ESCURO.items()
}


# ============================================================
# CLASSE PRINCIPAL
# ============================================================

class AnalisadorExcel:

    def __init__(self, root):

        self.root = root
        self.root.title("DataVista | Analisador Excel")
        self.root.geometry("1280x820")
        self.root.minsize(1040, 740)
        self.root.configure(fg_color=COR_FUNDO_CLARO)

        self.df = None
        self.caminho_arquivo = None
        self.colunas_numericas = []
        self.modo_escuro = False
        self.itens_menu = {}
        self._carregar_logo()

        self._configurar_estilo_graficos()
        self.criar_interface()

    def _carregar_logo(self):
        caminho_logo = (
            Path(__file__).resolve().parent
            / "assets"
            / "senai_logo.png"
        )
        with Image.open(caminho_logo) as logo_original:
            imagem = logo_original.convert("RGBA")

        area_visivel = imagem.getchannel("A").getbbox()
        if area_visivel is None:
            raise ValueError(f"O logotipo está vazio: {caminho_logo}")

        imagem = imagem.crop(area_visivel)
        self.logo_sidebar = ctk.CTkImage(
            light_image=imagem,
            dark_image=imagem,
            size=(66, 66)
        )
        self.logo_inicio = ctk.CTkImage(
            light_image=imagem,
            dark_image=imagem,
            size=(58, 58)
        )

    # ========================================================
    # INTERFACE PRINCIPAL
    # ========================================================

    def criar_interface(self):

        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(0, weight=1)

        self.criar_sidebar()

        conteudo = ctk.CTkScrollableFrame(
            self.root,
            fg_color="transparent",
            corner_radius=0,
            scrollbar_button_color=COR_PRIMARIA,
            scrollbar_button_hover_color=COR_PRIMARIA_HOVER
        )

        conteudo.grid(
            row=0,
            column=1,
            sticky="nsew",
            padx=(28, 30),
            pady=25
        )

        conteudo.grid_columnconfigure(0, weight=1)
        conteudo.grid_rowconfigure(3, weight=1)

        # Cabeçalho

        topo = ctk.CTkFrame(
            conteudo,
            fg_color=COR_GRAFITE,
            corner_radius=20,
            height=158
        )

        topo.grid(
            row=0,
            column=0,
            sticky="ew",
            pady=(0, 20)
        )

        topo.grid_columnconfigure(0, weight=1)
        topo.grid_columnconfigure(1, minsize=242)
        topo.grid_columnconfigure(2, minsize=88)
        topo.grid_propagate(False)

        ctk.CTkLabel(
            topo,
            text="SENAI  /  TECNOLOGIA • EDUCAÇÃO • DADOS",
            text_color="#FF6A72",
            font=ctk.CTkFont(
                family="Segoe UI",
                size=10,
                weight="bold"
            )
        ).grid(
            row=0,
            column=0,
            sticky="w",
            padx=24,
            pady=(18, 0)
        )

        ctk.CTkLabel(
            topo,
            text="Dados que viram\nnovas ideias.",
            text_color="#FFFFFF",
            font=ctk.CTkFont(
                family="Segoe UI",
                size=25,
                weight="bold"
            ),
            justify="left"
        ).grid(
            row=1,
            column=0,
            sticky="w",
            padx=24,
            pady=(4, 0)
        )

        ctk.CTkLabel(
            topo,
            text="Explore sua planilha, descubra tendências e compartilhe resultados.",
            text_color="#C1C9D6",
            font=ctk.CTkFont(size=11),
            wraplength=430,
            justify="left"
        ).grid(
            row=2,
            column=0,
            sticky="w",
            padx=24,
            pady=(4, 16)
        )

        self.status_topo = ctk.CTkLabel(
            topo,
            text="●  Aguardando arquivo",
            text_color="#D8E0EB",
            font=ctk.CTkFont(
                size=10,
                weight="bold"
            ),
            fg_color="#2A303C",
            corner_radius=10,
            padx=9,
            pady=6
        )

        self.status_topo.grid(
            row=1,
            column=2,
            sticky="n",
            padx=(0, 16),
            pady=(10, 0)
        )

        self._criar_ilustracao_hero(topo)

        logo_card = ctk.CTkFrame(
            topo,
            fg_color=COR_FUNDO_LOGO,
            corner_radius=16,
            width=70,
            height=70
        )
        logo_card.grid_propagate(False)
        logo_card.grid(
            row=0,
            column=2,
            sticky="n",
            padx=(0, 17),
            pady=(15, 0)
        )
        ctk.CTkLabel(
            logo_card,
            text="",
            image=self.logo_inicio
        ).place(
            relx=0.5,
            rely=0.5,
            anchor="center"
        )

        # Cartões de indicadores

        indicadores = ctk.CTkFrame(
            conteudo,
            fg_color="transparent"
        )

        indicadores.grid(
            row=1,
            column=0,
            sticky="ew",
            pady=(0, 20)
        )

        for coluna in range(3):
            indicadores.grid_columnconfigure(
                coluna,
                weight=1,
                uniform="cards"
            )

        self.card_linhas = self.criar_card(
            indicadores,
            0,
            "REGISTROS",
            "—",
            "Linhas da planilha"
        )

        self.card_colunas = self.criar_card(
            indicadores,
            1,
            "COLUNAS NUMÉRICAS",
            "—",
            "Variáveis disponíveis"
        )

        self.card_arquivo = self.criar_card(
            indicadores,
            2,
            "ARQUIVO",
            "Nenhum",
            "Planilha carregada"
        )

        # Área de importação

        painel_arquivo = ctk.CTkFrame(
            conteudo,
            fg_color="#FFFFFF",
            corner_radius=16,
            border_width=1,
            border_color="#E5EAF2"
        )

        painel_arquivo.grid(
            row=2,
            column=0,
            sticky="ew",
            pady=(0, 18)
        )

        painel_arquivo.grid_columnconfigure(0, weight=1)

        linha_arquivo = ctk.CTkFrame(
            painel_arquivo,
            fg_color="transparent"
        )

        linha_arquivo.grid(
            row=0,
            column=0,
            sticky="ew",
            padx=20,
            pady=17
        )

        linha_arquivo.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            linha_arquivo,
            text="XLS",
            width=54,
            height=50,
            corner_radius=13,
            fg_color="#FCEAEC",
            text_color=COR_PRIMARIA,
            font=ctk.CTkFont(size=13, weight="bold")
        ).grid(
            row=0,
            column=0,
            rowspan=2,
            padx=(0, 14)
        )

        ctk.CTkLabel(
            linha_arquivo,
            text="Sua planilha",
            text_color=COR_TEXTO,
            font=ctk.CTkFont(
                size=15,
                weight="bold"
            )
        ).grid(
            row=0,
            column=1,
            sticky="w"
        )

        self.lbl_arquivo = ctk.CTkLabel(
            linha_arquivo,
            text="Selecione um arquivo .xlsx ou .xls para começar.",
            text_color=COR_SECUNDARIA,
            font=ctk.CTkFont(size=12),
            anchor="w"
        )

        self.lbl_arquivo.grid(
            row=1,
            column=1,
            sticky="ew",
            pady=(3, 0)
        )

        self.btn_selecionar = ctk.CTkButton(
            linha_arquivo,
            text="＋  Importar planilha",
            command=self.selecionar_arquivo,
            height=42,
            width=190,
            corner_radius=10,
            fg_color=COR_PRIMARIA,
            hover_color=COR_PRIMARIA_HOVER,
            font=ctk.CTkFont(
                size=13,
                weight="bold"
            )
        )

        self.btn_selecionar.grid(
            row=0,
            column=2,
            rowspan=2,
            padx=(15, 0)
        )

        # Área inferior

        inferior = ctk.CTkFrame(
            conteudo,
            fg_color="transparent"
        )

        inferior.grid(
            row=3,
            column=0,
            sticky="nsew"
        )

        inferior.grid_columnconfigure(0, weight=5)
        inferior.grid_columnconfigure(1, weight=4)
        inferior.grid_rowconfigure(0, weight=1)

        self.criar_painel_configuracao(inferior)
        self.criar_painel_dicas(inferior)

        # Rodapé

        self.lbl_status = ctk.CTkLabel(
            conteudo,
            text="Pronto para começar. Importe uma planilha para liberar as análises.",
            text_color=COR_SECUNDARIA,
            font=ctk.CTkFont(size=11),
            anchor="w"
        )

        self.lbl_status.grid(
            row=4,
            column=0,
            sticky="ew",
            pady=(12, 0)
        )

    def _criar_ilustracao_hero(self, parent):
        moldura = ctk.CTkFrame(
            parent,
            fg_color=COR_GRAFITE_CLARO,
            corner_radius=15,
            width=232,
            height=126
        )
        moldura.place(
            relx=0.60,
            rely=0.08,
            relwidth=0.27,
            relheight=0.84
        )

        self.canvas_hero = tk.Canvas(
            moldura,
            background=COR_GRAFITE_CLARO,
            borderwidth=0,
            highlightthickness=0
        )
        self.canvas_hero.pack(
            fill="both",
            expand=True,
            padx=8,
            pady=8
        )
        self.canvas_hero.bind(
            "<Configure>",
            self._desenhar_ilustracao_dados
        )

    def _desenhar_ilustracao_dados(self, evento=None):
        canvas = self.canvas_hero
        largura = evento.width if evento is not None else canvas.winfo_width()
        altura = evento.height if evento is not None else canvas.winfo_height()
        if largura < 120 or altura < 80:
            return

        canvas.delete("all")
        x = lambda proporcao: largura * proporcao
        y = lambda proporcao: altura * proporcao

        horizonte = y(0.84)
        for indice in range(5):
            inicio_x = x(0.08 + indice * 0.18)
            canvas.create_line(
                inicio_x,
                y(0.28),
                x(0.50) + (inicio_x - x(0.50)) * 0.42,
                horizonte,
                fill="#344255",
                dash=(2, 6),
                width=1
            )
        for indice in range(4):
            linha_y = y(0.53 + indice * 0.09)
            canvas.create_line(
                x(0.08),
                linha_y,
                x(0.79),
                linha_y,
                fill="#344255",
                dash=(2, 6),
                width=1
            )

        altura_barra = (0.25, 0.43, 0.34, 0.60)
        cores_frente = (
            COR_PRIMARIA,
            COR_AZUL_ELETRICO,
            COR_ROXO,
            COR_CIANO
        )
        cores_lado = ("#9F252F", "#2358A8", "#6652AD", "#328DA0")
        largura_barra = x(0.072)
        profundidade_x = x(0.035)
        profundidade_y = y(0.045)
        pontos_linha = []

        for indice, fracao_altura in enumerate(altura_barra):
            inicio_x = x(0.15 + indice * 0.14)
            topo_y = horizonte - y(fracao_altura)
            fim_x = inicio_x + largura_barra
            base_y = horizonte
            pontos_linha.extend((inicio_x + largura_barra / 2, topo_y - profundidade_y))

            canvas.create_polygon(
                inicio_x,
                topo_y,
                fim_x,
                topo_y,
                fim_x,
                base_y,
                inicio_x,
                base_y,
                fill=cores_frente[indice],
                outline=""
            )
            canvas.create_polygon(
                fim_x,
                topo_y,
                fim_x + profundidade_x,
                topo_y - profundidade_y,
                fim_x + profundidade_x,
                base_y - profundidade_y,
                fim_x,
                base_y,
                fill=cores_lado[indice],
                outline=""
            )
            canvas.create_polygon(
                inicio_x,
                topo_y,
                inicio_x + profundidade_x,
                topo_y - profundidade_y,
                fim_x + profundidade_x,
                topo_y - profundidade_y,
                fim_x,
                topo_y,
                fill="#AEBECD",
                outline=""
            )

        for indice in range(0, len(pontos_linha) - 2, 2):
            canvas.create_line(
                pontos_linha[indice],
                pontos_linha[indice + 1],
                pontos_linha[indice + 2],
                pontos_linha[indice + 3],
                fill="#E0EBF8",
                width=2,
                smooth=True
            )
        for indice in range(0, len(pontos_linha), 2):
            ponto_x = pontos_linha[indice]
            ponto_y = pontos_linha[indice + 1]
            raio = max(2, largura * 0.012)
            canvas.create_oval(
                ponto_x - raio,
                ponto_y - raio,
                ponto_x + raio,
                ponto_y + raio,
                fill="#FFFFFF",
                outline=COR_AZUL_ELETRICO,
                width=2
            )

        orbita = (x(0.69), y(0.15), x(0.95), y(0.48))
        canvas.create_oval(
            *orbita,
            outline="#42617F",
            width=1
        )
        canvas.create_oval(
            x(0.765),
            y(0.18),
            x(0.88),
            y(0.34),
            fill="#294462",
            outline="#69D5E8",
            width=2
        )
        canvas.create_oval(
            x(0.795),
            y(0.225),
            x(0.85),
            y(0.30),
            fill=COR_CIANO,
            outline=""
        )
        canvas.create_oval(
            x(0.91),
            y(0.19),
            x(0.935),
            y(0.225),
            fill=COR_PRIMARIA,
            outline=""
        )
        canvas.create_oval(
            x(0.70),
            y(0.40),
            x(0.72),
            y(0.43),
            fill=COR_ROXO,
            outline=""
        )
        canvas.create_text(
            x(0.08),
            y(0.10),
            text="INSIGHTS EM MOVIMENTO",
            fill="#AEBBD0",
            anchor="w",
            font=("Segoe UI", 8, "bold")
        )

    # ========================================================
    # MENU LATERAL
    # ========================================================

    def criar_sidebar(self):

        sidebar = ctk.CTkFrame(
            self.root,
            width=225,
            corner_radius=0,
            fg_color="#151922"
        )

        sidebar.grid(
            row=0,
            column=0,
            sticky="nsew"
        )

        sidebar.grid_propagate(False)

        marca = ctk.CTkFrame(
            sidebar,
            fg_color="transparent"
        )

        marca.pack(
            fill="x",
            padx=17,
            pady=(23, 32)
        )

        logo_card = ctk.CTkFrame(
            marca,
            fg_color=COR_FUNDO_LOGO,
            corner_radius=13,
            width=60,
            height=60
        )
        logo_card.pack(side="left", padx=(0, 12))
        logo_card.pack_propagate(False)
        ctk.CTkLabel(
            logo_card,
            text="",
            image=self.logo_sidebar
        ).place(relx=0.5, rely=0.5, anchor="center")

        texto_marca = ctk.CTkFrame(
            marca,
            fg_color="transparent"
        )
        texto_marca.pack(side="left", fill="x", expand=True)

        ctk.CTkLabel(
            texto_marca,
            text="DataVista",
            text_color="#FFFFFF",
            font=ctk.CTkFont(
                size=17,
                weight="bold"
            )
        ).pack(anchor="w")

        ctk.CTkLabel(
            texto_marca,
            text="ANÁLISE EXCEL",
            text_color="#A9B2C1",
            font=ctk.CTkFont(
                size=9,
                weight="bold"
            )
        ).pack(
            anchor="w",
            pady=(3, 0)
        )

        ctk.CTkFrame(
            texto_marca,
            width=35,
            height=3,
            fg_color=COR_PRIMARIA,
            corner_radius=2
        ).pack(
            anchor="w",
            pady=(6, 0)
        )

        ctk.CTkLabel(
            sidebar,
            text="MENU PRINCIPAL",
            text_color="#7183A6",
            font=ctk.CTkFont(
                size=10,
                weight="bold"
            )
        ).pack(
            anchor="w",
            padx=23,
            pady=(0, 12)
        )

        self.item_menu(
            sidebar,
            "▦   Visão geral",
            ativo=True,
            command=self._mostrar_visao_geral
        )

        self.item_menu(
            sidebar,
            "↗   Criar gráficos",
            command=self._ir_para_graficos
        )

        self.item_menu(
            sidebar,
            "▤   Dados da planilha",
            command=self.visualizar_dados
        )

        self.switch_tema = ctk.CTkSwitch(
            sidebar,
            text="Tema escuro",
            command=self.alternar_tema,
            text_color="#CBD5E1",
            progress_color=COR_PRIMARIA,
            button_color=COR_PRIMARIA,
            button_hover_color=COR_PRIMARIA_HOVER,
            font=ctk.CTkFont(size=11)
        )
        self.switch_tema.pack(
            side="bottom",
            anchor="w",
            padx=20,
            pady=(0, 15)
        )
        if self.modo_escuro:
            self.switch_tema.select()

        rodape = ctk.CTkFrame(
            sidebar,
            fg_color="#212733",
            corner_radius=12
        )

        rodape.pack(
            side="bottom",
            fill="x",
            padx=14,
            pady=18
        )

        ctk.CTkLabel(
            rodape,
            text="ANÁLISE SIMPLES",
            text_color="#FFFFFF",
            font=ctk.CTkFont(
                size=11,
                weight="bold"
            )
        ).pack(
            anchor="w",
            padx=13,
            pady=(12, 3)
        )

        ctk.CTkLabel(
            rodape,
            text="Visualize seus dados\ncom clareza e praticidade.",
            text_color="#A9B8D2",
            justify="left",
            font=ctk.CTkFont(size=10)
        ).pack(
            anchor="w",
            padx=13,
            pady=(0, 12)
        )

    def item_menu(self, parent, texto, ativo=False, command=None):

        botao = ctk.CTkButton(
            parent,
            text=texto,
            command=None,
            anchor="w",
            height=40,
            corner_radius=8,
            fg_color=COR_PRIMARIA if ativo else "transparent",
            text_color="#FFFFFF" if ativo else "#C1C9D6",
            hover_color=COR_PRIMARIA_HOVER if ativo else "#282F3C",
            border_spacing=12,
            font=ctk.CTkFont(
                size=12,
                weight="bold" if ativo else "normal"
            )
        )
        botao.configure(
            command=lambda item=botao: self._ativar_item_menu(
                item,
                command
            )
        )
        self.itens_menu[texto] = botao
        botao.pack(
            fill="x",
            padx=13,
            pady=3
        )

    def _ativar_item_menu(self, selecionado, command):
        for botao in self.itens_menu.values():
            ativo = botao is selecionado
            botao.configure(
                fg_color=COR_PRIMARIA if ativo else "transparent",
                text_color="#FFFFFF" if ativo else "#C1C9D6",
                hover_color=COR_PRIMARIA_HOVER if ativo else "#282F3C",
                font=ctk.CTkFont(
                    size=12,
                    weight="bold" if ativo else "normal"
                )
            )

        if command is not None:
            command()

    def _mostrar_visao_geral(self):
        self.lbl_status.configure(
            text="Visão geral do analisador.",
            text_color=COR_SECUNDARIA
        )

    def _ir_para_graficos(self):
        self.lbl_status.configure(
            text="Configure os eixos e o tipo de gráfico no painel abaixo.",
            text_color=COR_SECUNDARIA
        )
        self.combo_tipo.focus_set()

    def alternar_tema(self):
        self.modo_escuro = bool(self.switch_tema.get())
        ctk.set_appearance_mode("dark" if self.modo_escuro else "light")

        paleta = (
            CORES_TEMA_ESCURO
            if self.modo_escuro
            else CORES_TEMA_CLARO
        )
        self._aplicar_paleta(self.root, paleta)
        self._configurar_estilo_tabela(ttk.Style(self.root))

    def _aplicar_paleta(self, widget, paleta):
        paleta_texto = (
            CORES_TEXTO_TEMA_ESCURO
            if self.modo_escuro
            else CORES_TEXTO_TEMA_CLARO
        )

        opcoes_cor = (
            "fg_color",
            "bg_color",
            "border_color",
            "text_color",
            "placeholder_text_color",
            "button_color",
            "button_hover_color",
            "dropdown_fg_color",
            "dropdown_text_color",
            "dropdown_hover_color",
            "text_color_disabled",
            "hover_color",
            "foreground",
            "background",
            "activeforeground",
            "activebackground",
            "highlightbackground",
            "highlightcolor",
            "selectforeground",
            "selectbackground",
        )

        for nome in opcoes_cor:
            try:
                cor_atual = widget.cget(nome)
            except (KeyError, tk.TclError, ValueError):
                continue

            cores = (
                paleta_texto
                if nome in ("text_color", "foreground")
                else paleta
            )
            if isinstance(cor_atual, str) and cor_atual in cores:
                widget.configure(**{nome: cores[cor_atual]})

        for filho in widget.winfo_children():
            self._aplicar_paleta(filho, paleta)

    def _configurar_estilo_tabela(self, estilo):
        estilo.theme_use("clam")

        fundo = "#111827" if self.modo_escuro else "#FFFFFF"
        texto = "#F1F5F9" if self.modo_escuro else COR_TEXTO
        fundo_cabecalho = "#1A2638" if self.modo_escuro else "#F1F5F9"
        cor_scrollbar = "#263247" if self.modo_escuro else "#E5EAF2"

        estilo.configure(
            "DataVista.Treeview",
            background=fundo,
            fieldbackground=fundo,
            foreground=texto,
            rowheight=32,
            font=("Segoe UI", 10),
            borderwidth=0
        )
        estilo.configure(
            "DataVista.Treeview.Heading",
            background=fundo_cabecalho,
            foreground=texto,
            font=("Segoe UI", 10, "bold"),
            relief="flat"
        )
        for orientacao in ("Horizontal", "Vertical"):
            estilo.configure(
                f"DataVista.{orientacao}.TScrollbar",
                background=cor_scrollbar,
                troughcolor=fundo,
                bordercolor=fundo,
                arrowcolor=texto
            )
        estilo.map(
            "DataVista.Treeview",
            background=[("selected", COR_PRIMARIA)],
            foreground=[("selected", "#FFFFFF")]
        )

    # ========================================================
    # CARTÕES
    # ========================================================

    def criar_card(
        self,
        parent,
        coluna,
        titulo,
        valor,
        descricao
    ):

        sombra = ctk.CTkFrame(
            parent,
            fg_color="#D8DDE5",
            corner_radius=14
        )

        sombra.grid(
            row=0,
            column=coluna,
            sticky="ew",
            padx=(0 if coluna == 0 else 7, 7)
        )

        card = ctk.CTkFrame(
            sombra,
            fg_color="#FFFFFF",
            corner_radius=14,
            border_width=1,
            border_color="#E5EAF2"
        )
        card.pack(fill="both", expand=True, padx=1, pady=(0, 3))

        ctk.CTkFrame(
            card,
            width=28,
            height=3,
            fg_color=COR_PRIMARIA,
            corner_radius=2
        ).pack(
            anchor="w",
            padx=16,
            pady=(12, 0)
        )

        ctk.CTkLabel(
            card,
            text=titulo,
            text_color=COR_SECUNDARIA,
            font=ctk.CTkFont(
                size=10,
                weight="bold"
            )
        ).pack(
            anchor="w",
            padx=16,
            pady=(9, 4)
        )

        valor_label = ctk.CTkLabel(
            card,
            text=valor,
            text_color=COR_TEXTO,
            font=ctk.CTkFont(
                size=23,
                weight="bold"
            )
        )

        valor_label.pack(
            anchor="w",
            padx=16
        )

        ctk.CTkLabel(
            card,
            text=descricao,
            text_color=COR_SECUNDARIA,
            font=ctk.CTkFont(size=10)
        ).pack(
            anchor="w",
            padx=16,
            pady=(2, 13)
        )

        return valor_label

    # ========================================================
    # PAINEL DE CONFIGURAÇÃO
    # ========================================================

    def criar_painel_configuracao(self, parent):

        painel = ctk.CTkFrame(
            parent,
            fg_color="#FFFFFF",
            corner_radius=16,
            border_width=1,
            border_color="#E5EAF2"
        )

        painel.grid(
            row=0,
            column=0,
            sticky="nsew",
            padx=(0, 10)
        )

        painel.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            painel,
            text="Monte sua visualização",
            text_color=COR_TEXTO,
            font=ctk.CTkFont(
                size=17,
                weight="bold"
            )
        ).pack(
            anchor="w",
            padx=20,
            pady=(20, 4)
        )

        ctk.CTkLabel(
            painel,
            text="Escolha como deseja comparar seus dados.",
            text_color=COR_SECUNDARIA,
            font=ctk.CTkFont(size=11)
        ).pack(
            anchor="w",
            padx=20,
            pady=(0, 18)
        )

        self.rotulo(painel, "TIPO DE GRÁFICO")

        self.combo_tipo = ctk.CTkOptionMenu(
            painel,
            values=[
                "Dispersão",
                "Linha",
                "Barras",
                "Histograma"
            ],
            height=39,
            corner_radius=9,
            fg_color="#EEF3FB",
            button_color=COR_PRIMARIA,
            button_hover_color=COR_PRIMARIA_HOVER,
            text_color=COR_TEXTO,
            dropdown_fg_color="#FFFFFF",
            dropdown_text_color=COR_TEXTO
        )

        self.combo_tipo.pack(
            fill="x",
            padx=20,
            pady=(5, 14)
        )

        self.rotulo(painel, "EIXO X")

        self.combo_x = ctk.CTkOptionMenu(
            painel,
            values=["Aguardando planilha"],
            height=39,
            corner_radius=9,
            fg_color="#EEF3FB",
            button_color=COR_PRIMARIA,
            button_hover_color=COR_PRIMARIA_HOVER,
            text_color=COR_TEXTO,
            dropdown_fg_color="#FFFFFF",
            dropdown_text_color=COR_TEXTO
        )

        self.combo_x.pack(
            fill="x",
            padx=20,
            pady=(5, 14)
        )

        self.rotulo(painel, "EIXO Y")

        self.combo_y = ctk.CTkOptionMenu(
            painel,
            values=["Aguardando planilha"],
            height=39,
            corner_radius=9,
            fg_color="#EEF3FB",
            button_color=COR_PRIMARIA,
            button_hover_color=COR_PRIMARIA_HOVER,
            text_color=COR_TEXTO,
            dropdown_fg_color="#FFFFFF",
            dropdown_text_color=COR_TEXTO
        )

        self.combo_y.pack(
            fill="x",
            padx=20,
            pady=(5, 18)
        )

        self.btn_grafico = ctk.CTkButton(
            painel,
            text="✦   GERAR GRÁFICO",
            command=self.gerar_grafico,
            height=47,
            corner_radius=10,
            fg_color=COR_PRIMARIA,
            hover_color=COR_PRIMARIA_HOVER,
            font=ctk.CTkFont(
                size=13,
                weight="bold"
            )
        )

        self.btn_grafico.pack(
            fill="x",
            padx=20,
            pady=(0, 10)
        )

        self.btn_visualizar = ctk.CTkButton(
            painel,
            text="Pré-visualizar dados",
            command=self.visualizar_dados,
            height=36,
            corner_radius=9,
            fg_color="#E9EFF8",
            hover_color="#DCE6F4",
            text_color="#334155",
            font=ctk.CTkFont(
                size=11,
                weight="bold"
            )
        )

        self.btn_visualizar.pack(
            fill="x",
            padx=20,
            pady=(0, 18)
        )

    # ========================================================
    # PAINEL DE DICAS
    # ========================================================

    def criar_painel_dicas(self, parent):

        painel = ctk.CTkFrame(
            parent,
            fg_color="#FFFFFF",
            corner_radius=16,
            border_width=1,
            border_color="#E5EAF2"
        )

        painel.grid(
            row=0,
            column=1,
            sticky="nsew",
            padx=(10, 0)
        )

        painel.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            painel,
            text="Dicas para uma boa análise",
            text_color=COR_TEXTO,
            font=ctk.CTkFont(
                size=16,
                weight="bold"
            )
        ).pack(
            anchor="w",
            padx=18,
            pady=(20, 15)
        )

        dicas = [
            (
                "01",
                "Dispersão",
                "Ideal para observar a relação entre duas variáveis."
            ),
            (
                "02",
                "Linha",
                "Ajuda a visualizar tendências e mudanças ao longo do tempo."
            ),
            (
                "03",
                "Barras",
                "Compare valores entre categorias ou registros."
            ),
            (
                "04",
                "Histograma",
                "Entenda como os valores de uma variável se distribuem."
            )
        ]

        for numero, titulo, descricao in dicas:

            linha = ctk.CTkFrame(
                painel,
                fg_color="transparent"
            )

            linha.pack(
                fill="x",
                padx=18,
                pady=7
            )

            ctk.CTkLabel(
                linha,
                text=numero,
                width=31,
                height=31,
                corner_radius=9,
                fg_color="#EAF1FF",
                text_color=COR_PRIMARIA,
                font=ctk.CTkFont(
                    size=11,
                    weight="bold"
                )
            ).pack(
                side="left",
                anchor="n",
                padx=(0, 10)
            )

            textos = ctk.CTkFrame(
                linha,
                fg_color="transparent"
            )

            textos.pack(
                side="left",
                fill="x",
                expand=True
            )

            ctk.CTkLabel(
                textos,
                text=titulo,
                text_color=COR_TEXTO,
                font=ctk.CTkFont(
                    size=12,
                    weight="bold"
                )
            ).pack(anchor="w")

            ctk.CTkLabel(
                textos,
                text=descricao,
                text_color=COR_SECUNDARIA,
                font=ctk.CTkFont(size=10),
                wraplength=205,
                justify="left"
            ).pack(
                anchor="w",
                pady=(2, 0)
            )

        ctk.CTkFrame(
            painel,
            height=1,
            fg_color="#E8EDF4"
        ).pack(
            fill="x",
            padx=18,
            pady=(16, 12)
        )

        ctk.CTkLabel(
            painel,
            text="Dica: use nomes claros nas colunas da planilha.",
            text_color=COR_SECUNDARIA,
            wraplength=250,
            justify="left",
            font=ctk.CTkFont(size=10)
        ).pack(
            anchor="w",
            padx=18,
            pady=(0, 18)
        )

    def rotulo(self, parent, texto):

        label = ctk.CTkLabel(
            parent,
            text=texto,
            text_color="#475569",
            font=ctk.CTkFont(
                size=10,
                weight="bold"
            )
        )

        label.pack(
            anchor="w",
            padx=20
        )

        return label

    # ========================================================
    # SELEÇÃO E VALIDAÇÃO DO EXCEL
    # ========================================================

    def selecionar_arquivo(self):

        caminho = filedialog.askopenfilename(
            title="Selecione uma planilha Excel",
            filetypes=[
                ("Planilhas Excel", "*.xlsx *.xls"),
                ("Todos os arquivos", "*.*")
            ]
        )

        if not caminho:
            return

        extensao = Path(caminho).suffix.lower()

        if extensao not in (".xlsx", ".xls"):

            messagebox.showerror(
                "Formato não suportado",
                "Selecione um arquivo Excel com extensão .xlsx ou .xls."
            )

            return

        try:

            motor = (
                "openpyxl"
                if extensao == ".xlsx"
                else "xlrd"
            )

            novo_df = pd.read_excel(
                caminho,
                engine=motor
            )

            if novo_df.empty:
                raise ValueError("A planilha está vazia.")

            numericas = novo_df.select_dtypes(
                include="number"
            ).columns.tolist()

            if not numericas:
                raise ValueError(
                    "Não encontrei colunas numéricas nessa planilha."
                )

            self.df = novo_df
            self.caminho_arquivo = caminho
            self.colunas_numericas = numericas

            nome = os.path.basename(caminho)

            self.lbl_arquivo.configure(
                text=f"{nome}  •  {len(novo_df):,} registros".replace(",", ".")
            )

            self.card_linhas.configure(
                text=f"{len(novo_df):,}".replace(",", ".")
            )

            self.card_colunas.configure(
                text=str(len(numericas))
            )

            self.card_arquivo.configure(
                text=nome if len(nome) <= 17 else nome[:14] + "…"
            )

            self.combo_x.configure(
                values=numericas
            )

            self.combo_y.configure(
                values=numericas
            )

            self.combo_x.set(numericas[0])

            self.combo_y.set(
                numericas[1]
                if len(numericas) > 1
                else numericas[0]
            )

            self.status_topo.configure(
                text="●  Planilha carregada",
                text_color="#6ED7A1"
            )

            self.lbl_status.configure(
                text=f"✓  Dados carregados: {len(novo_df)} linhas e {len(novo_df.columns)} colunas.",
                text_color=COR_SUCESSO
            )

        except Exception as erro:

            self.df = None
            self.caminho_arquivo = None

            self.status_topo.configure(
                text="●  Falha ao carregar",
                text_color="#FF737A"
            )

            messagebox.showerror(
                "Não foi possível abrir a planilha",
                str(erro)
            )

    # ========================================================
    # PREPARAÇÃO DOS DADOS
    # ========================================================

    def obter_dados(self, coluna_x, coluna_y=None):

        colunas = (
            [coluna_x]
            if coluna_y is None
            else [coluna_x, coluna_y]
        )

        dados = self.df[colunas].copy()

        for coluna in colunas:

            dados[coluna] = pd.to_numeric(
                dados[coluna],
                errors="coerce"
            )

        dados = dados.replace(
            [
                float("inf"),
                float("-inf")
            ],
            float("nan")
        ).dropna()

        return dados

    # ========================================================
    # GERAÇÃO DOS GRÁFICOS
    # ========================================================

    def gerar_grafico(self):

        if self.df is None:

            messagebox.showwarning(
                "Planilha necessária",
                "Primeiro importe uma planilha Excel."
            )

            return

        tipo = self.combo_tipo.get()
        coluna_x = self.combo_x.get()
        coluna_y = self.combo_y.get()

        if tipo != "Histograma" and coluna_x == coluna_y:

            messagebox.showwarning(
                "Ajuste os eixos",
                "Escolha duas colunas diferentes para o eixo X e Y."
            )

            return

        try:

            dados = self.obter_dados(
                coluna_x,
                None if tipo == "Histograma" else coluna_y
            )

            if dados.empty:

                raise ValueError(
                    "Não existem valores numéricos válidos para montar este gráfico."
                )

            if tipo == "Histograma" and len(dados) < 2:

                raise ValueError(
                    "São necessários pelo menos dois valores para o histograma."
                )

            self.abrir_janela_grafico(
                tipo,
                coluna_x,
                coluna_y,
                dados
            )

        except Exception as erro:

            messagebox.showerror(
                "Erro ao gerar gráfico",
                str(erro)
            )

    # ========================================================
    # JANELA DO GRÁFICO
    # ========================================================

    def abrir_janela_grafico(
        self,
        tipo,
        coluna_x,
        coluna_y,
        dados
    ):

        janela = ctk.CTkToplevel(self.root)

        janela.title(
            f"DataVista | Gráfico de {tipo}"
        )

        janela.geometry("1050x730")
        janela.minsize(820, 600)
        janela.configure(fg_color="#F4F7FB")
        janela.transient(self.root)

        # Cabeçalho

        topo = ctk.CTkFrame(
            janela,
            fg_color="transparent"
        )

        topo.pack(
            fill="x",
            padx=25,
            pady=(20, 5)
        )

        ctk.CTkLabel(
            topo,
            text=f"Gráfico de {tipo.lower()}",
            text_color=COR_TEXTO,
            font=ctk.CTkFont(
                size=22,
                weight="bold"
            )
        ).pack(anchor="w")

        descricao = (
            f"Distribuição de {coluna_x}"
            if tipo == "Histograma"
            else f"{coluna_x} × {coluna_y}"
        )

        ctk.CTkLabel(
            topo,
            text=descricao,
            text_color=COR_SECUNDARIA,
            font=ctk.CTkFont(size=12)
        ).pack(
            anchor="w",
            pady=(3, 0)
        )

        # Cartão do gráfico

        cartao = ctk.CTkFrame(
            janela,
            fg_color="#FFFFFF",
            corner_radius=15,
            border_width=1,
            border_color="#E5EAF2"
        )

        cartao.pack(
            fill="both",
            expand=True,
            padx=22,
            pady=15
        )

        figura, ax = plt.subplots(
            figsize=(9.5, 5.6),
            dpi=100
        )

        fundo_grafico = "#111827" if self.modo_escuro else "#FFFFFF"
        cor_texto_grafico = "#F1F5F9" if self.modo_escuro else COR_TEXTO
        cor_eixo_grafico = "#CBD5E1" if self.modo_escuro else "#475569"
        cor_grade_grafico = "#53647E" if self.modo_escuro else "#94A3B8"
        cor_rotulo_grafico = "#A6B4C7" if self.modo_escuro else COR_SECUNDARIA

        figura.patch.set_facecolor(fundo_grafico)
        ax.set_facecolor(fundo_grafico)

        azul = COR_AZUL_ELETRICO
        azul_claro = COR_CIANO
        roxo = COR_ROXO

        # GRÁFICO DE DISPERSÃO

        if tipo == "Dispersão":

            ax.scatter(
                dados[coluna_x],
                dados[coluna_y],
                s=65,
                alpha=0.78,
                color=COR_PRIMARIA,
                edgecolors=fundo_grafico,
                linewidths=0.8,
                label="Observações",
                zorder=3
            )

            if (
                len(dados) >= 2
                and dados[coluna_x].nunique() > 1
            ):

                import numpy as np

                coef = np.polyfit(
                    dados[coluna_x],
                    dados[coluna_y],
                    1
                )

                ordem = dados[coluna_x].sort_values()

                ax.plot(
                    ordem,
                    np.polyval(coef, ordem),
                    color=roxo,
                    linewidth=2.2,
                    label="Tendência linear"
                )

            ax.set_xlabel(coluna_x)
            ax.set_ylabel(coluna_y)

            ax.legend(
                frameon=False,
                loc="best",
                labelcolor=cor_texto_grafico
            )

        # GRÁFICO DE LINHA

        elif tipo == "Linha":

            ordenados = dados.sort_values(
                coluna_x
            )

            ax.plot(
                ordenados[coluna_x],
                ordenados[coluna_y],
                color=COR_PRIMARIA,
                linewidth=2.4,
                marker="o",
                markersize=5,
                markerfacecolor=fundo_grafico,
                markeredgecolor=azul,
                markeredgewidth=1.6,
                label=coluna_y
            )

            ax.set_xlabel(coluna_x)
            ax.set_ylabel(coluna_y)

            ax.legend(
                frameon=False,
                labelcolor=cor_texto_grafico
            )

        # GRÁFICO DE BARRAS

        elif tipo == "Barras":

            amostra = dados.head(30)

            ax.bar(
                amostra[coluna_x].astype(str),
                amostra[coluna_y],
                color=azul_claro,
                edgecolor=azul,
                linewidth=0.7,
                width=0.68,
                label=coluna_y
            )

            ax.set_xlabel(coluna_x)
            ax.set_ylabel(coluna_y)

            ax.tick_params(
                axis="x",
                rotation=35
            )

            ax.legend(
                frameon=False,
                labelcolor=cor_texto_grafico
            )

            if len(dados) > 30:

                ax.set_title(
                    "Exibindo os 30 primeiros registros",
                    loc="left",
                    fontsize=9,
                    color=cor_rotulo_grafico
                )

        # HISTOGRAMA

        elif tipo == "Histograma":

            ax.hist(
                dados[coluna_x],
                bins="auto",
                color=azul_claro,
                edgecolor=azul,
                linewidth=1.0,
                alpha=0.9
            )

            ax.set_xlabel(coluna_x)
            ax.set_ylabel("Frequência")

        # Acabamento visual

        ax.set_title(
            "Visualização dos dados",
            loc="left",
            fontsize=15,
            fontweight="bold",
            color=cor_texto_grafico,
            pad=16
        )

        ax.grid(
            axis="y",
            linestyle="--",
            linewidth=0.8,
            alpha=0.25,
            color=cor_grade_grafico
        )

        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)

        ax.spines["left"].set_color(cor_grade_grafico)
        ax.spines["bottom"].set_color(cor_grade_grafico)

        ax.tick_params(
            colors=cor_rotulo_grafico,
            labelsize=9
        )

        ax.xaxis.label.set_color(cor_eixo_grafico)
        ax.yaxis.label.set_color(cor_eixo_grafico)

        figura.tight_layout(pad=2.2)

        # Exibição do gráfico

        canvas = FigureCanvasTkAgg(
            figura,
            master=cartao
        )

        canvas.draw()

        canvas.get_tk_widget().pack(
            fill="both",
            expand=True,
            padx=12,
            pady=(12, 0)
        )

        # Barra de ferramentas

        barra = NavigationToolbar2Tk(
            canvas,
            cartao,
            pack_toolbar=False
        )

        barra.update()

        barra.pack(
            side="bottom",
            fill="x",
            padx=10,
            pady=(0, 8)
        )

        # Rodapé

        rodape = ctk.CTkFrame(
            janela,
            fg_color="transparent"
        )

        rodape.pack(
            fill="x",
            padx=25,
            pady=(0, 17)
        )

        ctk.CTkLabel(
            rodape,
            text=f"{len(dados):,} observações válidas".replace(",", "."),
            text_color=COR_SECUNDARIA,
            font=ctk.CTkFont(size=11)
        ).pack(side="left")

        # Exportação

        def salvar_grafico():

            destino = filedialog.asksaveasfilename(
                parent=janela,
                defaultextension=".png",
                filetypes=[
                    ("Imagem PNG", "*.png"),
                    ("Imagem PDF", "*.pdf")
                ],
                initialfile=f"grafico_{tipo.lower()}"
            )

            if destino:

                figura.savefig(
                    destino,
                    dpi=220,
                    bbox_inches="tight",
                    facecolor=figura.get_facecolor()
                )

                messagebox.showinfo(
                    "Gráfico salvo",
                    f"Arquivo salvo em:\n{destino}",
                    parent=janela
                )

        ctk.CTkButton(
            rodape,
            text="↓  Exportar gráfico",
            command=salvar_grafico,
            width=165,
            height=38,
            corner_radius=9,
            fg_color=COR_PRIMARIA,
            hover_color=COR_PRIMARIA_HOVER,
            font=ctk.CTkFont(
                size=12,
                weight="bold"
            )
        ).pack(side="right")

        def ao_fechar():

            plt.close(figura)
            janela.destroy()

        janela.protocol(
            "WM_DELETE_WINDOW",
            ao_fechar
        )
        if self.modo_escuro:
            self._aplicar_paleta(
                janela,
                CORES_TEMA_ESCURO
            )

    # ========================================================
    # PRÉ-VISUALIZAÇÃO DOS DADOS
    # ========================================================

    def visualizar_dados(self):

        if self.df is None:

            messagebox.showwarning(
                "Planilha necessária",
                "Importe uma planilha antes de visualizar os dados."
            )

            return

        janela = ctk.CTkToplevel(self.root)

        janela.title(
            "DataVista | Pré-visualização dos dados"
        )

        janela.geometry("900x560")
        janela.configure(fg_color="#F4F7FB")

        ctk.CTkLabel(
            janela,
            text="Pré-visualização da planilha",
            text_color=COR_TEXTO,
            font=ctk.CTkFont(
                size=21,
                weight="bold"
            )
        ).pack(
            anchor="w",
            padx=22,
            pady=(20, 3)
        )

        ctk.CTkLabel(
            janela,
            text=f"Mostrando até 100 linhas de {len(self.df)} registros.",
            text_color=COR_SECUNDARIA,
            font=ctk.CTkFont(size=11)
        ).pack(
            anchor="w",
            padx=22,
            pady=(0, 12)
        )

        area_tabela = ctk.CTkFrame(
            janela,
            fg_color="#FFFFFF",
            corner_radius=12,
            border_width=1,
            border_color="#E5EAF2"
        )
        area_tabela.pack(
            fill="both",
            expand=True,
            padx=20,
            pady=(0, 20)
        )

        self._configurar_estilo_tabela(ttk.Style(janela))

        colunas = [f"coluna_{indice}" for indice in range(len(self.df.columns))]
        tabela = ttk.Treeview(
            area_tabela,
            columns=colunas,
            show="headings",
            style="DataVista.Treeview"
        )
        barra_vertical = ttk.Scrollbar(
            area_tabela,
            orient="vertical",
            command=tabela.yview,
            style="DataVista.Vertical.TScrollbar"
        )
        barra_horizontal = ttk.Scrollbar(
            area_tabela,
            orient="horizontal",
            command=tabela.xview,
            style="DataVista.Horizontal.TScrollbar"
        )
        tabela.configure(
            yscrollcommand=barra_vertical.set,
            xscrollcommand=barra_horizontal.set
        )

        for indice, coluna in enumerate(self.df.columns):
            identificador = colunas[indice]
            tabela.heading(identificador, text=str(coluna), anchor="w")
            tabela.column(
                identificador,
                width=min(max(len(str(coluna)) * 9 + 32, 110), 260),
                minwidth=90,
                anchor="w",
                stretch=True
            )

        for valores in self.df.head(100).itertuples(index=False, name=None):
            linha = [
                "" if pd.isna(valor) else str(valor)
                for valor in valores
            ]
            tabela.insert("", "end", values=linha)

        tabela.grid(row=0, column=0, sticky="nsew")
        barra_vertical.grid(row=0, column=1, sticky="ns")
        barra_horizontal.grid(row=1, column=0, sticky="ew")
        area_tabela.grid_rowconfigure(0, weight=1)
        area_tabela.grid_columnconfigure(0, weight=1)

        if self.modo_escuro:
            self._aplicar_paleta(
                janela,
                CORES_TEMA_ESCURO
            )

    # ========================================================
    # ESTILO DOS GRÁFICOS
    # ========================================================

    def _configurar_estilo_graficos(self):

        plt.rcParams.update({
            "font.family": "DejaVu Sans",
            "font.size": 10,
            "axes.titlesize": 14,
            "axes.labelsize": 10,
            "figure.facecolor": "white",
            "savefig.facecolor": "white"
        })


# ============================================================
# EXECUÇÃO DO PROGRAMA
# ============================================================

if __name__ == "__main__":

    root = ctk.CTk()

    app = AnalisadorExcel(root)

    root.mainloop()