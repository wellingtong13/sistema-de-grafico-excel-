"""
DataVista | Qualidade & Indicadores — v2.0
Estudo didático: relação entre horas de estudo e nota final.
ATENÇÃO: os dados gerados são fictícios/simulados, não são evidência de alunos reais.
"""

from datetime import datetime
from pathlib import Path
import platform
import sys
from tkinter import filedialog, messagebox

import customtkinter as ctk
import matplotlib
matplotlib.use("TkAgg")
import numpy as np
import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure


APP = "DataVista | Qualidade & Indicadores"
VERSAO = "2.0.0"
SEMENTE = 42

PALETA = {
    "navy": "#14243A",
    "navy2": "#1D3553",
    "red": "#D71920",
    "red_hover": "#A90F17",
    "blue": "#2374AB",
    "green": "#23966B",
    "orange": "#E59A32",
    "bg": "#F2F5F9",
    "text": "#172536",
    "muted": "#68788A",
    "line": "#DCE3EC"
}

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


# -------------------------- REQUISITOS / REGRAS ----------------------------

# RF01: gerar base simulada reproduzível.
# RF02: validar colunas, valores ausentes, IDs duplicados e limites.
# RF03: calcular estatística descritiva, Pearson e regressão linear.
# RF04: exibir painel, gráfico e prévia da base.
# RF05: exportar Excel com Dados, Indicadores e Rastreabilidade.

# RN01: nota entre 0 e 100; cansaço entre 1 e 10; horas entre 0 e 168.
# RN02: resultados estatísticos descrevem associação, não causalidade.
# RN03: dados simulados devem ser identificados como tais em toda saída.


# -------------------------- GERAÇÃO DE DADOS -------------------------------

def gerar_base(qtd=100, seed=SEMENTE):

    if qtd < 3:
        raise ValueError("A base precisa ter pelo menos 3 registros.")

    rng = np.random.default_rng(seed)

    horas = rng.uniform(1, 20, qtd)

    notas = np.clip(
        30 + 3.2 * horas + rng.normal(0, 8, qtd),
        0,
        100
    )

    cansaco = np.clip(
        rng.uniform(1, 10, qtd) + horas * 0.1,
        1,
        10
    )

    df = pd.DataFrame({
        "ID_Aluno": [
            f"ALU-{i:03d}" for i in range(1, qtd + 1)
        ],
        "Horas_Estudo_Semanal": np.round(horas, 2),
        "Nota_Final": np.round(notas, 2),
        "Nivel_Cansaco": np.round(cansaco, 2)
    })

    validar_base(df)

    return df


# -------------------------- VALIDAÇÃO DA BASE ------------------------------

def validar_base(df):

    colunas = [
        "ID_Aluno",
        "Horas_Estudo_Semanal",
        "Nota_Final",
        "Nivel_Cansaco"
    ]

    faltantes = [
        c for c in colunas if c not in df.columns
    ]

    if faltantes:
        raise ValueError(
            "Colunas obrigatórias ausentes: " +
            ", ".join(faltantes)
        )

    if df.empty:
        raise ValueError("A base está vazia.")

    if df[colunas].isnull().any().any():
        raise ValueError(
            "Há valores ausentes nas colunas obrigatórias."
        )

    for c in colunas[1:]:

        if not pd.api.types.is_numeric_dtype(df[c]):
            raise ValueError(
                f"A coluna {c} deve ser numérica."
            )

    if df["ID_Aluno"].duplicated().any():
        raise ValueError(
            "Há identificadores duplicados."
        )

    regras = {
        "Horas_Estudo_Semanal": (0, 168),
        "Nota_Final": (0, 100),
        "Nivel_Cansaco": (1, 10)
    }

    for c, (minimo, maximo) in regras.items():

        if not df[c].between(minimo, maximo).all():
            raise ValueError(
                f"Valores fora do limite permitido em {c}."
            )

    return [
        "Colunas obrigatórias presentes",
        "Ausência de valores nulos",
        "Identificadores únicos",
        "Tipos numéricos conformes",
        "Limites de negócio respeitados"
    ]


# -------------------------- ANÁLISE ESTATÍSTICA ----------------------------

def analisar(df):

    x = df["Horas_Estudo_Semanal"].astype(float)
    y = df["Nota_Final"].astype(float)

    m, b = np.polyfit(x, y, 1)

    estimado = m * x + b

    ss_res = ((y - estimado) ** 2).sum()
    ss_tot = ((y - y.mean()) ** 2).sum()

    r2 = 1 - ss_res / ss_tot if ss_tot else 0.0

    corr = float(x.corr(y))

    if corr >= .7:
        leitura = "Associação positiva forte."

    elif corr >= .3:
        leitura = "Associação positiva moderada."

    elif corr > 0:
        leitura = "Associação positiva fraca."

    elif corr <= -.7:
        leitura = "Associação negativa forte."

    elif corr <= -.3:
        leitura = "Associação negativa moderada."

    else:
        leitura = (
            "Associação linear fraca ou praticamente inexistente."
        )

    return {
        "Registros": len(df),
        "Média de horas": x.mean(),
        "Média das notas": y.mean(),
        "Mediana das notas": y.median(),
        "Desvio padrão": y.std(ddof=1),
        "Nota mínima": y.min(),
        "Nota máxima": y.max(),
        "Média de cansaço": df["Nivel_Cansaco"].mean(),
        "Correlação de Pearson": corr,
        "Inclinação": float(m),
        "Intercepto": float(b),
        "R²": float(r2),
        "Interpretação": leitura
    }


# -------------------------- CONSTRUÇÃO DO GRÁFICO --------------------------

def criar_grafico(df, ind):

    fig = Figure(
        figsize=(8.6, 5.2),
        dpi=100,
        facecolor="white"
    )

    ax = fig.add_subplot(111)

    p = ax.scatter(
        df["Horas_Estudo_Semanal"],
        df["Nota_Final"],
        s=df["Nivel_Cansaco"] * 23,
        c=df["Nota_Final"],
        cmap="viridis",
        alpha=.78,
        edgecolors="#263746",
        linewidths=.45
    )

    xx = np.linspace(
        df["Horas_Estudo_Semanal"].min(),
        df["Horas_Estudo_Semanal"].max(),
        100
    )

    ax.plot(
        xx,
        ind["Inclinação"] * xx + ind["Intercepto"],
        color=PALETA["red"],
        linestyle="--",
        linewidth=2,
        label="Linha de tendência"
    )

    ax.set_title(
        "Horas de estudo × nota final",
        fontsize=14,
        weight="bold",
        color=PALETA["navy"],
        pad=13
    )

    ax.set_xlabel("Horas de estudo por semana")
    ax.set_ylabel("Nota final (0 a 100)")
    ax.set_ylim(0, 105)

    ax.grid(
        True,
        linestyle=":",
        alpha=.35
    )

    ax.legend()

    cb = fig.colorbar(p, ax=ax)
    cb.set_label("Nota final")

    fig.tight_layout()

    return fig


# -------------------------- EXPORTAÇÃO EXCEL -------------------------------

def exportar_excel(caminho, df, ind, validacoes):

    resumo = pd.DataFrame(
        [(k, v) for k, v in ind.items()],
        columns=["Indicador", "Resultado"]
    )

    auditoria = pd.DataFrame([
        ("Sistema", APP),
        ("Versão", VERSAO),
        (
            "Data/hora",
            datetime.now().astimezone().isoformat(
                timespec="seconds"
            )
        ),
        ("Classificação da base", "SIMULADA — fins didáticos"),
        ("Semente aleatória", SEMENTE),
        ("Método", "Uniforme + ruído normal"),
        ("Python", sys.version.split()[0]),
        ("Sistema operacional", platform.platform()),

        *[
            (f"Validação {i+1}", v)
            for i, v in enumerate(validacoes)
        ]

    ], columns=["Campo", "Registro"])

    with pd.ExcelWriter(
        caminho,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            sheet_name="Dados",
            index=False
        )

        resumo.to_excel(
            writer,
            sheet_name="Indicadores",
            index=False
        )

        auditoria.to_excel(
            writer,
            sheet_name="Rastreabilidade",
            index=False
        )

        from openpyxl.styles import Font, PatternFill

        for ws in writer.book.worksheets:

            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            for cells in ws.columns:

                letra = cells[0].column_letter

                ws.column_dimensions[letra].width = min(
                    max(
                        len(str(c.value or ""))
                        for c in cells
                    ) + 3,
                    55
                )

            for cell in ws[1]:

                cell.font = Font(
                    bold=True,
                    color="FFFFFF"
                )

                cell.fill = PatternFill(
                    "solid",
                    fgColor="14243A"
                )


# -------------------------- APLICAÇÃO PRINCIPAL ----------------------------

class DataVista(ctk.CTk):

    def __init__(self):

        super().__init__()

        self.title(APP)
        self.geometry("1340x840")
        self.minsize(1080, 700)

        self.configure(
            fg_color=PALETA["bg"]
        )

        self.df = gerar_base()
        self.ind = analisar(self.df)
        self.validacoes = validar_base(self.df)

        self.canvas = None

        self.status = ctk.StringVar(
            value="Base simulada carregada e validada."
        )

        self._montar()
        self._atualizar()

    # -------------------------- INTERFACE ----------------------------------

    def _montar(self):

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # MENU LATERAL

        side = ctk.CTkFrame(
            self,
            width=220,
            corner_radius=0,
            fg_color=PALETA["navy"]
        )

        side.grid(
            row=0,
            column=0,
            sticky="nsew"
        )

        side.grid_propagate(False)

        ctk.CTkLabel(
            side,
            text="DATAVISTA",
            text_color="white",
            font=ctk.CTkFont(
                size=24,
                weight="bold"
            )
        ).pack(
            anchor="w",
            padx=22,
            pady=(28, 2)
        )

        ctk.CTkLabel(
            side,
            text="QUALIDADE • INDICADORES",
            text_color="#B9C8D8",
            font=ctk.CTkFont(
                size=10,
                weight="bold"
            )
        ).pack(
            anchor="w",
            padx=23,
            pady=(0, 30)
        )

        botoes = [
            ("▦  Painel executivo", self._atualizar),
            ("▤  Visualizar dados", self._ver_dados),
            ("✓  Checklist de validação", self._ver_validacao),
            ("⇩  Exportar Excel", self._exportar)
        ]

        for label, cmd in botoes:

            ctk.CTkButton(
                side,
                text=label,
                command=cmd,
                anchor="w",
                fg_color="transparent",
                hover_color="#29435F",
                text_color="white",
                height=42
            ).pack(
                fill="x",
                padx=12,
                pady=4
            )

        ctk.CTkFrame(
            side,
            height=1,
            fg_color="#3B5069"
        ).pack(
            fill="x",
            padx=18,
            pady=24
        )

        ctk.CTkLabel(
            side,
            text=(
                "CONTROLE DO ESTUDO\n"
                "Origem: dados simulados\n"
                "Semente: 42\n"
                "Registros: 100\n\n"
                "VERSÃO " + VERSAO
            ),
            text_color="#D7E1EB",
            justify="left",
            font=ctk.CTkFont(size=11)
        ).pack(
            anchor="w",
            padx=22
        )

        # ÁREA PRINCIPAL

        main = ctk.CTkFrame(
            self,
            fg_color="transparent"
        )

        main.grid(
            row=0,
            column=1,
            sticky="nsew",
            padx=25,
            pady=22
        )

        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(3, weight=1)

        ctk.CTkLabel(
            main,
            text="Painel de desempenho acadêmico",
            text_color=PALETA["navy"],
            font=ctk.CTkFont(
                size=25,
                weight="bold"
            )
        ).grid(
            row=0,
            column=0,
            sticky="w"
        )

        ctk.CTkLabel(
            main,
            text=(
                "Análise estatística • consistência • rastreabilidade"
            ),
            text_color=PALETA["muted"]
        ).grid(
            row=0,
            column=0,
            sticky="w",
            pady=(39, 12)
        )

        # AVISO SOBRE DADOS SIMULADOS

        aviso = ctk.CTkFrame(
            main,
            fg_color="#FFF4E5",
            border_color="#F1D3A5",
            border_width=1
        )

        aviso.grid(
            row=1,
            column=0,
            sticky="ew",
            pady=(0, 15)
        )

        ctk.CTkLabel(
            aviso,
            text=(
                "ATENÇÃO: dados fictícios/simulados para demonstração. "
                "Não representam registros reais nem constituem evidência de auditoria."
            ),
            text_color="#80520C",
            font=ctk.CTkFont(
                size=12,
                weight="bold"
            ),
            wraplength=1000
        ).pack(
            anchor="w",
            padx=13,
            pady=10
        )

        # CARDS DE INDICADORES

        self.kpi_area = ctk.CTkFrame(
            main,
            fg_color="transparent"
        )

        self.kpi_area.grid(
            row=2,
            column=0,
            sticky="ew",
            pady=(0, 15)
        )

        for i in range(4):
            self.kpi_area.grid_columnconfigure(
                i,
                weight=1
            )

        self.kpis = {}

        indicadores = [
            ("REGISTROS", "Registros", PALETA["blue"]),
            ("NOTA MÉDIA", "Média das notas", PALETA["green"]),
            ("HORAS MÉDIAS", "Média de horas", PALETA["red"]),
            ("CORRELAÇÃO", "Correlação de Pearson", PALETA["orange"])
        ]

        for i, (titulo, chave, cor) in enumerate(indicadores):

            card = ctk.CTkFrame(
                self.kpi_area,
                fg_color="white",
                corner_radius=12,
                border_width=1,
                border_color=PALETA["line"]
            )

            card.grid(
                row=0,
                column=i,
                sticky="ew",
                padx=(0 if i == 0 else 8, 0)
            )

            ctk.CTkLabel(
                card,
                text=titulo,
                text_color=PALETA["muted"],
                font=ctk.CTkFont(
                    size=10,
                    weight="bold"
                )
            ).pack(
                anchor="w",
                padx=14,
                pady=(13, 3)
            )

            lab = ctk.CTkLabel(
                card,
                text="—",
                text_color=cor,
                font=ctk.CTkFont(
                    size=24,
                    weight="bold"
                )
            )

            lab.pack(
                anchor="w",
                padx=14,
                pady=(0, 13)
            )

            self.kpis[chave] = lab

        # ÁREA DO GRÁFICO E INDICADORES

        body = ctk.CTkFrame(
            main,
            fg_color="transparent"
        )

        body.grid(
            row=3,
            column=0,
            sticky="nsew"
        )

        body.grid_columnconfigure(0, weight=3)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)

        chart_card = ctk.CTkFrame(
            body,
            fg_color="white",
            corner_radius=12,
            border_width=1,
            border_color=PALETA["line"]
        )

        chart_card.grid(
            row=0,
            column=0,
            sticky="nsew",
            padx=(0, 12)
        )

        chart_card.grid_rowconfigure(1, weight=1)
        chart_card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            chart_card,
            text="ANÁLISE GRÁFICA",
            text_color=PALETA["navy"],
            font=ctk.CTkFont(
                size=13,
                weight="bold"
            )
        ).grid(
            row=0,
            column=0,
            sticky="w",
            padx=16,
            pady=12
        )

        self.chart_area = ctk.CTkFrame(
            chart_card,
            fg_color="white"
        )

        self.chart_area.grid(
            row=1,
            column=0,
            sticky="nsew",
            padx=8,
            pady=5
        )

        # PAINEL DE LEITURA

        insight = ctk.CTkFrame(
            body,
            fg_color="white",
            corner_radius=12,
            border_width=1,
            border_color=PALETA["line"]
        )

        insight.grid(
            row=0,
            column=1,
            sticky="nsew"
        )

        ctk.CTkLabel(
            insight,
            text="INDICADORES E LEITURA",
            text_color=PALETA["navy"],
            font=ctk.CTkFont(
                size=12,
                weight="bold"
            )
        ).pack(
            anchor="w",
            padx=13,
            pady=13
        )

        self.texto = ctk.CTkTextbox(
            insight,
            wrap="word",
            fg_color="#F7F9FC",
            text_color=PALETA["text"],
            font=ctk.CTkFont(size=12)
        )

        self.texto.pack(
            fill="both",
            expand=True,
            padx=10,
            pady=(0, 10)
        )

        # RODAPÉ

        rodape = ctk.CTkFrame(
            main,
            fg_color="transparent"
        )

        rodape.grid(
            row=4,
            column=0,
            sticky="ew",
            pady=(12, 0)
        )

        ctk.CTkLabel(
            rodape,
            textvariable=self.status,
            text_color=PALETA["muted"]
        ).pack(side="left")

        ctk.CTkButton(
            rodape,
            text="Regenerar base",
            command=self._nova_base,
            fg_color=PALETA["red"],
            hover_color=PALETA["red_hover"],
            height=34
        ).pack(
            side="right",
            padx=(8, 0)
        )

        ctk.CTkButton(
            rodape,
            text="Exportar Excel",
            command=self._exportar,
            fg_color=PALETA["navy2"],
            height=34
        ).pack(side="right")

    # -------------------------- ATUALIZAÇÃO DO PAINEL ----------------------

    def _atualizar(self):

        self.ind = analisar(self.df)

        for chave, label in self.kpis.items():

            v = self.ind[chave]

            if chave == "Média das notas":
                txt = f"{v:.1f}/100"

            elif chave == "Média de horas":
                txt = f"{v:.1f} h"

            elif chave == "Correlação de Pearson":
                txt = f"{v:.3f}"

            else:
                txt = str(v)

            label.configure(text=txt)

        if self.canvas:
            self.canvas.get_tk_widget().destroy()

        fig = criar_grafico(
            self.df,
            self.ind
        )

        self.canvas = FigureCanvasTkAgg(
            fig,
            master=self.chart_area
        )

        self.canvas.draw()

        self.canvas.get_tk_widget().pack(
            fill="both",
            expand=True
        )

        i = self.ind

        conteudo = (
            f'MÉDIA DAS NOTAS\n{i["Média das notas"]:.2f}\n\n'
            f'MEDIANA\n{i["Mediana das notas"]:.2f}\n\n'
            f'DESVIO PADRÃO\n{i["Desvio padrão"]:.2f}\n\n'
            f'NOTA MÍNIMA / MÁXIMA\n'
            f'{i["Nota mínima"]:.2f} / {i["Nota máxima"]:.2f}\n\n'
            f'MÉDIA DE CANSAÇO\n{i["Média de cansaço"]:.2f}/10\n\n'
            f'REGRESSÃO LINEAR\n'
            f'Nota = {i["Intercepto"]:.2f} + '
            f'{i["Inclinação"]:.2f} × horas\n'
            f'R² = {i["R²"]:.3f}\n\n'
            f'INTERPRETAÇÃO\n{i["Interpretação"]}\n\n'
            'A correlação representa associação estatística; '
            'não comprova causa e efeito.'
        )

        self.texto.configure(state="normal")

        self.texto.delete(
            "1.0",
            "end"
        )

        self.texto.insert(
            "1.0",
            conteudo
        )

        self.texto.configure(state="disabled")

        self.status.set(
            "Painel atualizado • validações executadas."
        )

    # -------------------------- VISUALIZAR DADOS ---------------------------

    def _ver_dados(self):

        w = ctk.CTkToplevel(self)

        w.title("Base de dados — prévia")
        w.geometry("850x570")

        t = ctk.CTkTextbox(
            w,
            font=("Consolas", 11)
        )

        t.pack(
            fill="both",
            expand=True,
            padx=15,
            pady=15
        )

        t.insert(
            "1.0",
            self.df.to_string(index=False)
        )

        t.configure(state="disabled")

    # -------------------------- CHECKLIST DE VALIDAÇÃO ---------------------

    def _ver_validacao(self):

        self.validacoes = validar_base(self.df)

        mensagem = (
            "CHECKLIST DE QUALIDADE\n\n" +
            "\n".join(
                "✓ " + x for x in self.validacoes
            )
        )

        mensagem += (
            "\n\nLimitação: estas verificações avaliam consistência "
            "estrutural e regras programadas. Não representam certificação, "
            "conformidade normativa ou aprovação formal de auditoria."
        )

        messagebox.showinfo(
            "Validação",
            mensagem
        )

    # -------------------------- GERAR NOVA BASE ----------------------------

    def _nova_base(self):

        self.df = gerar_base()

        self.ind = analisar(self.df)

        self.validacoes = validar_base(self.df)

        self._atualizar()

        messagebox.showinfo(
            "Base gerada",
            "Base simulada regenerada. A semente fixa mantém "
            "os resultados reproduzíveis."
        )

    # -------------------------- EXPORTAR RELATÓRIO -------------------------

    def _exportar(self):

        destino = filedialog.asksaveasfilename(
            title="Salvar relatório de auditoria",
            defaultextension=".xlsx",
            filetypes=[
                ("Planilha Excel", "*.xlsx")
            ],
            initialfile="DataVista_Relatorio.xlsx"
        )

        if not destino:
            return

        try:

            exportar_excel(
                destino,
                self.df,
                self.ind,
                self.validacoes
            )

            self.status.set(
                "Relatório Excel exportado: " +
                Path(destino).name
            )

            messagebox.showinfo(
                "Exportação concluída",
                "Arquivo salvo em:\n" + destino
            )

        except Exception as exc:

            messagebox.showerror(
                "Erro na exportação",
                str(exc)
            )


# -------------------------- EXECUÇÃO DO SISTEMA ----------------------------

if __name__ == "__main__":

    try:

        app = DataVista()
        app.mainloop()

    except Exception as erro:

        messagebox.showerror(
            "Erro ao iniciar o DataVista",
            str(erro)
        )

        raise