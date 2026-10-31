from io import BytesIO
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="DataVista | Analisador Excel",
    page_icon="📊",
    layout="wide",
)


def preparar_dados(
    df: pd.DataFrame,
    coluna_x: str,
    coluna_y: str | None = None,
) -> pd.DataFrame:
    colunas = [coluna_x] if coluna_y is None else [coluna_x, coluna_y]
    dados = df[colunas].copy()

    for coluna in colunas:
        dados[coluna] = pd.to_numeric(dados[coluna], errors="coerce")

    return dados.replace([np.inf, -np.inf], np.nan).dropna()


def criar_grafico(
    dados: pd.DataFrame,
    tipo: str,
    coluna_x: str,
    coluna_y: str | None,
):
    figura, ax = plt.subplots(figsize=(10, 5.5))
    cor_principal = "#E1262F"
    cor_secundaria = "#237BFA"

    if tipo == "Dispersão":
        ax.scatter(
            dados[coluna_x],
            dados[coluna_y],
            color=cor_principal,
            alpha=0.78,
            edgecolors="white",
            linewidths=0.5,
            label="Observações",
        )
        if len(dados) >= 2 and dados[coluna_x].nunique() > 1:
            coeficientes = np.polyfit(dados[coluna_x], dados[coluna_y], 1)
            ordem = dados[coluna_x].sort_values()
            ax.plot(
                ordem,
                np.polyval(coeficientes, ordem),
                color=cor_secundaria,
                linewidth=2,
                label="Tendência linear",
            )
        ax.set_xlabel(coluna_x)
        ax.set_ylabel(coluna_y)
        ax.legend(frameon=False)
    elif tipo == "Linha":
        ordenados = dados.sort_values(coluna_x)
        ax.plot(
            ordenados[coluna_x],
            ordenados[coluna_y],
            color=cor_principal,
            marker="o",
            markersize=3,
            linewidth=2,
        )
        ax.set_xlabel(coluna_x)
        ax.set_ylabel(coluna_y)
    elif tipo == "Barras":
        amostra = dados.head(30)
        ax.bar(
            amostra[coluna_x].astype(str),
            amostra[coluna_y],
            color=cor_secundaria,
        )
        ax.set_xlabel(coluna_x)
        ax.set_ylabel(coluna_y)
        ax.tick_params(axis="x", rotation=35)
        if len(dados) > 30:
            ax.set_title("Exibindo os 30 primeiros registros")
    else:
        ax.hist(
            dados[coluna_x],
            bins="auto",
            color=cor_secundaria,
            edgecolor="white",
        )
        ax.set_xlabel(coluna_x)
        ax.set_ylabel("Frequência")

    ax.set_title("Visualização dos dados", loc="left", fontweight="bold")
    ax.grid(axis="y", linestyle="--", alpha=0.25)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    figura.tight_layout()
    return figura


def main() -> None:
    st.title("DataVista | Analisador Excel")
    st.write(
        "Importe uma planilha Excel para visualizar os dados e criar gráficos."
    )

    logo = Path(__file__).resolve().parent / "assets" / "senai_logo.png"
    if logo.is_file():
        st.sidebar.image(str(logo), width=100)

    arquivo = st.file_uploader(
        "Selecione uma planilha Excel (.xlsx ou .xls)",
        type=["xlsx", "xls"],
    )
    if arquivo is None:
        st.info("Envie uma planilha para começar a análise.")
        return

    try:
        df = pd.read_excel(arquivo)
    except (ImportError, ValueError, OSError) as erro:
        st.error(f"Não foi possível abrir a planilha: {erro}")
        return

    if df.empty:
        st.error("A planilha está vazia.")
        return

    colunas_numericas = df.select_dtypes(include="number").columns.tolist()
    if not colunas_numericas:
        st.error("Não encontrei colunas numéricas nessa planilha.")
        return

    st.caption(f"Arquivo carregado: **{arquivo.name}**")
    metricas = st.columns(3)
    metricas[0].metric("Registros", f"{len(df):,}".replace(",", "."))
    metricas[1].metric("Colunas", len(df.columns))
    metricas[2].metric("Colunas numéricas", len(colunas_numericas))

    st.subheader("Pré-visualização")
    st.dataframe(df.head(100), use_container_width=True)
    st.caption(f"Mostrando até 100 linhas de {len(df):,} registros.")

    st.subheader("Monte sua visualização")
    controles = st.columns(3)
    tipo = controles[0].selectbox(
        "Tipo de gráfico",
        ["Dispersão", "Linha", "Barras", "Histograma"],
    )
    coluna_x = controles[1].selectbox("Eixo X", colunas_numericas)
    coluna_y = controles[2].selectbox(
        "Eixo Y",
        colunas_numericas,
        index=1 if len(colunas_numericas) > 1 else 0,
        disabled=tipo == "Histograma",
    )

    if tipo != "Histograma" and coluna_x == coluna_y:
        st.warning("Escolha colunas diferentes para os eixos X e Y.")
        return

    coluna_y_grafico = None if tipo == "Histograma" else coluna_y
    dados = preparar_dados(df, coluna_x, coluna_y_grafico)
    if dados.empty:
        st.error("Não há valores numéricos válidos para montar este gráfico.")
        return
    if tipo == "Histograma" and len(dados) < 2:
        st.error("O histograma precisa de pelo menos dois valores válidos.")
        return

    figura = criar_grafico(dados, tipo, coluna_x, coluna_y_grafico)
    imagem = BytesIO()
    figura.savefig(imagem, format="png", dpi=180, bbox_inches="tight")
    st.pyplot(figura)
    plt.close(figura)

    st.caption(f"{len(dados):,} observações válidas.".replace(",", "."))
    st.download_button(
        "Baixar gráfico em PNG",
        data=imagem.getvalue(),
        file_name=f"grafico_{tipo.lower()}.png",
        mime="image/png",
    )


if __name__ == "__main__":
    main()
