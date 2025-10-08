import pandas as pd
from pathlib import Path
from datetime import datetime

def analisar_planilha(caminho_arquivo, relatorio):
    """Lê e analisa uma planilha Excel, escrevendo o resultado no relatório."""
    if not caminho_arquivo.exists():
        relatorio.write(f"\n❌ Arquivo não encontrado: {caminho_arquivo}\n")
        return

    try:
        df = pd.read_excel(caminho_arquivo)
    except Exception as e:
        relatorio.write(f"\nErro ao ler o arquivo {caminho_arquivo.name}: {e}\n")
        return

    relatorio.write(f"\n\n📂 === Analisando arquivo: {caminho_arquivo.name} ===\n")

    relatorio.write("\n📊 --- ANÁLISE BÁSICA ---\n")
    relatorio.write(f"➡ Linhas: {df.shape[0]}\n")
    relatorio.write(f"➡ Colunas: {df.shape[1]}\n")

    relatorio.write("\n🔍 --- NOMES DAS COLUNAS ---\n")
    relatorio.write(", ".join(map(str, df.columns)) + "\n")

    relatorio.write("\n📈 --- TIPOS DE DADOS ---\n")
    relatorio.write(df.dtypes.to_string() + "\n")

    relatorio.write("\n🧮 --- VALORES NULOS POR COLUNA ---\n")
    relatorio.write(df.isnull().sum().to_string() + "\n")

    relatorio.write("\n✨ --- AMOSTRA DE DADOS ---\n")
    relatorio.write(df.head().to_string() + "\n")

    relatorio.write("\n✅ Análise concluída com sucesso!\n")


if __name__ == "__main__":
    print("=== Analisador de Planilhas Excel ===")

    # Caminho automático para a pasta Downloads do usuário atual
    pasta_downloads = Path.home() / "Downloads"

    # Busca pelos dois arquivos na pasta Downloads
    arquivos = [
        pasta_downloads / "exemplo1.xlsx",
        pasta_downloads / "exemplo2.xlsx"
    ]

    print("\n📁 Verificando arquivos na pasta Downloads...")
    for caminho in arquivos:
        print(f" - {caminho}")

    # Nome do relatório com data/hora
    nome_relatorio = f"relatorio_analise_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    caminho_relatorio = Path(__file__).parent / nome_relatorio

    # Cria e grava o relatório
    with open(caminho_relatorio, "w", encoding="utf-8") as rel:
        rel.write("=== RELATÓRIO DE ANÁLISE DE PLANILHAS ===\n")
        rel.write(f"Data de geração: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
        rel.write("=" * 60 + "\n")

        for caminho in arquivos:
            analisar_planilha(caminho, rel)

    print(f"\n✅ Relatório salvo em: {caminho_relatorio.resolve()}")
