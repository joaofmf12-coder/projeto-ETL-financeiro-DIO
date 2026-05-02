import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# Fonte de Dados
url_github = "https://github.com/joaofmf12-coder/projeto-ETL-financeiro-DIO/raw/refs/heads/main/Controle%20Financeiro.xlsx"

# 1. EXTRACT
recebimentos = pd.read_excel(url_github, sheet_name='Recebimentos')
gastos = pd.read_excel(url_github, sheet_name='Gastos')
sobras = pd.read_excel(url_github, sheet_name='Sobras')
investimentos = pd.read_excel(url_github, sheet_name='Investimentos')

# 2. TRANSFORM
    # a. Limpeza de dados nulos
df_sobras_clean = sobras.dropna()
df_invest_clean = investimentos.dropna()

    # b. Integração das tabelas (Merge)
df_analise = pd.merge(df_sobras_clean, df_invest_clean, on=['Usuário', 'Mês'], suffixes=('_Sobra', '_Invest'))

    # c. Criação de indicador: Eficiência Financeira
df_analise['%_Investido'] = (df_analise['Valor_Invest'] / df_analise['Valor_Sobra']) * 100
df_analise['%_Investido'] = df_analise['%_Investido'].round(2)

print("--- ANÁLISE DE EFICIÊNCIA: SOBRA VS INVESTIMENTO ---")
print(df_analise[['Usuário', 'Mês', 'Valor_Sobra', 'Valor_Invest', '%_Investido']])

    # d. Geração do Gráfico
x = np.arange(len(df_analise))
largura = 0.35
fig, ax = plt.subplots(figsize=(12, 7))

barra_saldo = ax.bar(x - largura/2, df_analise['Valor_Sobra'], largura, label='Saldo', color='skyblue')
barra_invest = ax.bar(x + largura/2, df_analise['Valor_Invest'], largura, label='Investimento', color='salmon')

ax.set_ylabel('Valores em R$')
ax.set_title('Comparativo Mensal por Usuário: Saldo vs Investimento')
ax.set_xticks(x)
etiquetas = [f"{u}\n{m}" for u, m in zip(df_analise['Usuário'], df_analise['Mês'])]
ax.set_xticklabels(etiquetas)
ax.legend()
ax.yaxis.grid(True, linestyle='--', alpha=0.7)
plt.tight_layout()

    # IMPORTANTE: Salvar a imagem ANTES de plt.show() para não perdê-la
plt.savefig("grafico_temp.png")
plt.show()

# 5. LOAD (Carregamento Consolidado no Excel)
nome_arquivo = "Relatorio_Financeiro_Final.xlsx"

    # Inserir imagens no Excel
with pd.ExcelWriter(nome_arquivo, engine='xlsxwriter') as writer:
        # Aba 1: Dados
    df_analise.to_excel(writer, sheet_name='Analise_Saldos', index=False)

        # Aba 2: Visualização
    workbook = writer.book
    worksheet_grafico = workbook.add_worksheet('Visualizacao_Grafica')
    worksheet_grafico.insert_image('B2', 'grafico_temp.png')

print(f"Sucesso! O arquivo '{nome_arquivo}' foi gerado com duas abas.")