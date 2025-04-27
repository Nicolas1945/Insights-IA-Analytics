import sys
import subprocess
import logging
from datetime import datetime
import os

try:
    import pkg_resources
    required = {'pandas', 'matplotlib', 'seaborn', 'fpdf2'}
    installed = {pkg.key for pkg in pkg_resources.working_set}
    missing = required - installed

    if missing:
        print(f"Instalando pacotes faltantes: {missing}")
        subprocess.check_call([sys.executable, "-m", "pip", "install", *missing])
    
    import pandas as pd
    import matplotlib.pyplot as plt
    import seaborn as sns
    from fpdf import FPDF

except Exception as e:
    print(f"Erro ao configurar ambiente: {str(e)}")
    print("Por favor, instale manualmente os pacotes com:")
    print("pip install pandas matplotlib seaborn fpdf2")
    sys.exit(1)

# Configuração do logger
logger = logging.getLogger('meganium_analysis')
logger.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

file_handler = logging.FileHandler('meganium_analysis.log')
file_handler.setFormatter(formatter)

console_handler = logging.StreamHandler()
console_handler.setFormatter(formatter)

logger.addHandler(file_handler)
logger.addHandler(console_handler)

class PDFReport(FPDF):
    def __init__(self):
        super().__init__()
        self.set_auto_page_break(auto=True, margin=15)
        self.add_page()
        self.set_font('Arial', '', 12)
    
    def header(self):
        self.set_font('Arial', 'B', 15)
        self.cell(0, 10, 'Relatório de Vendas - Meganium Games', 0, 1, 'C')
        self.ln(5)
    
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')
    
    def add_insights_section(self, title, insights):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, title, 0, 1)
        self.set_font('Arial', '', 10)
        
        for key, value in insights.items():
            self.cell(0, 6, f"- {key}: {value}", 0, 1)
        
        self.ln(5)
    
    def add_image_section(self, title, image_path):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, title, 0, 1)
        self.image(image_path, x=10, w=190)
        self.ln(5)

def load_data():
    """Carrega e prepara os dados"""
    try:
        logger.info("Carregando dados do arquivo CSV...")
        df = pd.read_csv('data/Meganium_Sales_Data.csv')
        
        required_columns = ['product_sold', 'date', 'quantity', 'total_price', 
                          'currency', 'site', 'discount_value', 'delivery_country']
        
        missing_cols = [col for col in required_columns if col not in df.columns]
        if missing_cols:
            raise ValueError(f"Colunas obrigatórias faltando: {missing_cols}")
        
        logger.info("Convertendo dados de data...")
        df['date'] = pd.to_datetime(df['date'], errors='coerce')
        invalid_dates = df['date'].isna().sum()
        if invalid_dates > 0:
            logger.warning(f"Removendo {invalid_dates} registros com datas inválidas")
            df = df.dropna(subset=['date'])
        
        df['month'] = df['date'].dt.month
        df['day_of_week'] = df['date'].dt.day_name()
        df['year'] = df['date'].dt.year
        
        logger.info(f"Dados carregados com sucesso. Total de registros: {len(df)}")
        return df
    
    except Exception as e:
        logger.error(f"Erro ao carregar dados: {str(e)}", exc_info=True)
        raise

def generate_insights(df):
    """Gera insights analíticos"""
    logger.info("Gerando insights dos dados...")
    
    insights = {
        'Produto mais vendido': df['product_sold'].value_counts().idxmax(),
        'Quantidade total vendida': f"{int(df['quantity'].sum()):,} unidades",
        'Receita total': f"{df['total_price'].sum():,.2f} {df['currency'].iloc[0]}",
        'Ticket médio': f"{df['total_price'].mean():,.2f} {df['currency'].iloc[0]}",
        'Total de descontos': f"{df['discount_value'].sum():,.2f} {df['currency'].iloc[0]}",
        'País com mais vendas': df['delivery_country'].value_counts().idxmax(),
        'Site com mais vendas': df['site'].value_counts().idxmax(),
        'Mês com mais vendas': df.groupby('month')['total_price'].sum().idxmax(),
        'Dia da semana com mais vendas': df.groupby('day_of_week')['total_price'].sum().idxmax(),
        'Ano com mais vendas': df.groupby('year')['total_price'].sum().idxmax()
    }
    
    logger.info("Insights gerados com sucesso")
    return insights

def create_visualizations(df):
    """Cria visualizações e salva como imagens"""
    logger.info("Criando visualizações...")
    os.makedirs('output', exist_ok=True)
    
    try:
        # Gráfico 1: Top 10 produtos
        plt.figure(figsize=(10, 6))
        sns.barplot(
            y=df['product_sold'].value_counts().head(10).index,
            x=df['product_sold'].value_counts().head(10).values,
            palette='viridis'
        )
        plt.title('Top 10 Produtos Mais Vendidos', pad=20)
        plt.xlabel('Quantidade Vendida')
        plt.tight_layout()
        plt.savefig('output/top_produtos.png', dpi=300, bbox_inches='tight')
        plt.close()
        
        # Gráfico 2: Vendas por país
        plt.figure(figsize=(8, 8))
        top_countries = df['delivery_country'].value_counts().head(5)
        plt.pie(
            top_countries,
            labels=top_countries.index,
            autopct='%1.1f%%',
            startangle=90,
            colors=sns.color_palette('pastel')
        )
        plt.title('Distribuição de Vendas por País (Top 5)', pad=20)
        plt.savefig('output/vendas_paises.png', dpi=300, bbox_inches='tight')
        plt.close()
        
        # Gráfico 3: Vendas por mês
        plt.figure(figsize=(12, 6))
        monthly_sales = df.groupby('month')['total_price'].sum()
        sns.lineplot(
            x=monthly_sales.index,
            y=monthly_sales.values,
            marker='o',
            color='royalblue',
            linewidth=2.5
        )
        plt.title('Vendas Mensais', pad=20)
        plt.xlabel('Mês')
        plt.ylabel('Receita Total')
        plt.xticks(range(1, 13))
        plt.grid(True, linestyle='--', alpha=0.7)
        plt.tight_layout()
        plt.savefig('output/vendas_mes.png', dpi=300, bbox_inches='tight')
        plt.close()
        
        logger.info("Visualizações criadas com sucesso")
        
    except Exception as e:
        logger.error(f"Erro ao criar visualizações: {str(e)}", exc_info=True)
        raise

def generate_pdf_report(insights, output_path='Meganium_Sales_Report.pdf'):
    """Gera relatório em PDF com os insights"""
    logger.info("Gerando relatório PDF...")
    
    try:
        pdf = PDFReport()
        
        # Capa
        pdf.set_font('Arial', 'B', 16)
        pdf.cell(0, 40, '', 0, 1)  # Espaçamento
        pdf.cell(0, 10, 'Relatório Completo de Vendas', 0, 1, 'C')
        pdf.set_font('Arial', '', 12)
        pdf.cell(0, 10, 'Meganium Games - Análise de Dados', 0, 1, 'C')
        pdf.cell(0, 10, datetime.now().strftime('%d/%m/%Y'), 0, 1, 'C')
        pdf.add_page()
        
        # Seção de Insights
        pdf.add_insights_section('Principais Insights', insights)
        
        # Adicionar gráficos
        pdf.add_image_section('Top 10 Produtos Mais Vendidos', 'output/top_produtos.png')
        pdf.add_image_section('Distribuição de Vendas por País', 'output/vendas_paises.png')
        pdf.add_image_section('Vendas Mensais', 'output/vendas_mes.png')
        
        # Salvar PDF
        pdf.output(output_path)
        logger.info(f"Relatório PDF gerado com sucesso: {output_path}")
        
        return output_path
    
    except Exception as e:
        logger.error(f"Erro ao gerar PDF: {str(e)}", exc_info=True)
        raise

def main():
    try:
        logger.info("Iniciando análise de vendas Meganium Games")
        
        # Carregar e processar dados
        df = load_data()
        
        # Gerar insights
        insights = generate_insights(df)
        
        # Criar visualizações
        create_visualizations(df)
        
        # Gerar relatório PDF
        report_path = generate_pdf_report(insights)
        
        logger.info(f"Análise concluída com sucesso. Relatório disponível em: {report_path}")
        
        # Exibir resumo no console
        print("\n=== RESUMO DA ANÁLISE ===")
        for key, value in insights.items():
            print(f"{key}: {value}")
        print(f"\nRelatório completo gerado em: {report_path}")
        
    except Exception as e:
        logger.error(f"Falha na execução da análise: {str(e)}", exc_info=True)
        print("Ocorreu um erro durante a análise. Verifique o arquivo de log para detalhes.")

if __name__ == "__main__":
    main()