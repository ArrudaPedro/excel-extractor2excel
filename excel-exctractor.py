import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

def interacao_pizzaria(caminho_entrada):
    # 1. Ler a planilha original
    try:
        df = pd.read_excel(caminho_entrada)
    except Exception as e:
        print(f"Erro ao ler o arquivo: {e}")
        return

    # Validar se as colunas necessárias existem (ajuste os nomes se sua planilha for diferente)
    colunas_necessarias = ['Produto', 'Tamanho', 'Ingredientes']
    for col in colunas_necessarias:
        if col not in df.columns:
            print(f"Erro: A coluna '{col}' não foi encontrada na planilha.")
            return

    # 2. Identificar e listar todos os sabores únicos (removendo duplicados)
    sabores_unicos = df['Produto'].dropna().unique()
    
    if len(sabores_unicos) == 0:
        print("Nenhum produto encontrado na planilha.")
        return

    print("\n--- SABORES ENCONTRADOS NA PLANILHA ---")
    for i, sabor in enumerate(sabores_unicos, start=1):
        print(f"[{i}] {sabor}")
    print("---------------------------------------")

    # 3. Interação com o usuário para seleção
    while True:
        try:
            escolha = int(input(f"Digite o número do sabor que deseja extrair (1 a {len(sabores_unicos)}): "))
            if 1 <= escolha <= len(sabores_unicos):
                sabor_selecionado = sabores_unicos[escolha - 1]
                break
            else:
                print("Número inválido. Escolha uma opção da lista.")
        except ValueError:
            print("Por favor, digite um número válido.")

    print(f"\nVocê selecionou: {sabor_selecionado}")
    
    # 4. Filtrar a planilha original com base na escolha
    df_filtrado = df[df['Produto'] == sabor_selecionado]
    df_final = df_filtrado[colunas_necessarias]

    # Nome do arquivo de saída baseado no sabor escolhido (limpando espaços e caracteres)
    nome_arquivo_saida = f"Planilha_{sabor_selecionado.replace(' ', '_')}.xlsx"

    # 5. Salvar e formatar a nova planilha
    with pd.ExcelWriter(nome_arquivo_saida, engine='openpyxl') as writer:
        df_final.to_excel(writer, index=False, sheet_name='Dados Extraídos')
        
        worksheet = writer.sheets['Dados Extraídos']

        # Estilização profissional do cabeçalho
        header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid') # Azul escuro
        header_font = Font(color='FFFFFF', bold=True, size=11)
        border_thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border_thin

        # Formatação das linhas de dados
        for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row):
            for cell in row:
                cell.border = border_thin
                cell.alignment = Alignment(vertical='center', wrap_text=True)

        # Ajuste inteligente de largura das colunas
        worksheet.column_dimensions['A'].width = 30 # Produto
        worksheet.column_dimensions['B'].width = 15 # Tamanho
        worksheet.column_dimensions['C'].width = 60 # Ingredientes

    print(f"\nSucesso! O arquivo '{nome_arquivo_saida}' foi gerado apenas com o sabor escolhido.")


interacao_pizzaria('FICHA TÉCNICA PIZZA.xlsx')