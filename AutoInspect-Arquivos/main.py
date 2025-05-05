import os
import pandas as pd
from openpyxl import load_workbook
from automacao import preencher_formulario, inicializar_navegador, acessar_site

def rodar_automacao(matricula, arquivo_excel):
    from automacao import preencher_formulario, inicializar_navegador, acessar_site

    # Carregar o dataframe diretamente do arquivo recebido
    df = pd.read_excel(arquivo_excel)

    # Inicializar o navegador
    driver = inicializar_navegador()

    # Verificar se a coluna "Inspecao_Lancada" existe no DataFrame, se não, criar
    if "Inspecao_Lancada" not in df.columns:
        df["Inspecao_Lancada"] = ""  # Cria a coluna com valor vazio caso não exista

    # Abrir o arquivo Excel existente com openpyxl
    wb = load_workbook(arquivo_excel)

    # Obter a planilha ativa
    sheet = wb.active

    # Loop para preencher os dados da planilha
    for index, row in df.iterrows():
        print(f"Verificando para matrícula {matricula} - Linha {index + 1}")

        if str(row.get("Inspecao_Lancada", "")).strip().upper() == "OK":
            print(f"Inspeção {row['ID']} já lançada. Pulando.")
            continue  # Se já foi lançada, pula essa linha

        # Reabre o site para lançar a nova inspeção
        acessar_site(driver)

        # Caso não tenha sido lançada, preenche o formulário
        print(f"Preenchendo para matrícula {matricula} - Linha {index + 1}")
        preencher_formulario(driver, row, matricula)

        # Marca a inspeção como "OK" após o preenchimento do formulário
        df.at[index, "Inspecao_Lancada"] = "OK"
        print(f"Inspeção {row['ID']} lançada com sucesso!")

        # Fecha o navegador após cada inspeção
        driver.quit()

        # Reabre o navegador para a próxima inspeção
        driver = inicializar_navegador()

    # Agora, após modificar a coluna Inspecao_Lancada, vamos atualizar a planilha no Excel.
    with pd.ExcelWriter(arquivo_excel, engine='openpyxl') as writer:
        writer._book = wb
        df.to_excel(writer, index=False, sheet_name=sheet.title)

    driver.quit()

    print("Processo finalizado e Excel atualizado, mantendo a formatação e as planilhas!")