import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from datetime import datetime

# Guardar apenas a data (sem a hora)
data_atual = datetime.now().strftime('%d/%m/%Y')


def inicializar_navegador():
    options = webdriver.EdgeOptions()
    options.add_experimental_option("detach", True)  # Mantém o navegador aberto
    driver = webdriver.Edge(options=options)
    return driver


def acessar_site(driver):
    print("Acessando o site...")
    driver.get("http://inspecoesrapidas.cpfl.com.br/Inspecao/cadastro")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "fname")))

# Garantir que a coluna "DataInspecao" esteja no formato DD/MM/YYYY
# df["DataInspecao"] = df["DataInspecao"].dt.strftime("%d/%m/%Y")
def preencher_formulario(driver, row, matricula):
    try:
        print(f"Iniciando preenchimento para a inspeção {row['ID']}")

        data_inspecao = row["DataInspecao"]
        operacao_de_campo = row["OperacaoDeCampo"]
        area_de_operacao = row["AreaDeOperacao"]
        ea = row["EA"]
        municipio = row["Municipio"]
        endereco = row["Endereco"] if pd.notna(row["Endereco"]) else municipio
        numero = row["Numero"] if pd.notna(row["Numero"]) else ""
        bairro = row["Bairro"] if pd.notna(row["Bairro"]) else ""
        alimentador = row["Alimentador"]
        numero_operativo = row["Numero_Operativo"]
        tronco_ramal = row["Tronco/Ramal"]
        executor = row["Executor"]
        defeito = row["Defeito"]
        nivel = row['Nivel']
        quantidade = row["Quantidade"]

        # Preencher a matrícula
        campo_matricula = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "fname")))

        campo_matricula.clear()
        campo_matricula.send_keys(matricula)
        campo_matricula.send_keys(Keys.TAB)

        # Selecionar empresa
        dropdown_empresa = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "dropdown_empresa")))
        dropdown_empresa.click()
        empresa_option = driver.find_element(By.XPATH,
                                             "//input[@class='dropdown-item empresa' and @value='CPFL PAULISTA']")
        empresa_option.click()

        # Preencher a data
        campo_data = driver.find_element(By.ID, "DataInspecao")
        campo_data.clear()
        campo_data.send_keys(data_atual)
        campo_data.send_keys(Keys.TAB)

        # Selecionar operação de campo
        dropdown_operacao = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "dropdown_operacaocampo")))
        dropdown_operacao.click()
        operacao_option = driver.find_element(By.XPATH,
                                              f"//input[@class='dropdown-item operacaocampo' and @value='{operacao_de_campo}']")
        operacao_option.click()

        # Selecionar área de operação
        dropdown_area_operacao = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "dropdown_areaoperacaocampo")))
        dropdown_area_operacao.click()
        area_operacao_option = driver.find_element(By.XPATH,
                                                   f"//input[@class='dropdown-item areaoperacaocampo' and @value='{area_de_operacao}']")
        area_operacao_option.click()

        # Selecionar EA
        dropdown_ea = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "dropdown_ea")))
        dropdown_ea.click()
        ea_option = driver.find_element(By.XPATH, f"//input[@class='dropdown-item ea' and @value='{ea}']")
        ea_option.click()

        # Selecionar município
        dropdown_municipio = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "dropdown_municipio")))
        dropdown_municipio.click()
        municipio_option = driver.find_element(By.XPATH,
                                               f"//input[@class='dropdown-item municipio' and @value='{municipio}']")
        municipio_option.click()

        # Aguarde até que o botão de "Endereço da Inspeção" esteja visível e clicável
        btn_equipes = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "nav-equipes-tab")))
        btn_equipes.click()
        time.sleep(2)  # Pequeno delay para garantir carregamento

        # Localiza o campo de endereço pelo ID e preenche
        input_endereco = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Logradouro")))
        input_endereco.clear()
        input_endereco.send_keys(endereco)

        # Preencher o campo "Número" com a coluna "Numero", se houver dado
        input_numero = driver.find_element(By.ID, "Numero")
        input_numero.clear()
        input_numero.send_keys(numero)

        # Preencher o campo "Bairro" com a coluna "Bairro", se houver dado
        input_bairro = driver.find_element(By.ID, "Bairro")
        input_bairro.clear()
        input_bairro.send_keys(bairro)

        if pd.notna(alimentador):  # Se houver um alimentador na planilha
            # Clicar no botão dropdown para alimentar
            dropdown_alimentador = driver.find_element(By.ID, "dropdown_alimentador")
            dropdown_alimentador.click()

            # Esperar que os itens do dropdown carreguem
            alimentador_option = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, f"//input[@class='dropdown-item alimentador' and @value='{alimentador}']")))
            alimentador_option.click()

            # Preencher o campo "Operativo"
            input_operativo = driver.find_element(By.ID, "NumeroOperativo")
            input_operativo.clear()
            input_operativo.send_keys(numero_operativo)

            # Preencher o campo de observação, se existir
            observacao = row["Observacao"] if "Observacao" in row and pd.notna(row["Observacao"]) else ""

            descricao_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "DescricaoSolicitacao")))
            descricao_input.clear()
            descricao_input.send_keys(observacao)

            # Aguarde até que o botão de "Informações da Inspeção" esteja visível e clicável
            btn_informacoes = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "nav-administrador-tab")))
            btn_informacoes.click()
            time.sleep(2)  # Pequeno delay para garantir carregamento

            from selenium.webdriver.support.ui import Select

            # Selecionar Tronco/Ramal
            select_tronco_ramal = driver.find_element(By.ID, "IdTroncoRamal")
            if tronco_ramal == "Tronco":
                select_tronco_ramal.send_keys("Tronco")
            elif tronco_ramal == "Ramal":
                select_tronco_ramal.send_keys("Ramal")
            else:
                select_tronco_ramal.send_keys("")  # Deixar em branco se não houver valor

            # Preencher o campo Executor
            dropdown_executor = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "IdExecutor")))

            # Usando Select para manipular o dropdown de forma correta
            select_executor = Select(dropdown_executor)

            # Mapeamento de valores de Executor para os valores no select
            executores_map = {
                "LM-CPFL": "2",
                "LM-PARCEIRA": "4",
                "LV-CPFL": "1",
                "LV-PARCEIRA": "3"
            }

            if executor in executores_map:
                valor_executor = executores_map[executor]

                # Seleciona a opção com o valor correto
                select_executor.select_by_value(valor_executor)
            else:
                print(f"Executor não encontrado para o valor: {executor}")

            # Supondo que a severidade esteja na coluna 'Severidade' e que a primeira linha seja cabeçalho
            severidade = df['Severidade'].iloc[0]  # Aqui você pega o primeiro valor da coluna Severidade

            # Encontrar o dropdown de Severidade
            dropdown_severidade = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "IdSeveridade")))

            # Usar o Select para manipular o dropdown
            select_severidade = Select(dropdown_severidade)

            # Mapeamento de valores de Severidade
            severidade_map = {
                "Alta": "1",
                "Média": "2",
                "Baixa": "3"
            }

            # Verificar se a severidade extraída do Excel existe no mapeamento e selecionar a opção correspondente
            if severidade in severidade_map:
                valor_severidade = severidade_map[severidade]

                # Selecionar a opção com o valor correto
                select_severidade.select_by_value(valor_severidade)
            else:
                print(f"Severidade não encontrada para o valor: {severidade}")

            dropdown_nivel = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "IdNivel")))

            # Usar o Select para manipular o dropdown
            select_nivel = Select(dropdown_nivel)

            # Mapeamento de valores de Nivel
            nivel_map = {
                "Primário": "1",
                "Secundário": "2"
            }

            # Verificar se o nível extraído do Excel existe no mapeamento e selecionar a opção correspondente
            if nivel in nivel_map:
                valor_nivel = nivel_map[nivel]

                # Selecionar a opção com o valor correto
                select_nivel.select_by_value(valor_nivel)

            dropdown_defeito_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "dropdown_defeito")))
            dropdown_defeito_button.click()

            # Agora, procurar o botão de defeito correspondente dentro do menu
            defeito_buttons = driver.find_elements(By.CSS_SELECTOR, "input.dropdown-item.defeito")

            # Verificar se o defeito extraído do Excel está presente e clicar no botão correspondente
            defeito_encontrado = False
            for button in defeito_buttons:
                if button.get_attribute("value") == defeito:
                    button.click()
                    defeito_encontrado = True
                    break

            if not defeito_encontrado:
                print(f"Defeito '{defeito}' não encontrado na lista de opções.")

            campo_quantidade = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Quantidade")))

            campo_quantidade.clear()  # Limpa o campo antes de preencher
            campo_quantidade.send_keys(str(quantidade))  # Preenche com a quantidade correta da linha atual

            # Localizar o botão 'Adicionar' e clicar
            botao_adicionar = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, "adicionar")))
            botao_adicionar.click()

            # Espera o botão "Salvar" ficar clicável e então clica nele
            salvar_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.btn-md.m-1.btn-success[type='submit']"))
            )
            salvar_button.click()

            # ATÉ TER CERTEZA QUE NÃO VAI REPETIR CÓDIGO ESSA FUNÇÃO ESTÁ COMENTADA

            # Espera o link "Sim" ficar clicável e então clica nele
            confirmar_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "lnkConfirmSim"))
            )
            confirmar_button.click()

            print(f"Inspeção {row['ID']} preenchida com sucesso!")

    except Exception as e:
        print(f"Erro ao preencher o formulário para a inspeção {row['ID']}: {str(e)}")