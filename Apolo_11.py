import pandas as pd
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import time
from selenium.webdriver.chrome.service import Service

# Leitura do arquivo Excel e criação do DataFrame
diretorio = 'C:/Users/.../Desktop/Material de Trabalho/SMEL/Equipamentos.xlsx'
armazenador = pd.read_excel(diretorio)
data_frame = pd.DataFrame(armazenador)

# Dicionário de mapeamento de valores
mapeamento_valores = {
    'Campo': '30',
    'Academia a Céu Aberto': '53',
    'Campo de Futebol Society': '60',
    'Ginásio': '5',
    'Praça': '26',
    'Pista de Skate': '49',
    'Piscina': '56',
    'Pista de Malha': '57',
    'Quadra Recreativa': '59',
    'Quadra Poliesportiva': '50',
    'Quadra de Vôlei': '58',
    'Quadra de Tênis': '54',
    'Quadra de Peteca': '55',
    'Quadra de Futsal': '52'
}

# VALOR LINHA DO DATA FRAME
valorLinha = 0

# CRIACAO DE UMA VARIAVEL PARA CONTAR QUANTAS UNIDADES FORAM VINCULADAS
cadastrados = 0

# CRIACAO DE UMA VARIAVEL PARA CONTAR QUANTAS UNIDADES NÃO FORAM VINCULADAS
naoCadastrados = 0

unidadesCadastradas = 0
ListaUnidade = []

# CASO OCORRA ALGUMA QUEBRA O ALGORITIMO RETOMA A CONTAGEM INICIANDO NO ULTIMO PONTO DE PARADA WHILE
pontoDeParadaWhile = 0

# ENQUANTO O PONTO DE PARA FOR DIFERENTE DE 2030, CONTINUE TENTANDO REALIZAR O CADASTRO


while (pontoDeParadaWhile != 140):

    for x in range(pontoDeParadaWhile, 140):

        # INDICADOR DE QUEBRA MOSTRANDO ONDE QUEBROU E ONDE DEVE RETORNAR:
        print("Estamos cadastrando a unidade:", pontoDeParadaWhile, "ao quebrar retorne a partir do: ",
              pontoDeParadaWhile)

        # PEGANDO OS DADOS DO DATA FRAME
        valorLinha = pontoDeParadaWhile

        servico = Service(ChromeDriverManager().install())
        navegador = webdriver.Chrome(service=servico)

        # ACESSAR O SITE DO CIC:
        navegador.get('Site privado, pois trata de informações confidenciais')

        try:
            # ISERIR LOGIN
            navegador.find_element(By.NAME, 'josso_username').send_keys('...')
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        try:
            # INSERIR A SENHA
            navegador.find_element(By.NAME, 'josso_password').send_keys('...')
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        try:
            # CLICAR NO BOTAO PARA ACESSAR SISTEMA
            navegador.find_element(By.CLASS_NAME, "botao").click()
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        try:
            # CLICAR NA OPCAO GERAL
            navegador.find_element('xpath', '//*[@id="geral"]/div[2]/ul/li[2]/a').click()
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        try:
            # CLICAR NA OPCAO UNIDADE
            navegador.find_element('xpath', '//*[@id="geral"]/div[2]/ul/li[2]/ul/li[2]/a').click()
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        try:
            # CLICAR NA OPCAO NOVA UNIDADE
            navegador.find_element('xpath', '//*[@id="novo"]').click()
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        # -------------------------------------------------------------------------

        nome = data_frame.at[valorLinha, 'NOME_EQUIPAMENTO']

        tipo = data_frame.at[valorLinha, 'TIPO_EQUIPAMENTO']

        log = data_frame.at[valorLinha, 'TIPO_LOGRADOURO']

        nomeLog = data_frame.at[valorLinha, 'NOME_LOGRADOURO']

        bairro = data_frame.at[valorLinha, 'BAIRRO']

        numlog = data_frame.at[valorLinha, 'NUMERO']

        time.sleep(5)
        # ---------------------------------------------------------------------------------

        try:
            # INSERIR A UNIDADE NO CAMPO NOME UNIDADE
            navegador.find_element('xpath', '//*[@id="mestre-nome"]').send_keys(nome)
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break
        time.sleep(3)

        try:
            # CLICAR NA OPCAO TITULARIDADE
            navegador.find_element('xpath', '//*[@id="mestre-titularidade"]/option[4]').click()
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break
        time.sleep(3)

        try:
            # CLICAR NA OPCAO TIPO DE UNIDADE
            navegador.find_element('xpath', '//*[@id="mestre-tipo_unidade"]').click()
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        time.sleep(5)

        try:
            # INSERIR O TIPO DE UNIDADE
            navegador.find_element('xpath', '//*[@id="mestre-tipo_unidade"]').send_keys(tipo)
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break
        time.sleep(5)

        try:
            # CLICAR NA OPCAO LOGRADOURO
            navegador.find_element('xpath', '//*[@id="aba-endereco"]').click()
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        time.sleep(5)

        try:
            # CLICAR NO BOTAO NOVA UNIDADE
            navegador.find_element('xpath', '//*[@id="novo"]').click()
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        try:
            # CLICANDO NO BOTAO PARA CADASTRAR NOVA UNIDADE
            navegador.find_element('xpath', '//*[@id="detalhe-1-vinculado"]').click()
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break
        time.sleep(2)

        try:
            # ENTRANDO NO I-FRAME PARA INSERIR DOS DADOS
            navegador.switch_to.frame(0)
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        try:
            # INSERINDO O LOGRADOURO
            navegador.find_element(By.XPATH, '//*[@id="logradouro"]').send_keys(nomeLog)
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile, "retorne ao ponto:", valorLinha)

            break

        try:
            # INSERINDO O BAIRRO
            navegador.find_element('xpath', '//*[@id="bairro"]').send_keys(bairro)
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile, "retorne ao ponto:", valorLinha)

            break

        try:
            # TRANSFORMANDO O LOGRADOURO EM NUMERO
            numlog = int(numlog)
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        try:
            # INSERINDO O NUMERO DO LOGRADOURO NO IFRAME
            navegador.find_element('xpath', '//*[@id="numeroInicial"]').send_keys(numlog)  # numlog
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        try:
            # Limpeza do campo "numeroFinal"
            navegador.find_element('xpath', '//*[@id="numeroFinal"]').clear()
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        try:
            # Preencha o campo "numeroFinal" com o valor desejado (numlog)
            navegador.find_element('xpath', '//*[@id="numeroFinal"]').send_keys(numlog)
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        try:
            # CLICAR NA OPCAO PESQUISAR
            navegador.find_element('xpath', '//*[@id="pesquisar"]').click()
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        # Descanse por 30 segundos (ou o tempo necessário para a página carregar)
        time.sleep(5)

        try:
            # Verifique se o tipo de equipamento está no dicionário de mapeamento
            if tipo in mapeamento_valores:
                valor_para_preencher = mapeamento_valores[tipo]
                # Localize o campo apropriado no site (substitua o XPath correto)
                campo_input = navegador.find_element('xpath', '//*[@id="mestre-tipo_unidade"]')
                campo_input.send_keys(valor_para_preencher)
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        time.sleep(3)

        try:
            navegador.find_element('xpath', '//*[@id="conteudo"]/table/tbody/tr[2]').click()
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        time.sleep(3)

        try:
            # FINALIZANDO O IFRAME
            navegador.switch_to.default_content()
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        time.sleep(5)

        try:
            navegador.find_element('xpath', '// *[ @ id = "detalhe-1-principal"]').click()
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        time.sleep(3)

        try:
            # CLICANDO NA OPCAO GRAVAR
            navegador.find_element('xpath', '//*[@id="gravar"]').click()

            # --------RASPAGEM DE DADOS ---------------------
            time.sleep(5)
            site = BeautifulSoup(navegador.page_source, "html.parser")
            buscador = site.find_all("p")
            unidadesCadastradas = buscador[1]
            ListaUnidade.append(unidadesCadastradas)
            print(unidadesCadastradas)
            # ------------------------------------------------

        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        time.sleep(3)
        pontoDeParadaWhile = pontoDeParadaWhile + 1
        try:
            navegador.quit()
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

for unidade in ListaUnidade:
    print(unidade)

print("Total Analisados", len(unidade))

