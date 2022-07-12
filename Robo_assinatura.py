from imap_tools import MailBox, AND
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.support.select import Select
from pynput.keyboard import Key, Controller
from credenciais import senha, login, caminho
from datetime import datetime, date
from time import sleep
import logging


# Ao iniciar o Google Chrome, o exemplar iniciado será uma página com o plugin
chrome_options = ChromeOptions()
chrome_options.add_extension(f'{caminho}plugincertisign.crx')
# logando no email, via IMAP
login = login()
senha = senha()
meu_email = MailBox('outlook.office365.com').login(login, senha)
timeout = 10
contador = 0

def simula():
    Controller().press(Key.alt)
    Controller().press(Key.tab)
    Controller().release(Key.tab)
    Controller().release(Key.alt)
    sleep(0.6)
    happen = datetime.now().strftime('%d/%m/%Y %H:%M')
    print(f"Selecionando allow no pluging em {happen}")
    Controller().press(Key.alt)
    Controller().press(Key.tab)
    Controller().release(Key.tab)
    Controller().release(Key.alt)
    sleep(0.6)
    Controller().tap(Key.tab)
    sleep(0.7)
    Controller().tap(Key.tab)
    sleep(0.7)
    Controller().tap(Key.enter)
    sleep(1.5)
    Controller().tap(Key.enter)
    sleep(14)


while True:
    try:
        atch = meu_email.fetch(AND(
            subject="Portal Vertsign informa: Assinatura pendente.", seen=False))
            # Verificando a caixa de entrada
        for msg in atch:
            logging.basicConfig(level=logging.INFO, filename=
            f'{caminho}{date.today()}.log', datefmt='%H:%M:%S',
                                filemode='a', format='%(asctime)s,%(msecs)d %(name)s %(levelname)s %(message)s')
            lista_email = msg
            posicao_inicio = lista_email.html.rfind(
                '<tr><td style="padding:0 5px"><a href="') + 39
            posicao_fim = lista_email.html[posicao_inicio:-1].find('" target="_blank">')
            # Abrindo o link
            driver = webdriver.Chrome(f'{caminho}chromedriver.exe', options=chrome_options)
            happened = datetime.now().strftime('%d/%m/%Y %H:%M')
            print(f'Abrindo em {happened}')
            logging.info(f'Abrindo em {happened}')
            try:
                driver.get(
                    lista_email.html[posicao_inicio:posicao_inicio + posicao_fim])
                # Clicando na opção de aceitar os cookies
                element_present = ec.presence_of_element_located((
                    By.CSS_SELECTOR, '#onetrust-accept-btn-handler'))
                WebDriverWait(driver, timeout).until(element_present).click()
                happened = datetime.now().strftime('%d/%m/%Y %H:%M')
                print(f'Aceitando Cookies em {happened}')
                logging.info(f'Aceitando Cookies em {happened}')
            except:
                driver.close()
                meu_email.move(lista_email.uid, 'Não assinados')
                happened = datetime.now().strftime('%d/%m/%Y %H:%M')
                print(f'Não achou Cookies em {happened}')
                logging.info(f'Não achou Cookies em {happened}')
                print(30 * "=")
                logging.info(30 * "=")
                break
            try:
                assinante = driver.find_element(by=By.CLASS_NAME, value='form-horizontal').text
                selec = Select(driver.find_element(by=By.ID, value='certificateSelect'))
                posiini = assinante.rfind('Signatário:\n ') + 13
                posifim = assinante.find('\nData:')
                nomeassin = assinante[posiini:posifim].upper()
                sleep(2)
                # Selecionando o assinante
                if "D. M. P. II" in nomeassin:
                    print(nomeassin)
                    logging.info(nomeassin)
                    selec.select_by_index(0)
                if "D. M. S." in nomeassin:
                    print(nomeassin)
                    logging.info(nomeassin)
                    selec.select_by_index(1)
                if "F. F. C." in nomeassin:
                    print(nomeassin)
                    logging.info(nomeassin)
                    selec.select_by_index(2)
            except:
                driver.close()
                meu_email.move(lista_email.uid, 'Não assinados')
                happened = datetime.now().strftime('%d/%m/%Y %H:%M')
                print(f'Não foi achado assinantes no select em {happened}')
                logging.info(f'Não foi achado assinantes no select em {happened}')
                print(30 * "=")
                logging.info(30 * "=")
                break
            try:
                elem = driver.find_element(by=By.CLASS_NAME, value='btn-warning')
                # Procura o botão de assinar
                sleep(1)
                if elem.is_displayed():
                    try:
                        sleep(1.5)
                        elem.click()
                        sleep(4.5)
                        simula()
                        driver.close()
                        meu_email.move(lista_email.uid, 'Assinados')
                        sleep(1)
                        happened = datetime.now().strftime('%d/%m/%Y %H:%M')
                        print(f'Assinado em {happened}')
                        logging.info(f'Assinado em {happened}')
                        print(30 * "=")
                        logging.info(30 * "=")
                    except:
                        driver.close()
                        meu_email.move(lista_email.uid, 'Não assinados')
                        happened = datetime.now().strftime('%d/%m/%Y %H:%M')
                        print(f'Assinatura bloqueada para clicar em {happened}')
                        logging.info(f'Assinatura bloqueada para clicar em {happened}')
                        print(30 * "=")
                        logging.info(30 * "=")
                        break
                elif driver.find_element(by=By.CLASS_NAME, value='btn-warning'):
                    # Caso o documento já foi assiando
                    driver.close()
                    meu_email.move(lista_email.uid, 'Assinados')
                    happened = datetime.now().strftime('%d/%m/%Y %H:%M')
                    print(f'Elemento já assinado {happened}')
                    logging.info(f'Elemento já assinado {happened}')
                    print(30 * "=")
                    logging.info(30 * "=")
                else:
                    # Caso o documento é aberto e não encontra como assinar, ele é lançado para 'Não assinados'
                    driver.close()
                    meu_email.move(lista_email.uid, 'Não assinados')
                    happened = datetime.now().strftime('%d/%m/%Y %H:%M')
                    print(f'Não encontrou onde assinar em {happened}')
                    logging.info(f'Não encontrou onde assinar em {happened}')
                    print(30 * "=")
                    logging.info(30 * "=")
            except:
                driver.close()
                meu_email.move(lista_email.uid, 'Não assinados')
                happened = datetime.now().strftime('%d/%m/%Y %H:%M')
                print(f'Elemento não estava na tela em {happened}')
                logging.info(f'Elemento não estava na tela em {happened}')
                print(30 * "=")
                logging.info(30 * "=")
                break
            sleep(1)
    except:
        # Caso o script pare por excesso de tentativas de verificar o email, ele reloga no email
        contador += 1
        logging.basicConfig(level=logging.INFO, filename=
        f'{caminho}{date.today()}.log', datefmt='%H:%M:%S',
                            filemode='a', format='%(asctime)s,%(msecs)d %(name)s %(levelname)s %(message)s')
        print(f'Tentativa {contador}')
        logging.info(f'Tentativa {contador}')
        sleep(2)
        if contador >= 10:
            login = login()
            senha = senha()
            meu_email = MailBox('outlook.office365.com').login(login, senha)
            happened = datetime.now().strftime('%d/%m/%Y %H:%M')
            print(f"Logando novamente em {happened}")
            logging.info(f"Logando novamente em {happened}")
            contador -= 10
        else:
            continue
