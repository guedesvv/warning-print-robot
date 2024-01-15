from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.edge.options import Options
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.edge.service import Service
from time import sleep
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.common.exceptions import SessionNotCreatedException
from PIL import Image
from datetime import datetime
import pyautogui
import time

write = pyautogui.write
data_hora_atual = datetime.now()
formato_nome_arquivo = data_hora_atual.strftime("%Y-%m-%d_%H-%M-%S")
url_user = 'joao.vguedes'
url_pass = "SN1747"
t = 60

tentativas_maximas = 5
tentativa_atual = 0

while tentativa_atual < tentativas_maximas:
    try:
        print(f"Tentativa {tentativa_atual + 1} de iniciar o driver.")
        # Caminho para o seu perfil do Edge
        edge_profile_path = r'C:\Users\GPS\AppData\Local\Microsoft\Edge\User Data\Trabalho'

        # Configurar as opções do Edge com o caminho do perfil
        edge_options = Options()
        edge_options.add_argument(f'--user-data-dir={edge_profile_path}')

        # Inicializar o driver do Edge com as opções configuradas
        driver = webdriver.Edge(options=edge_options)
        t = 10

        print("Driver iniciado, tentando abrir o URL.")
        driver.get(
            'https://portal.gpssa.com.br/GPS/Login.aspx?ReturnUrl=%2fGPS%2fPortal.aspx')
        driver.maximize_window()

        print("URL aberto com sucesso.")
        break  # Se a sessão for criada com sucesso, saia do loop

    except SessionNotCreatedException as e:
        print(f"Falha na criação da sessão: {e}")
        tentativa_atual += 1
        sleep(1)
        if tentativa_atual == tentativas_maximas:
            print(
                "Número máximo de tentativas atingido. Não foi possível criar a sessão.")

    except Exception as e:
        print(f"Ocorreu um erro inesperado: {e}")
        break

# Certifique-se de que o driver seja fechado em caso de falha
if 'driver' in locals() and tentativa_atual == tentativas_maximas:
    driver.quit()


# Insira credenciais
target = '//*[@id="txtUsername-inputEl"]'  # Usuario
WebDriverWait(driver, t).until(EC.presence_of_element_located(
    (By.XPATH, target))).send_keys(list(url_user))

target = '//*[@id="txtPassword-inputEl"]'  # Senha
elem = WebDriverWait(driver, t).until(
    EC.presence_of_element_located((By.XPATH, target)))
elem.send_keys(list(url_pass))

target = '//*[@id="btnLogin"]'  # Btn Login
WebDriverWait(driver, t).until(
    EC.presence_of_element_located((By.XPATH, target))).click()

sleep(5)

target = "//span[text()='Relatório de Acompanhamento']"
elem = WebDriverWait(driver, t).until(
    EC.presence_of_element_located((By.XPATH, target)))
ActionChains(driver).double_click(on_element=elem).perform()
sleep(2)

# GPS VISTA
# pasta de relatorios e acompanhamentoss
target = "//span[text()='16 - GPS Vista']"
elem = WebDriverWait(driver, t).until(
    EC.presence_of_element_located((By.XPATH, target)))
ActionChains(driver).double_click(on_element=elem).perform()

sleep(2)
# Iris
target = "//span[text()='16.1 - ÍRIS']"
elem = WebDriverWait(driver, t).until(
    EC.presence_of_element_located((By.XPATH, target)))
ActionChains(driver).double_click(on_element=elem).perform()

# Tarefas escalonadas
target = "//span[text()='16.1.2 - Tarefas Escalonadas']"
elem = WebDriverWait(driver, t).until(
    EC.element_to_be_clickable((By.XPATH, target))).click()

sleep(5)
# Mude para o iframe pai
iframe_pai = driver.find_element(By.XPATH, '//*[@id="box-1053"]')
driver.switch_to.frame(iframe_pai)

# Mude para o iframe filho
# Aguarde pelo iframe filho
# Ajuste o tempo de espera conforme necessário
wait = WebDriverWait(driver, 10)
xpath_iframe_filho = "//iframe[contains(@src, 'powerbi.com')]"
iframe_filho = wait.until(
    EC.presence_of_element_located((By.XPATH, xpath_iframe_filho)))

# Mude para o iframe filho
driver.switch_to.frame(iframe_filho)

# Editar filtros do BI

sleep(25)

target = '(//*[@class="visual customPadding allow-deferred-rendering visual-image"])[1]'
elem = (
    WebDriverWait(driver, t)
    .until(EC.element_to_be_clickable((By.XPATH, target)))
    .click()
)

# Filtro diretor
target = "//div[@class='slicer-dropdown-menu' and @aria-label='DIRETOR']"
elem = (
    WebDriverWait(driver, t)
    .until(EC.element_to_be_clickable((By.XPATH, target)))
    .click()
)


sleep(4)
# Colocando Filtros
pyautogui.typewrite("JAZEEL")
sleep(4)
pyautogui.press("down")
time.sleep(1)
pyautogui.press("enter")
sleep(10)

# Fechando guia dos filtros

target = "/html/body/div[2]/report-embed/div/div[1]/div/div/div/div/exploration-container/div/div/docking-container/div/div/div/div/exploration-host/div/div/exploration/div/explore-canvas/div/div[2]/div/div[2]/div[2]/visual-container-repeat/visual-container-group[1]/transform/div/div[2]/visual-container[3]/transform/div/div[3]/div/div/visual-modern/div/div"
elem = (
    WebDriverWait(driver, t)
    .until(EC.element_to_be_clickable((By.XPATH, target)))
    .click()
)

sleep(5)

# Print tarefas escalonadas
target = '//*[@id="pbiAppPlaceHolder"]'
elem = WebDriverWait(driver, t).until(
    EC.presence_of_element_located((By.XPATH, target))
)
screenshot_filename = "screenshot_tarefas_escalonadas.png"
elem.screenshot(screenshot_filename)

WebDriverWait(driver, t).until(
    EC.visibility_of_element_located((By.CSS_SELECTOR, "svg.mainGraphicsContext"))
)


# Agora que o <rect> está visível, localiza o elemento desejado para o screenshot
target = '//*[@id="pbiAppPlaceHolder"]'
elem = WebDriverWait(driver, t).until(
    EC.presence_of_element_located((By.XPATH, target))
)

# Tira o screenshot do elemento específico
screenshot_filename = "screenshot_tarefas_escalonadas.png"
elem.screenshot(screenshot_filename)

# Para voltar ao iframe pai
driver.switch_to.parent_frame()
driver.switch_to.default_content()
sleep(5)

# Visitas operacionais
target = "//span[text()='16.1.3 - Visita Liderança']"
elem = WebDriverWait(driver, t).until(
    EC.presence_of_element_located((By.XPATH, target))
)
ActionChains(driver).double_click(on_element=elem).perform()
sleep(2)

# visit operacional lideranca
target = "//span[text()='16.1.3.1 - Visita Oper. Liderança']"
elem = (
    WebDriverWait(driver, t)
    .until(EC.presence_of_element_located((By.XPATH, target)))
    .click()
)
sleep(2)

# Mude para o iframe pai
iframe_pai = driver.find_element(By.XPATH, '//*[@id="box-1056"]')
driver.switch_to.frame(iframe_pai)

# Mude para o iframe filho
# Aguarde pelo iframe filho
# Ajuste o tempo de espera conforme necessário
wait = WebDriverWait(driver, 10)
xpath_iframe_filho = "//iframe[contains(@src, 'powerbi.com')]"
iframe_filho = wait.until(
    EC.presence_of_element_located((By.XPATH, xpath_iframe_filho)))

# Mude para o iframe filho
driver.switch_to.frame(iframe_filho)


# Editar filtros do BI

sleep(25)

target = "/html/body/div[2]/report-embed/div/div[1]/div/div/div/div/exploration-container/div/div/docking-container/div/div/div/div/exploration-host/div/div/exploration/div/explore-canvas/div/div[2]/div/div[2]/div[2]/visual-container-repeat/visual-container-group/transform/div/div[2]/visual-container[2]/transform/div/div[3]/div/div"
elem = (
    WebDriverWait(driver, t)
    .until(EC.element_to_be_clickable((By.XPATH, target)))
    .click()
)

# Filtro diretor
target = "//div[@class='slicer-dropdown-menu' and @aria-label='DIRETOR']"
elem = (
    WebDriverWait(driver, t)
    .until(EC.element_to_be_clickable((By.XPATH, target)))
    .click()
)


sleep(4)
# Colocando Filtros
pyautogui.typewrite("JAZEEL")
sleep(4)
pyautogui.press("down")
time.sleep(1)
pyautogui.press("enter")
sleep(10)

# Fechando guia dos filtros

target = "/html/body/div[2]/report-embed/div/div[1]/div/div/div/div/exploration-container/div/div/docking-container/div/div/div/div/exploration-host/div/div/exploration/div/explore-canvas/div/div[2]/div/div[2]/div[2]/visual-container-repeat/visual-container-group[1]/transform/div/div[2]/visual-container[3]/transform/div/div[3]/div/div/visual-modern/div/div"
elem = (
    WebDriverWait(driver, t)
    .until(EC.element_to_be_clickable((By.XPATH, target)))
    .click()
)

# Print tarefas escalonadas
target = '//*[@id="pbiAppPlaceHolder"]'
elem = WebDriverWait(driver, t).until(
    EC.presence_of_element_located((By.XPATH, target))
)
screenshot_filename = "screenshot_visita_operacao_liderança.png"
elem.screenshot(screenshot_filename)

WebDriverWait(driver, t).until(
    EC.visibility_of_element_located((By.CSS_SELECTOR, "svg.mainGraphicsContext"))
)


# Agora que o <rect> está visível, localiza o elemento desejado para o screenshot
target = '//*[@id="pbiAppPlaceHolder"]'
elem = WebDriverWait(driver, t).until(
    EC.presence_of_element_located((By.XPATH, target))
)

# Tira o screenshot do elemento específico
screenshot_filename = "screenshot_visita_operacao_liderança.png"
elem.screenshot(screenshot_filename)

# Para voltar ao iframe pai
driver.switch_to.parent_frame()
driver.switch_to.default_content()
sleep(5)

# Whatsapp Web
driver.get("https://web.whatsapp.com/")
driver.maximize_window()

elem.send_keys(Keys.ENTER)
sleep(60)
