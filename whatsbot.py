# Importar bibliotecas
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

PATH = "chromedriver.exe"
driver = webdriver.Chrome(PATH)


def abrir_whatsapp():  # Navegar até o whatsapp web
    driver.get("https://web.whatsapp.com/")
    driver.maximize_window()
    time.sleep(10)


def buscar_contato(contato):  # Definir contato
    campo_pesquisa = driver.find_element_by_xpath(
        '//div[contains(@class, "copyable-text selectable-text")]')
    time.sleep(3)
    campo_pesquisa.click()
    campo_pesquisa.send_keys(contato)
    campo_pesquisa.send_keys(Keys.ENTER)


def enviar_mensagem(mensagem):  # Enviar mensagem
    campo_mensagem = driver.find_elements_by_xpath(
        '//div[contains(@class, "copyable-text selectable-text")]')
    campo_mensagem[1].click()
    time.sleep(3)
    campo_mensagem[1].send_keys(mensagem)
    campo_mensagem[1].send_keys(Keys.ENTER)
