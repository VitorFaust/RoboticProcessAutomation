# Boa noite, Pessoal!
# Me chamo VitorSousaFaust, Fiz essas funções para agilizar o processo de desenvolvimento de RPA no meu trabalho.
# Espero que gostem!
# Dicas de melhoria no codigo ou duvidas,serão muito bem vinda! basta enviar para algum dos meus contatos abaixo.
# E-mail: vitorfaust96@gmail.com
# Linkedin: https://www.linkedin.com/in/vitor-faust-7ba838209/
# GitHub: https://github.com/VitorFaust

import os
from datetime import datetime
from time import *
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
import pyautogui


#Aguarda o aparecimento de uma imagem na tela, para que continue o processo.
def waitImage(path, cf):
    #waitImage(string,int)
    while True:
        imagemlocation = pyautogui.locateCenterOnScreen(path, confidence=cf)
        if imagemlocation != None:
            print("imagem encontrada!")
            break
        else:
            continue


#clicka em uma imagem quando a mesma aparece na tela.
#O uso ideal dessa função, e que seja utilizada posteriormente á função waitImage().
def clickImage(path, cf):
    #clickImage(path->string, cf->int)
    while True:
        parametro = pyautogui.locateCenterOnScreen(path, confidence=cf)
        if parametro != None:
            sleep(0.1)
            pyautogui.moveTo(parametro)
            pyautogui.click(parametro)
            break
        else:
            continue


#Espera uma imagem aparecer na tela, por tempo determinado.
def waitTime(path, cf, tempo):
    #waitTime(path->string, cf->int, tempo->int)
    contador = 0
    while True:
        parametro = pyautogui.locateCenterOnScreen(path, confidence=cf)
        if parametro != None:
            sleep(0.1)
            break
        else:
            if contador == tempo:
                break
            sleep(0.1)
            contador = contador + 0.1


#Essa função serve para clicar quantas vezes desejar, em uma imagem especifica.
#ideal para paginas que oferecem apenas setas para nagevar entre elas.
def clickTimes(path, cf, vezes):
    #clickTimes(path->string, cf->int, vezes->int):
    x = -1
    while True:
        x = x + 1
        while True:
            parametro = pyautogui.locateCenterOnScreen(path, confidence=cf)
            if parametro != None:
                pyautogui.moveTo(parametro)
                pyautogui.click(parametro)
                break
            else:
                continue
        if x == vezes:
            break
        else:
            continue


#Essa função server para clicar em uma imagem dentro de outra imagem.
#O ideal é utilizar essa função em casos que tenham imagens parecidas externamente da imagem desejada.
def clickInsideImage(imagem, dentro_da_imagem, cf):
    #clickInsideImage(imagem->string, dentro_da_imagem->string, cf->int)
    while True:

        parametro = pyautogui.locateOnScreen(+imagem, confidence=cf)
        parametro2 = pyautogui.locateOnScreen(+dentro_da_imagem, region=parametro)

        print('parametro1', parametro, 'parametro2', parametro2)
        if parametro != None:
            pyautogui.moveTo(pyautogui.center(parametro2))
            pyautogui.click(pyautogui.center(parametro2))
            break
        else:
            continue


#Essa função serve para automatizar o envio de E-mails.
#a mensagem tem que serem formatadas em html.
def send_email_rpa(remetente, senhaRemetente, destinatario, assunto, mensagem):
    #send_email_rpa(remetente->string, senhaRemetente->string, destinatario->string, assunto->string, mensagem->string)
    # está configurado para o outlook. caso queira enviar através de outras plataformas, tem que mudar a config smtp.

    msg = MIMEMultipart()

    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Subject'] = assunto
    msg.attach(MIMEText(mensagem, 'html'))

    server = smtplib.SMTP('smtp.office365.com', 587)
    server.starttls()
    server.login(remetente, senhaRemetente)
    server.sendmail(msg['From'], msg['to'], msg.as_string())
    server.quit()
