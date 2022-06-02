# -*- coding: utf-8 -*-
from email import encoders
from email.mime.base import MIMEBase
from optparse import Values
import os
from os import link
import pandas as pd
from openpyxl import workbook, load_workbook
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from inspect import classify_class_attrs

login = open('Login.txt','r')
email = login.readline()
senha = login.readline()
email = str(email)
print(email)

f = open('config.txt','r')
arquivo=f.read()
x = pd.read_excel(arquivo+'.xlsx')

Texto_email = open('Texto_email.txt')
Texto_email = Texto_email.read()

Posto_email = dict([
('CENTRO DE SAUDE I DR LOURENCO QUILICCI', 'centrodesaude.saude@gmail.com'), 
('EACS PLANEJADA II','planejada2.saude@gmail.com'),
('ESF PLANEJADA I',"planejada1.saude@gmail.com"),
('ESF AGUA COMPRIDA',"aguacomprida.saude@gmail.com"),
('ESF AGUA CLARA I',"aguasclaras.saude@gmail.com"),
('ESF CASA DE JESUS',"casadejesus.saude@gmail.com"),
('ESF CDHU SAADA NADER ABI CHEDID',"cdhu.saude@gmail.com"),
('ESF CIDADE JARDIM','cidadejardim.saude@gmail.com'),
('ESF HENEDINA RODRIGUES CORTEZ','henedinacortez.saude@gmail.com'),
('ESF HIPICA JAGUARI','hipicajaguari.saude2015@gmail.com'),
('ESF NILDA COLLI','nildacolli.saude@gmail.com'),
('ESF MADRE PAULINA JD FRATERNIDADE','madrepaulina.saude@gmail.com'),
('ESF PARQUE DOS ESTADOS II','parque2.saude@gmail.com'),
('ESF PARQUE DOS ESTADOS I','parque1.saude@gmail.com'),
('ESF PEDRO MEGALE','pedromegale.saude@gmail.com'),
('ESF PLANEJADA I',"planejada1.saude@gmail.com"),
('ESF SAO FRANCISCO DE ASSIS','esfescolausf@gmail.com'),
('ESF SAO LOURENCO','saolourenco.saude@gmail.com'),
('ESF SAO MIGUEL','saomiguelbp.saude@gmail.com'),
('ESF SAO VICENTE','saovicentebp.saude@gmail.com'),
('ESF TORO II','toro.saude@gmail.com'),
('ESF VILA BIANCHI','vilabianchi.saude@gmail.com'),
('ESF VILA DAVID I','viladavi1.saude@gmail.com'),
('ESF VILA DAVID II','viladavi2.saude@gmail.com'),
('ESF VILA MOTTA','vilamotta.saude@gmail.com'),
('ESPACO DO ADOLESCENTE','ubsespacoadolescente@gmail.com'),
('SAE SERVICO DE ATENCAO ESPECIALIZADA','sae.equipebp@gmail.com'),
('UBS ARARA DOS MORI','araradosmori.saude@gmail.com'),
('UBS BIRICA DO VALADO','biricadovalado.saude@gmail.com'),
('UBS MAE DOS HOMENS','maedoshomens.saude@gmail.com'),
('UBS MORRO GRANDE DA BOA VISTA','morrogrande.saude@gmail.com'),
('UBS SANTA LUZIA','santaluziabp.saude@gmail.com'),
('UBS VILA APARECIDA','vilaaparecida.saude@gmail.com')
])

def enviar_email(Unidade,UnidadeL):  

    print("\n\nEmail :",Unidade)
    email_para = "edu.py.codigolivre@gmail.com" #Unidade
    email_de = email
    subject = "FILA DE ESPERA SISREG" #titulo
    
    body =  Texto_email 

    message = MIMEMultipart()
    message["From"] = email_de
    message['To'] = email_para
    message["Subject"] = subject

    Cam_arquivo = UnidadeL + '.xlsx'
    attchment = open(Cam_arquivo,'rb')
    att = MIMEBase('aplication','octet-stream')
    att.set_payload(attchment.read())
    encoders.encode_base64(att)

    att.add_header('Content-Disposition', f'attchment; filename = {Cam_arquivo}')
    attchment.close()
    message.attach(att)

    message.attach(MIMEText(body, 'plain'))
    text = message.as_string()

    print("login efetuado")
    mail = smtplib.SMTP('smtp.gmail.com', 587)
    mail.ehlo()
    mail.starttls()
    mail.ehlo()
    mail.login(email_de,senha)
    mail.sendmail(message['From'],email_para, text)
    mail.close()

def Autopy():

    Lista = list(Posto_email.items())
    for item in Lista:

        Unidade = item[0]
        
        filtro = x.loc[x['NOME UNIDADE SOLICITANTE'].str.contains(Unidade)]
        pasta_destino = Unidade + '.xlsx'

        print(pasta_destino)
        filtro.to_excel(pasta_destino)

        unidade_email= item[1]
        
        enviar_email(unidade_email,Unidade)

        os.remove(pasta_destino)

Autopy()