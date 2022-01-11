#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pyautogui as pg
import pyperclip as pc
import time

pg.PAUSE = 1 #Determinar um delay para nao dar problema na web

# Passo 1: Entrar no sistema da empresa (link do drive)
pg.hotkey("ctrl", "t")
time.sleep(10)
#Se atente com os caracteres especiais nos links, caso eles sejam de endereco br
#O pyperclip sabe lidar com esses caracteres especiais
pc.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pg.hotkey("ctrl", "v")
pg.press("enter")

#carregar o sistema (o delay depende da internet)
time.sleep(10) 

# Passo 2: Navegar no sistema( até a pasta exportar)
pg.click(x=398, y=267, clicks=2)
time.sleep(10)

# Passo 3: Fazer download da base de vendas
pg.click(x=407, y=333)
pg.click(x=1159, y=139)
pg.click(x=1009, y=557)
time.sleep(10)
pg.click(x=499, y=447)

time.sleep(15)


# In[ ]:


# Passo 4: Importar a base de vendas pro python
import pandas as pd
import openpyxl

dataset=pd.read_excel(r"C:\Users\arils\Downloads\Vendas - Dez.xlsx", engine='openpyxl')

#r serve para dizer para o python nao ler caracter especial
display(dataset)

# Passo 5: Calcular o faturamento e quantidade de produtos 
#vendidos( os indicadores)
faturamento = dataset["Valor Final"].sum()
qtd_produtos = dataset["Quantidade"].sum()


# In[ ]:


#Passo 6: Enviar email para diretoria

pg.hotkey("ctrl", "t")
pc.copy("https://mail.google.com/mail/u/4/#inbox")
pg.hotkey("ctrl", "v")
pg.press("enter")

time.sleep(15)

# Clicar no botao escrever
pg.click(x=88, y=170)

# Preencher o destino
pg.write("h.candido20@unifesp.br")
pg.press("tab")
pg.press("tab")
        
# Preencher o assunto
pc.copy("Relatório de Vendas - Excel")
pg.hotkey("ctrl","v")
pg.press("tab")
        
# Escrever o email #USE f PARA FORMATAR O TEXTO COM VARIAVEIS
texto = f"""Teste 

Faturamento:{faturamento:,.2f}
Quantidade produto: {qtd_produtos:,.2f}

Att
"""
pc.copy(texto)
pg.hotkey("ctrl","v")
        
# Clicar em enviar
pg.click(x=843, y=625)
        

