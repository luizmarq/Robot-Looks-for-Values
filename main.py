from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
from tkinter.filedialog import askopenfilename
import time
import pandas as pd
import unidecode

ctk.set_appearance_mode('light')
ctk.set_default_color_theme('blue')

cor_preto = '#282829'
cor_branco = '#FAFAFA'
cor_verde = '#0A7641'
cor_verde_escuro = '#086034'

def caminho_arquivo():
    caminho = askopenfilename(title="Selecione o arquivo Excel")
    return caminho

def posicionamento_janela(window, width, height):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    x = (screen_width/2) - (width/2)
    y = (screen_height/2) - (height/2)
    window.geometry('%dx%d+%d+%d' % (width, height, x, y))

def buscar_navegador(linha, produto, planilha):
    # Configura automaticamente o ChromeDriver na versão correta
    driver = webdriver.Chrome(ChromeDriverManager().install())

    # MAGAZINE LUIZA
    produtoMagazine = unidecode.unidecode(produto).replace(' ', '+')
    driver.get("https://www.magazineluiza.com.br/busca/" + produtoMagazine + "/?from=submit")
    time.sleep(5)

    # Get Value Magazine Luiza
    try:
        valorMagazine = driver.find_elements(By.CSS_SELECTOR, 'p[data-testid="price-value"].sc-kpDqfm.eCPtRw.sc-camqpD.cFgZBi')
        print(valorMagazine[0].text)
        planilha.at[linha, 'Magazine Luiza'] = valorMagazine[0].text
        
    except:
        planilha.at[linha, 'Magazine Luiza'] = 'N/I'


    # PONTO FRIO
    produtoPonto = produtoMagazine.replace('+', '-')
    driver.get("https://www.pontofrio.com.br/" + produtoPonto + '/b')
    time.sleep(5)

    # Get Value Ponto Frio
    try:
        valorPonto = driver.find_elements(By.CSS_SELECTOR, 'div.product-card__highlight-price[aria-hidden="true"]') 
        print(valorPonto[0].text)
        planilha.at[linha, 'Ponto Frio'] = valorPonto[0].text
    
    except:
        planilha.at[linha, 'Ponto Frio'] = 'N/I'


    # CASAS BAHIA
    driver.get("https://www.casasbahia.com.br/" + produtoPonto + '/b')
    time.sleep(5)

    # Get Value Ponto Frio
    try:
        valorCasas = driver.find_elements(By.CSS_SELECTOR, 'div.product-card__highlight-price[aria-hidden="true"]')
        print(valorCasas[0].text)
        planilha.at[linha, 'Casas Bahia'] = valorCasas[0].text

    except:
        planilha.at[linha, 'Casas Bahia'] = 'N/I'

    # Fecha o navegador
    driver.quit()

class MainWindow(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.criando_widgets()

    def criando_widgets(self):
        w = 500
        h = 200
        posicionamento_janela(self, w, h) #abrindo no meio da tela do PC
        self.title("")
        self.minsize(w, h)
        self.resizable(False, False)
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        #self.grid_rowconfigure(0, weight=1)

        self.lbl_arquivo = ctk.CTkLabel(self, text="", text_color=cor_preto, font=('Times New Roman', 12))
        self.lbl_arquivo.grid(columnspan=2, column=0, row=0, sticky=ctk.W+ctk.E)
        self.lbl_arquivo.grid_configure(pady=(40), padx=(30))

        bt_selecionar = ctk.CTkButton(self, text='Selecione a planilha', 
                                      command=lambda: self.localizar_arquivo(self.lbl_arquivo), 
                                      font=('Times New Roman', 16, 'bold'),
                                      fg_color=cor_verde,
                                      hover_color=cor_verde_escuro)
        bt_selecionar.grid(column=0, row=1, sticky=ctk.W+ctk.E)
        bt_selecionar.grid_configure(padx=(30, 5))

        bt_espelho = ctk.CTkButton(self, text='Pesquisar', command=lambda: self.rodar_bot(self.lbl_arquivo.cget('text')), font=('Times New Roman', 16, 'bold'))
        bt_espelho.grid(column=1, row=1, sticky=ctk.W+ctk.E)
        bt_espelho.grid_configure(padx=(5, 30))

        lbl_feedback = ctk.CTkLabel(self, text="", text_color=cor_preto, font=('Times New Roman', 12))
        lbl_feedback.grid(columnspan=2, column=0, row=2, sticky=ctk.W+ctk.E)
        lbl_feedback.grid_configure(pady=(40), padx=(30))

    def localizar_arquivo(self, objeto):
        texto = caminho_arquivo()
        objeto.configure(text=texto)

    def rodar_bot(self, arquivo):
        if not arquivo:
            messagebox.showerror('Erro #1', 'Selecione uma planilha válida')
        else:
            #buscar_navegador()

            planilha = pd.read_excel(arquivo)
            total_prod = len(planilha['Produtos'].dropna())
            
            # Se não houver nenhum produto na planilha
            if total_prod == 0:
                messagebox.showerror('Erro #2', 'Não há produtos relacionados')
            else:

                # Rodar para cada produto
                for index, produto in enumerate(planilha["Produtos"]):
                    buscar_navegador(index, produto, planilha)

                    # Salvando a planilha
                    planilha.to_excel(arquivo, index=False)

if __name__ == '__main__':
    app = MainWindow()
    app.mainloop()


