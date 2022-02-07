import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from tkinter.filedialog import askopenfilename
import pandas as pd
import requests
from datetime import datetime
import numpy as np

requisicao = requests.get('https://economia.awesomeapi.com.br/json/all')
dicionario_moedas = requisicao.json()

lista_moedas = list(dicionario_moedas.keys())


def pegar_cotacao():
    moeda = combobox_selecionemoeda.get()
    data_cotacao = Calendario_moeda.get()
    ano = data_cotacao[-4:]
    mes = data_cotacao[3:5]
    dia = data_cotacao[:2]
    link = f'https://economia.awesomeapi.com.br/{moeda}-BRL/10?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}'
    requicao_moeda = requests.get(link)
    cotacao = requicao_moeda.json()
    valor_moeda = cotacao[0]['bid']
    label_textocotacao['text'] = f'A cotação da {moeda} no dia {data_cotacao} foi de: R$ {valor_moeda}'

def selecionar_arquivo():

    caminho_arquivo = askopenfilename(title='Selecione o arquivo de Moeda ')
    var_caminhoarquivo.set(caminho_arquivo)
    if caminho_arquivo:
        label_arquivoselecionado['text'] = f'Arquivo selecionado: {caminho_arquivo}'


def atualizar_cotacoes():

    try:

        df = pd.read_excel(var_caminhoarquivo.get())
        moedas = df.iloc[:, 0]
        data_inicial = Calendario_datainicial.get()
        data_final = Calendario_datafinal.get()

        ano_inicial = data_inicial[-4:]
        mes_inicial = data_inicial[3:5]
        dia_inicial = data_inicial[:2]

        ano_final = data_final[-4:]
        mes_final = data_final[3:5]
        dia_final = data_final[:2]


        for moeda in moedas:

            link = f'https://economia.awesomeapi.com.br/{moeda}-BRL/10?' \
                   f'start_date={ano_inicial}{mes_inicial}{dia_inicial}&' \
                   f'end_date={ano_final}{mes_final}{dia_final}'
            requicao_moeda = requests.get(link)
            cotacoes = requicao_moeda.json()

            for cotacao in cotacoes:
                timestamp = int(cotacao['timestamp'])
                bid = float(cotacao['bid'])
                data = datetime.fromtimestamp(timestamp)
                data = data.strftime('%d/%m/%Y')
                if data not in df:

                    df[data] = np.nan

                df.loc[df.iloc[:, 0] == moeda, data] = bid

        df.to_excel('Teste.xlsx')
        label_atualizarcotacoes['text'] = 'Arquivo atualizado com sucesso'


    except:

        label_atualizarcotacoes['text'] = 'Selecione um arquivo excel no formato correto'
        
        
janela = tk.Tk()

janela.title('Ferramenta de cotação de Moedas')
label_cotacaomoeda = tk.Label(text= "Cotação de moeda especifica", borderwidth=2, relief='solid')
label_cotacaomoeda.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky='nswe',)

label_selecionarmoeda = tk.Label(text= "Selecionar moeda: ", anchor= 'e' )
label_selecionarmoeda.grid(row=1, column=0, padx=10, pady=10, sticky='nswe', columnspan=2, )

combobox_selecionemoeda = ttk.Combobox(values=lista_moedas)
combobox_selecionemoeda.grid(row=1, column=2, padx=10, pady=10, sticky='nswe')

label_selecionardia = tk.Label(text= "Selecionar o dia que deseja pegar a cotação: ", anchor= 'e' )
label_selecionardia.grid(row=2, column=0, padx=10, pady=10, sticky='nswe', columnspan=2, )

Calendario_moeda = DateEntry(year=2022, locale='pt_br')
Calendario_moeda.grid(row=2, column=2, padx=10, pady=10, sticky='nswe')

label_textocotacao = tk.Label(text="")
label_textocotacao.grid(row=3, column= 0, padx=10, pady=10, sticky='nswe', columnspan=2,)

botao_pegarcotacao = tk.Button(text='Pegar cotação', command= pegar_cotacao)
botao_pegarcotacao.grid(row=3, column=2, padx=10, pady=10, sticky='nswe' )

#cotação de varias moedas

label_multiplascotacoes = tk.Label(text= "Multiplas cotações de Moedas", borderwidth= 2, relief='solid')
label_multiplascotacoes.grid(row=4, column=0, padx=10, pady=10, sticky='nswe', columnspan=3,)

label_selecionararquivo = tk.Label(text='Selecione um arquivo em excel com as Moedas na Coluna A')
label_selecionararquivo.grid(row=5, column=0 ,padx=10, pady=10, columnspan=2, sticky='nswe', )

var_caminhoarquivo = tk.StringVar()

botao_selecionararquivo = tk.Button(text='Clique aqui para selecionar: ', command=selecionar_arquivo )
botao_selecionararquivo.grid(row=5, column=2, padx=10, pady=10, sticky='nswe',)

label_arquivoselecionado = tk.Label(text='Nenhum arquivo selecionado', anchor='e')
label_arquivoselecionado.grid(row=6, column=0, padx=10, pady=10, columnspan=3, sticky='nswe', )

label_datainicial = tk.Label(text='Data inicial:', anchor='e')
label_datafinal = tk.Label(text='Data final:', anchor= 'e')

label_datainicial.grid(row=7, column=0, padx=10, pady=10, sticky='nswe',)
label_datafinal.grid(row=8, column=0, padx=10, pady=10, sticky='nswe',)

Calendario_datainicial = DateEntry(year=2022, locale='pt_br')
Calendario_datafinal = DateEntry(year=2022, locale='pt_br')

Calendario_datainicial.grid(row=7, column=1, padx=10, pady=10, sticky='nswe')
Calendario_datafinal.grid(row=8, column=1, padx=10, pady=10, sticky='nswe')

botao_atualizarcotacoes= tk.Button(text= 'Atualizar cotações', command=atualizar_cotacoes)
botao_atualizarcotacoes.grid(row=9, column=0, padx=10, pady=10, sticky='nswe')

label_atualizarcotacoes = tk.Label(text='')
label_atualizarcotacoes.grid(row=9, column=1, columnspan=2, padx=10, pady=10, sticky='nswe')

botao_fechar = tk.Button(text= 'Fechar', command= janela.quit)
botao_fechar.grid(row=10, column=2, padx=10, pady=10, sticky='nswe')

janela.mainloop()

