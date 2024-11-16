# -*- coding: utf-8 -*-
"""
Created on Wed Aug 19 16:54:32 2020
"""

import tkinter as tk
from tkinter import ttk
import pandas as pd
import os

class PrincipalRAD:
    def __init__(self, win):
        # Componentes
        self.lblNome = tk.Label(win, text='Nome do Aluno:')
        self.lblNota1 = tk.Label(win, text='Nota 1')
        self.lblNota2 = tk.Label(win, text='Nota 2')
        self.lblMedia = tk.Label(win, text='Média')
        self.txtNome = tk.Entry(bd=3)
        self.txtNota1 = tk.Entry()
        self.txtNota2 = tk.Entry()
        self.btnCalcular = tk.Button(win, text='Calcular Média', command=self.fCalcularMedia)
        
        # Componente TreeView
        self.dadosColunas = ("Aluno", "Nota1", "Nota2", "Média", "Situação")
        self.treeMedias = ttk.Treeview(win, columns=self.dadosColunas, selectmode='browse')
        self.verscrlbar = ttk.Scrollbar(win, orient="vertical", command=self.treeMedias.yview)
        self.verscrlbar.pack(side='right', fill='x')
        self.treeMedias.configure(yscrollcommand=self.verscrlbar.set)
        
        # Cabeçalhos
        for col in self.dadosColunas:
            self.treeMedias.heading(col, text=col)
            self.treeMedias.column(col, minwidth=0, width=100)
        
        self.treeMedias.pack(padx=10, pady=10)

        # Posicionamento dos componentes
        self.lblNome.place(x=100, y=50)
        self.txtNome.place(x=200, y=50)
        self.lblNota1.place(x=100, y=100)
        self.txtNota1.place(x=200, y=100)
        self.lblNota2.place(x=100, y=150)
        self.txtNota2.place(x=200, y=150)
        self.btnCalcular.place(x=100, y=200)
        self.treeMedias.place(x=100, y=300)
        self.verscrlbar.place(x=805, y=300, height=225)

        # Variáveis para controle de registros
        self.id = 0
        self.iid = 0

        # Carrega dados iniciais
        self.carregarDadosIniciais()

    def carregarDadosIniciais(self):
        try:
            fsave = 'C:/temp/PlanilhaAlunos.xlsx'
            dados = pd.read_excel(fsave)
            for i in range(len(dados)):
                nome = dados['Aluno'][i]
                nota1 = str(dados['Nota1'][i])
                nota2 = str(dados['Nota2'][i])
                media = str(dados['Média'][i])
                situacao = dados['Situação'][i]
                self.treeMedias.insert('', 'end', iid=self.iid, values=(nome, nota1, nota2, media, situacao))
                self.iid += 1
                self.id += 1
        except Exception as e:
            print(f'Erro ao carregar dados: {e}')

    def fSalvarDados(self):
        try:
            fsave = 'C:/temp/PlanilhaAlunos.xlsx'
            os.makedirs(os.path.dirname(fsave), exist_ok=True)
            dados = []
            for line in self.treeMedias.get_children():
                lstDados = [value for value in self.treeMedias.item(line)['values']]
                dados.append(lstDados)
            df = pd.DataFrame(data=dados, columns=self.dadosColunas)
            with pd.ExcelWriter(fsave, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Plan1')
                workbook = writer.book
                worksheet = writer.sheets['Plan1']
                for idx, col in enumerate(df.columns):
                    max_len = max((df[col].astype(str).map(len).max(), len(col)))
                    worksheet.set_column(idx, idx, max_len + 2)
            print('Dados salvos com sucesso.')
        except Exception as e:
            print(f'Erro ao salvar os dados: {e}')

    def fVerificarSituacao(self, nota1, nota2):
        media = (nota1 + nota2) / 2
        if media >= 7.0:
            situacao = 'Aprovado'
        elif media >= 5.0:
            situacao = 'Em Recuperação'
        else:
            situacao = 'Reprovado'
        return media, situacao

    def fCalcularMedia(self):
        try:
            nome = self.txtNome.get()
            nota1 = float(self.txtNota1.get())
            nota2 = float(self.txtNota2.get())
            media, situacao = self.fVerificarSituacao(nota1, nota2)

            # Verifica se o aluno já existe no Treeview
            aluno_encontrado = False
            for line in self.treeMedias.get_children():
                if self.treeMedias.item(line)['values'][0] == nome:
                    # Atualiza os dados do aluno existente
                    self.treeMedias.item(line, values=(nome, f"{nota1:.1f}", f"{nota2:.1f}", f"{media:.1f}", situacao))
                    aluno_encontrado = True
                    break

            # Se o aluno não existir, cria um novo registro
            if not aluno_encontrado:
                self.treeMedias.insert('', 'end',
                                       iid=self.iid,
                                       values=(nome, f"{nota1:.1f}", f"{nota2:.1f}", f"{media:.1f}", situacao))
                self.iid += 1
                self.id += 1

            # Salva os dados no arquivo Excel
            self.fSalvarDados()
        except ValueError:
            print('Por favor, insira valores válidos para nome e notas.')
        finally:
            self.txtNome.delete(0, 'end')
            self.txtNota1.delete(0, 'end')
            self.txtNota2.delete(0, 'end')

# Programa Principal
janela = tk.Tk()
principal = PrincipalRAD(janela)
janela.title('Bem Vindo ao RAD')
janela.geometry("820x600+10+10")
janela.mainloop()

