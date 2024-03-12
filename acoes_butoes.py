import customtkinter as ctk
from tkinter import *
from tkinter import messagebox


class configurar_butoes:
    def __init__(self) -> None:
        self.name_value = StringVar()
        self.name_contato = StringVar()
        self.nameidade = StringVar()
        self.endereco = StringVar()
        self.obs = StringVar()
        
    def acao_salvar(self):
        pass
    
    def atualizar_imput(self):
        self.name = self.name_value.get()
        contato = self.contato.get()
        idade = self.idade.get()
        endereco = self.endereco.get()
        obs = self.obs.get(0.0,END)
        
    def acao_limpar(self):
        pass