from biblio import importest
import PySimpleGUI as sg
janela = importest.tela_pedido()
while True:
    evento, valores = janela.read()