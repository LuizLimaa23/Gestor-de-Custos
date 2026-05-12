import openpyxl
import os
import sys
from datetime import datetime
mapadiasy ={
"segunda" : 6,
"terça" : 7,
"quarta" : 8,
"quinta" : 9,
"sexta" : 10,
"sabádo" : 11,
"domingo" : 12 }
mapasemanax = {
    "1" : 3,
    "2" : 12 }







def add_entrada(aba,semana,dia,valor):
     linha = mapadiasy[dia]
     coluna =mapasemanax[semana]
     for i in range(5):
         colunaatual= coluna + i
         celula = aba.cell(row=linha, column=colunaatual)
         if celula.value is None or celula.value == 0:
                celula.value = valor
                print("Valor adicionado")
                return
     print("Todos os valores diarios preenchidos!!")    
         
         

        
     


def registrar_gastos():
    arquivoexcel = "controledegastos.xlsx"

    if not os.path.exists(arquivoexcel):
        print("ERRO: O arquivo não pode ser encontrado, por favor insira ele na pasta do executável!")
        input("\n Pressione enter pra continuar")
        return
    try:
        wb=openpyxl.load_workbook(arquivoexcel)
        aba= wb['entradas']
        print("                           ||Sistema de controle de gastos||           \n")
        inicio= input("       Pressione E para continuar ou S para sair: ").lower()
        if inicio == 's': return
        elif inicio == "e":
                print("Digite a opção desejada: \n"
                "1 para adicionar valores \n" \
                "2 para visualizar entradas\n" \
                "3 para visualizar gastos")
                opcao = input(": ")
                if opcao == "1":
                     while True:
                        semana= input("Escolha a semana!(1,2,3,4): ")
                        dia = input("Escolha o dia da semana: ").lower()
                        while True:
                             valor1 = input("Digite o valor a ser adicionado: ").replace(',','.')
                             try: 
                                  valor2=float(valor1)
                                  break
                             except: print("Digite apenas números!")
                            
                        add_entrada(aba, semana,dia,valor2)
                        wb.save(arquivoexcel)
                        loop= input("Valor adicionado, deseja [c]ontinuar ou [s]air? ").lower()
                        if loop == "c":
                             continue
                        elif loop == "s":
                             break
                        else: print ("Digite uma resposta válida! ")
        else: print("Digite uma opção válida!")
    except: print("Ocorreu um erro! ")


if __name__ == "__main__":
    registrar_gastos()
                            


    
        
