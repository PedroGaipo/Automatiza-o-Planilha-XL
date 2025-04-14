from openpyxl import load_workbook #importa a biblioteca opepyxl com o comando load_workbook que serve para ler a planilha
from copy import copy
planilha = load_workbook("alunos_escola_ficticia_105.xlsx")
abadados = planilha["Sheet1"] #todos os dados que vamos pegar para alterar a planilha serão da aba Sheet1
def criaraba(Sala, planilha, formatacao): #Função que serve para criar uma aba para organizar os alunos de uma escola por suas respectivas salas e formatar as células
    if Sala not in planilha.sheetnames: #se não existir uma planilha chamada sala então será criado uma nova aba chamada sala com as celulas: nome, idade, sala, bairro, sexo, notafinal
        planilha.create_sheet(Sala)
        novaaba = planilha[Sala]
        novaaba ["A1"].value = "Nome"
        novaaba ["B1"].value = "Idade"
        novaaba ["C1"].value = "Sala"
        novaaba ["D1"].value = "Bairro"
        novaaba ["E1"].value = "Sexo"
        novaaba ["F1"].value = "NotaFinal"
        novaaba ["A1"]._style = formatacao
        novaaba ["B1"]._style = formatacao
        novaaba ["C1"]._style = formatacao
        novaaba ["D1"]._style = formatacao
        novaaba ["E1"]._style = formatacao
        novaaba ["F1"]._style = formatacao

def transinfo(abadeorigem, abadestino, linhaorigem):
    #Nessa parte o código precisa saber de qual linha da planilha original precisa ser pegada e jogada para a nova aba da sala do aluno.
    linhadestino = abadestino.max_row + 1 #na linha de destino será adicionada a informação que foi pegada e sempre adicionar na linha abaixo a próxima informação
    for coluna in range(1, 7):#Aqui o código percorrerá as colunas: Nome, Idade, Sala, Bairro, Sexo e NotaFinal
        abadestino.cell(row=linhadestino, column=coluna).value = abadeorigem.cell(row=linhaorigem, column=coluna).value#Aqui o código colocará na aba de destino(Sala) toda a informação que foi percorrida nas abas de origem e que tem a ver com a sala da nova aba
        
todasaslinhas = abadados.max_row #comando para ver todas as linhas da planilha
print("Atualmente existem ", todasaslinhas, " linhas na planilha")

formatacao = copy(abadados["A1"]._style)#Definindo que a variável formatacao é para ser igual a da célula A1 da primeira aba

for linha in range(2, todasaslinhas + 1):#Aqui ele percorre todas as linhas 2 para baixo
    Sala = abadados.cell(row=linha, column=3).value#identifica que a coluna sala é a terceira da planilha
    if not Sala:#se não existir sala então pare a execução do código
        break

    criaraba(Sala, planilha, formatacao)#se existir uma sala, a função de criar aba será iniciada e a de transferir as informações tambem
    abadestino = planilha[Sala]
    transinfo(abadados, abadestino, linha)

planilha.save("alunos_escola_ficticia_105(2).xlsx") #salva as alterações que o código fez na planilha 