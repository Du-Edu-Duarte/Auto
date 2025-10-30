import openpyxl
import pyperclip
import pyautogui
from time import sleep
# entrar na planilha 

workbook = openpyxl.load_workbook("produtos_ficticios.xlsx")
sheet_produtos = workbook['Produtos']
# copiar informações de um campo e colar no seu campo corrMaster CalçaExtra Camisaespondente 
for linha in sheet_produtos.iter_rows(min_row=2):

    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(3526,315,duration=1)
Extra Bolsa    pyautogui.hotkey("ctrl", "v")

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(3534,431,duration=1)
Última moda da estação.    pyautogui.hotkey("ctrl", "v")

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(3541,559,duration=1)
    pyautogui.hotkey("ctrl", "v")

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(3535,663,duration=1)
    pyautogui.hotkey("ctrl", "v")


    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(3551,760,duration=1)
    pyautogui.hotkey("ctrl", "v")


    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(3531,758,duration=1)
    pyautogui.hotkey("ctrl", "v")

    pyautogui.click(3489,923,duration=1)
    sleep(2)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(3519,339,duration=1)
    pyautogui.hotkey("ctrl", "v")

    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(3519,440,duration=1)
    pyautogui.hotkey("ctrl", "v")

    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(3506,540,duration=1)
    pyautogui.hotkey("ctrl", "v")

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(3500,634,duration=1)
    pyautogui.hotkey("ctrl", "v")

    tamanho = linha[10].value
    if tamanho == "Pequeno":
        pyautogui.click(3500,765,duration=1)
    elif tamanho == "Médio":
        pyautogui.click(3515,801,duration=1)
    else:
        pyautogui.click(3514,831,duration=1)
    
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(3529,829,duration=1)
    pyautogui.hotkey("ctrl", "v")

    pyautogui.click(3515,891,duration=1)
    sleep(3)
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(3519,366,duration=1)
    pyautogui.hotkey("ctrl", "v")

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(3515,459,duration=1)
    pyautogui.hotkey("ctrl", "v")

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(3528,568,duration=1)
    pyautogui.hotkey("ctrl", "v")

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(3524,709,duration=1)
    pyautogui.hotkey("ctrl", "v")

    localizacao = linha[16].value
    pyperclip.copy(localizacao)
    pyautogui.click(3534,801,duration=1)
    pyautogui.hotkey("ctrl", "v")

    pyautogui.click(3525,871,duration=1)

    pyautogui.click(4030,219,duration=1)

# repetir esses passos para outros campos até preencher campos daquela página 
# clicar próximo 
# repetir os mesmo passos e ir para a próxima página (pagina 2)
# repetir os mesmos passos e finalizar o cadastro daquele produto  e clicar em concluir
# clicar em ok para finalizar o processo
# clicar no ok mais uma vez na mensagem de confirmação de salvamento no banco de dados 
# clicar em adicionar mais um e repetir o processo até finalizar a planilha  
