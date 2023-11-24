# importando bibliotecas, bibliotecas sublinhas de vermelho, 
# precisam de pip (ser instaladas).
# as bibliotecas apos usadas ficam em negrito
import pandas as pd
import re # reconhece o email, se ele corresponde ao padrão, reconhece qualquer coisa!
import os
import smtplib
import openpyxl
import email
from email.message import EmailMessage
from time import sleep

#criando a classe do projeto - CLASSE > METODO. O OBJETO que chama os metodos da classe.
class To_do: # na classe o parentese é opicional. Aqui nos construirenos os métodos.
    def iniciar(self):
        self.lista_tarefas = [] # será alimentado com as tarefas digitadas
        self.email_destino()
        self.menu()
        self.criar_planilha()
        sleep(2) #para esperar
        self.enviar_email()

    def email_destino(self):
        while True:
            self.email = str(input('Email de destino: ')).lower() # O self permite acessar a variavel
            #dentro de outros metodos

            padrao_email = re.search(
                "^[a-zA-Z0-9._%+]+@[a-zA-Z0-9._-]+\.[a-z]{2,}$", self.email
            )
            if padrao_email:
                print('Email válido...')
                break
            else:
                print('Email inválido...')

            # https://pt.stackoverflow.com/questions/1386/express%C3%A3o-regular-para-valida%C3%A7%C3%A3o-de-e-mail
            # verifica se o email é aceitavel, 
            # se está dentro do padrão
            # var emailRegex = /^[a-z0-9.]+@[a-z0-9]+\.[a-z]+\.([a-z]+)?$/i;

    def menu(self):
        while True:
            menu_principal = int(input('''
            MENU PRINCIPAL
            [1] CADASTRAR
            [2] VISUALIZAR
            [3] SAIR              
            Opção: '''))

            match menu_principal:
                case 1: self.cadastrar_tarefas()
                case 2: self.visualizar_tarefas()
                case 3: break
                case _: print('Opção inválida')

    def cadastrar_tarefas(self):
        while True:
            self.tarefa = str(input('Tarefa ou [S] para sair: ')).capitalize() # a ideia do self é deixar ele disponivel

            if self.tarefa == 'S':
                break
            else:
                self.lista_tarefas.append(self.tarefa)
                try:
                    with open('./src/tarefas/todo.txt', 'a', encoding='utf8') as arquivo:
                        arquivo.write(f'{self.tarefa}\n')
                    # C:\Users\integral\Desktop\LanSchool Files\Projeto_Final\src\tarefas
                    # sempre inverter a barra
                    # src\tarefas\todo.txt
                except FileNotFoundError as erro:
                    print(f'Erro: {erro}')

    def visualizar_tarefas(self):
        try: #serve para mostrar a mensagem de erro e nao travar quando der erro
            with open('./src/tarefas/todo.txt', 'r', encoding='utf8') as arquivo:
                print(arquivo.read()) # para visualizar o conteudo .read()
                # self.lista_tarefas.append(arquivo.read()) #toda vez que o programa para, ele zera todo mundo
                # como a lista estará vazia, chamo o 2 e jogo as tarefas que estão no arquivo na lista de tarefas.
        except FileNotFoundError as erro:
            print(f'Erro: r ou rb')

    def criar_planilha(self):
        if len(self.lista_tarefas) > 0:
            try:
                df = pd.DataFrame({'Tarefas' : self.lista_tarefas}) #chave e valor
                nome_arquivo = str(input('Nome do arquivo (sem.XLSX): '))
                df.to_excel(nome_arquivo+'.xlsx', index=False)
                self.nome_planilha = nome_arquivo+'.xlsx'
                print('Planilha criada com sucesso')
            except Exception as e:
                print(f'Erro: criar_planilha')
    
    def enviar_email(self):
        endereco = 'meuemail@gmail.com' #alterar
        
        with open('./src/senha.txt') as arquivo:
            s = arquivo.readlines()
            #src\senha.txt

        senha = s[0]

        msg = EmailMessage()
        msg['Subject'] = "Planilha de tarefas"
        msg['From'] = endereco
        msg['To'] = self.email
        msg.set_content('Envio de planilha via Python')
        arquivos = [self.nome_planilha]  # lista
        for item in arquivos:
            with open(item, 'rb') as arq:
                dados = arq.read()                  
                nome_arquivo = arq.name

            msg.add_attachment(
                dados,
                maintype = 'application',
                subtype = 'octet-stream',
                filename = nome_arquivo
            )
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(endereco, senha, initial_response_ok=True)
        server.send_message(msg) 
        print('Email enviado com sucesso')
        # https://myaccount.google.com/apppasswords
        # senha criada unica vez: https://myaccount.google.com/apppasswords...
        # requer autenticação de 2 fatores
              




# A classe será chamada da seguinte maneira: 
# A variavel start é um objeto da classe to_do, 
# ela pode receber os metodos, 
# que no caso é o # iniciar. 
# Por ser objeto, ela terá acesso a todos os métodos da classe. 
start = To_do()
start.iniciar()