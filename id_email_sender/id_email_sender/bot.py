from botcity.core import DesktopBot
import pandas as pd

class Bot(DesktopBot):
    def action(self, execution=None):
        
        tabela = pd.read_excel(str(r'C:\Users\efranca\Documents\titan\id_email_sender\id_email_sender\database_gmail_sender.xlsx'), dtype=str)

        self.browse("https://titan.hostgator.com.br/mail/")
        
        for i in tabela.index:

            email = str(tabela.loc[i, 'email'])
            subject = 'Novidades - ALELO Natal'
            message = """Olá somos a ALELO e temos Novidades!!!

Estamos próximos do Natal.Você já pensou em transformar a expectativa do seu colaborador em realidade?

Com o Cartão Natal, aceito em toda a MultiRede Alelo, sua equipe celebra o fim do ano do jeito que quiser: em restaurante, abastecendo o carro ou em supermercados.

A ALELO possui vários Benefícios para a sua Empresa, como Taxas e Condições Diferenciadas, que nós da CONTACT, como Parceiro Exclusivo da ALELO, podemos oferecer.

Faça uma Consulta e entre em Contato Conosco."""
        
            print('====Inicio====')
            print(f'email: {email}')
            print(f'subject: {subject}')
            print(f'message: {message}')
            print('==============')

            if self.find( "writtingFind", matching=0.97, waiting_time=30000):
                self.not_found("writtingFind")
                self.click()
            self.sleep(1000)

            if self.find( "newEmail", matching=0.97, waiting_time=10000):
                self.not_found("newEmail")

            self.paste(email)
            self.sleep(500)
            self.tab()
            self.paste(subject)
            self.sleep(500)
            self.tab()
            self.paste(message)
            self.sleep(500)
                       
            if self.find( "sendFind", matching=0.97, waiting_time=10000):
                self.not_found("sendFind")
            self.click()
            self.sleep(1440)

        print('====Fim====')
            
    def not_found(self, label):
        print(f"Element found: {label}")

if __name__ == '__main__':
    Bot.main()
