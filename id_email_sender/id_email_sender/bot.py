from botcity.core import DesktopBot
import pandas as pd

class Bot(DesktopBot):
    def action(self, execution=None):
        
        tabela = pd.read_excel(str(r'C:\Users\efranca\Documents\titan\id_email_sender\id_email_sender\database_gmail_sender.xlsx'), dtype=str)

        self.browse("https://titan.hostgator.com.br/mail/")
        
        for i in tabela.index:

            email = str(tabela.loc[i, 'email'])
            subject = str(tabela.loc[i, 'subject'])
            message = str(tabela.loc[i, 'message'])
        
            print('====Inicio====')
            print(f'email: {email}')
            print(f'subject: {subject}')
            print(f'message: {message}')
            print('==============')

            if self.find( "writtingFind", matching=0.97, waiting_time=30000):
                self.not_found("writtingFind")
                self.click()
                self.wait(1000)

            if self.find( "newEmail", matching=0.97, waiting_time=10000):
                self.not_found("newEmail")

            self.paste(email)
            self.wait(500)
            self.tab()
            self.paste(subject)
            self.wait(500)
            self.tab()
            self.paste(message)
            self.wait(500)
                       
            if self.find( "sendFind", matching=0.97, waiting_time=10000):
                self.not_found("sendFind")
            self.click()

        print('====Fim====')
            
    def not_found(self, label):
        print(f"Element found: {label}")

if __name__ == '__main__':
    Bot.main()



