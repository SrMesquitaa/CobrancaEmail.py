import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import datetime  

def obter_mes_atual():

    mes_atual = datetime.datetime.now().strftime('%B')  


    meses_pt = {
        "January": "Janeiro",
        "February": "Fevereiro",
        "March": "Março",
        "April": "Abril",
        "May": "Maio",
        "June": "Junho",
        "July": "Julho",
        "August": "Agosto",
        "September": "Setembro",
        "October": "Outubro",
        "November": "Novembro",
        "December": "Dezembro"
    }

    return meses_pt.get(mes_atual, mes_atual) 


lojas_gerentes = {
    '* e *': {
        'gerente': '*',
        'email': '*@gmail.com',
        'assunto': 'Metas Gerenciais - Arezzo e Magrella',
        'corpo': """<html>
                    <body>
                        <p>Olá *.</p>
                        <p>Segue a planilha para atualização das metas das lojas Arezzo e Magrella.</p>
                        <p>Por favor, preencha a planilha e nos envie de volta para manutenção das metas.</p>
                        <p>Obrigado e boas vendas!</p>
                        <p><i>Caso já tenha respondido este e-mail, por favor, ignore esta mensagem.</i></p>
                    </body>
                    </html>"""
    },
    'Vans': {
        'gerente': '*',
        'email': '**',
        'assunto': 'Metas Gerenciais - **',
        'corpo': """<html>
                    <body>
                        <p>Olá *.</p>
                        <p>Segue a planilha para atualização das metas da loja *.</p>
                        <p>Por favor, preencha a planilha e nos envie de volta para manutenção das metas.</p>
                        <p>Obrigado e boas vendas!</p>
                        <p><i>Caso já tenha respondido este e-mail, por favor, ignore esta mensagem.</i></p>
                    </body>
                    </html>"""
    }
}

arquivo_modelo = 'modelo_metas.xlsx'  




def enviar_email_com_anexo(destinatario, assunto, corpo, arquivo, remetente, senha):

    servidor = smtplib.SMTP('smtp.gmail.com', 587)
    servidor.starttls()

    try:

        servidor.login(remetente, senha)


        mensagem = MIMEMultipart()
        mensagem['From'] = remetente
        mensagem['To'] = destinatario
        mensagem['Subject'] = assunto  


        mensagem.attach(MIMEText(corpo, 'html'))  


        with open(arquivo, 'rb') as anexo:
            parte = MIMEBase('application', 'octet-stream')
            parte.set_payload(anexo.read())
            encoders.encode_base64(parte)
            parte.add_header('Content-Disposition',
                             f'attachment; filename={arquivo}')
            mensagem.attach(parte)


        servidor.sendmail(remetente, destinatario, mensagem.as_string())
        print(f"E-mail enviado com sucesso para {destinatario}!")

    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")
    finally:

        servidor.quit()



email_remetente = '**'
senha_remetente = '**'  


mes_atual = obter_mes_atual()


for loja, info in lojas_gerentes.items():

    assunto_atualizado = f"{info['assunto']} - {mes_atual}"
    enviar_email_com_anexo(info['email'], assunto_atualizado, info['corpo'],
                           arquivo_modelo, email_remetente, senha_remetente)
