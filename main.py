import pandas as pd
import smtplib
import email.message
import email
import pandas as pd
import unidecode
import time 
import gspread
import pandas as pd

#Utiliazando gspread para adicionar informações nas linhas da tabela
CODE = "1stZuHZP9yFWdoIDO2SIPzSM-VD3Vyng4qjV6cJxtjLI"
gc = gspread.service_account(filename=r'C:\Users\antoniolucas\OneDrive - Sistema FIEMA\Documentos\PYTHON\automacao de envio de email para os interrassado no cursos do senai\chave.json')
sh = gc.open_by_key(CODE)
ws = sh.worksheet('pagina1')
#Enviar Email
def enviar_email(nome,user_email,curso):
    corpo_email = f"""
    <p>Prezado {nome},
    Recebemos sua solicitação para se iniciar o curso de  na marca de 7 dias restantes para o término da turma é necessário que esta seja monitorada para a realização da 1ª Fase da pesquisa do SAPES, considerando que a execução da 2ª e 3ª Fase dependem disto.  

    <p>Identificamos que a turma  tem previsão de encerramento em <b></b>.<p>
    <p>Aguardamos devolutiva até <b></b><p>
    
    <p>Ratificamos a importância, de realização do monitoramento contínuo no sistema SAPES e atualização dos dados dos alunos no sistema SGE.</p>
    <p>Favor responder para o email <b>josuebarreto@fiema.org.br</b><p>

    """
    msg = email.message.Message()
    msg['Subject'] = f'⚠️ Sua inscrição no curso foi confirmada no curso {curso} ⚠️'
    msg['From'] = 'ia.sapes@ma.edu.senai.br'
    msg['To'] = user_email
    password = 'klythyhnumthjqki' 
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email)
    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    # Login Credentials for sending the mail
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print(f'email enviado para {email}')

###################################################################################################
#Entrada de Dados
id_forms = "1stZuHZP9yFWdoIDO2SIPzSM-VD3Vyng4qjV6cJxtjLI"
cont = 0
leitura_forms = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{id_forms}/export?format=csv")
leitura_forms
for i in leitura_forms.index:
   nome = leitura_forms.loc[i,'Nome Completo:']
   telefone = leitura_forms.loc[i,'Telefone:(Whatssap)']
   user_email = leitura_forms.loc[i,'Email:']
   curso = leitura_forms.loc[i,'Selecione o Curso:']
   status_email = leitura_forms.loc[i,'Status email']
   cont +=1
#1 CASO 
# É Quando o usuario ainda não recebeu o email e o campo status email está vazio
#2 CASO
# É Quando o usuario ja recebeu o email e o campo status email está prenchido
   if status_email !='email enviado':
        print(f"Nome:{nome}\nTelefone:{telefone}\ne-mail:{user_email}\nCurso:{curso}")
        lista = ['email enviado']
        print(cont)
        ws.append_row(lista,table_range=f"H{cont}")
        enviar_email(nome,user_email,curso)
   else:
        print("Todos os candidatos ja receberam o aviso por email")
