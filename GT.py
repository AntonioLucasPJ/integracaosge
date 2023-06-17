from calendar import month
from dataclasses import replace
from operator import index
from tkinter.messagebox import YES
from traceback import print_tb
import smtplib
import email.message
import email
import pandas as pd
import unidecode
from time import sleep as sono
from datetime import datetime,timedelta
#Lista de cursos que não serão enviados
turmas_aperferçoamento =[]
turmas_iniciacaoprofissional=[]
#Lista de Colaboradores 
cont =0
colaboradores= ['antoniolucas@fiema.org.br','mariarodrigues@fiema.org.br']
geren_rft = ['doralice@fiema.org.br','ritasousa@fiema.org.br']
geren_di = ['clezenildesales@fiema.org.br','wennedy@fiema.org.br']
geren_imperatriz = ['deglisonnascimento@fiema.org.br','erikabrito@fiema.org.br']
geren_acai = ['rosimeirelima@fiema.org.br']
geren_bals = ['josenazareno@fiema.org.br','lenisalacerda@fiema.org.br']
geren_baca = ['andredossantos@fiema.org.br','josearimatea@fiema.org.br']
geren_caxi = ['wilberthraiol@fiema.org.br','denise@fiema.org.br']
geren_rosa = ['josilenemos@fiema.org.br']
#Introdução a Data de Hoje
hoje = datetime.now()
#Remover horas da data e transformar AAA/M/D
em_trez_dias = hoje + timedelta(days=2)
em_trez_dias = em_trez_dias.strftime('%d/%m/%Y')
def enviar_email(unidade,c_curso,cod_turma,f_turmas,lista_n):
    corpo_email = f"""
    <p>Prezado time {unidade},

Como é da ciência deste Centro, na marca de 7 dias restantes para o término da turma é necessário que esta seja monitorada para a realização da 1ª Fase da pesquisa do SAPES, considerando que a execução da 2ª e 3ª Fase dependem disto.  

    <p>Identificamos que a turma {cod_turma} tem previsão de encerramento em <b>{f_turmas}</b>.<p>
    <p>Aguardamos devolutiva até <b>{em_trez_dias}</b><p>
    
    <p>Ratificamos a importância, de realização do monitoramento contínuo no sistema SAPES e atualização dos dados dos alunos no sistema SGE.</p>
    <p>Favor responder para o email <b>mariarodrigues@fiema.org.br</b><p>

    """
    msg = email.message.Message()
    msg['Subject'] = f'⚠️ AVISO DE ENCERRAMENTO DA TURMA:{cod_turma} ⚠️'
    msg['From'] = 'ia.sapes@ma.edu.senai.br'
    msg['To'] = lista_n
    password = 'klythyhnumthjqki' 
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email)
    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    # Login Credentials for sending the mail
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print(f'Turma:{cod_turma}, email enviado para {lista_n}')
#Planilha
tabela = pd.read_excel(r'C:\Users\antoniolucas\Documents\SAPES PYTHON\GT.xlsx')
for i in tabela.index:
    filial = tabela.loc[i,'FILIAL']
    curso = tabela.loc[i,'CURSO']
    cdt = tabela.loc[i,'CODTURMA']
    modalidade = tabela.loc[i,'MODALIDADE']
    data_final = tabela.loc[i,'DATA FINAL']
#Tratamento de datas
# and data_final <=p_semana:
    cont+=1
    unidade = filial
    cod_turma = cdt
    curso = unidecode.unidecode(curso)
    t_curso = len(curso)
    f_turmas = data_final
    c_curso = curso
    m_modalidade = modalidade
#Foi socilitado que as turmas de Aperfeiçoamento profissional e iniciação profissional não sejam enviados 
    print(f"Nome:{unidade} \nCodigo Turma:{cod_turma}\nModalidade:{m_modalidade}\nData Final:{f_turmas}")
    for n in colaboradores:
        lista_n = n
        enviar_email(unidade,c_curso,cod_turma,f_turmas,lista_n)
        sono(2)
    if unidade == 'CEPT- Raimundo Franco Texeira':
        for t in geren_rft:
            lista_h = t
            enviar_email(unidade,c_curso,cod_turma,f_turmas,lista_h)
            sono(2)
        print(f"Quantidadade de Turmas no RFT:{cont} turmas")
    if unidade == 'CEPT- Distrito Industrial':
        for u in geren_di:
            lista_i = u 
            enviar_email(unidade, c_curso, cod_turma, f_turmas, lista_i)
            sono(2)
    if unidade == 'CEPT- Imperatriz':
        for g in geren_imperatriz:
            lista_j = g
            enviar_email(unidade, c_curso, cod_turma, f_turmas, lista_j)
            sono(2)
    if unidade == 'CEPT- Acailandia':
        for f in geren_acai:
            list_k = f
            enviar_email(unidade, c_curso, cod_turma, f_turmas, list_k)
            sono(2)
    if unidade == 'CEPT- Bacabal':
        for k in geren_baca:
            list_l = k 
            enviar_email(unidade, c_curso, cod_turma, f_turmas, list_l)
            sono(2)
    if unidade == 'CEPT- Balsas':
        for n in geren_bals:
            list_m = n
            enviar_email(unidade, c_curso, cod_turma, f_turmas, list_m)
            sono(2)
    if unidade == 'CEPT- Caxias':
        for n in geren_caxi:
            list_n = n 
            enviar_email(unidade, c_curso, cod_turma, f_turmas, list_n)
            sono(2)
    if unidade == 'CEPT- Rosario':
        for m in geren_rosa:
            list_o = m
            enviar_email(unidade, c_curso, cod_turma, f_turmas, list_o)
            sono(2)
print(f'Total de Turmas ={cont}')
print(f'Existem {len(turmas_iniciacaoprofissional)} turmas de Iniciação Profissional e {len(turmas_aperferçoamento)} turmas de Aperfeiçoamento/Especialização Profissional')


            
 