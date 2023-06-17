from sqlalchemy.engine import URL
connection_string = '''DRIVER={SQL Server};SERVER=10.33.30.20;UID=antoniogusmao;Pwd=Lucas1144675167;DATABASE=CorporeRM'''
connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": connection_string})
import pyodbc
from sqlalchemy import create_engine
engine = create_engine(connection_url)
import pandas as pd
import sqlalchemy as sa



with engine.begin() as conn:
    df = pd.read_sql(sa.text('''
    DECLARE @setedias datetime;
    SET @setedias = DATEADD(DD,DATEDIFF(DD,0,GETDATE()),0)+7
    SELECT 
		CASE 
			WHEN NOMEFANTASIA = 'CEPT-CAXIAS' THEN 'CEPT- Caxias'
			WHEN NOMEFANTASIA = 'CEPT-RFT' THEN 'CEPT- Raimundo Franco Texeira'
			WHEN NOMEFANTASIA = 'CEPT-DI' THEN 'CEPT- Distrito Industrial'
			WHEN NOMEFANTASIA = 'CEPT-IMPERATRIZ' THEN 'CEPT- Imperatriz'
			WHEN NOMEFANTASIA = 'CEPT-BACABAL' THEN 'CEPT- Bacabal'
			WHEN NOMEFANTASIA = 'CEPT-BALSAS' THEN 'CEPT- Balsas'
			WHEN NOMEFANTASIA = 'CEPT-ROSARIO' THEN 'CEPT- Rosario'
			WHEN NOMEFANTASIA = 'CEPT-AÇAILÂNDIA' THEN 'CEPT- Acailandia'
		END AS FILIAL,
        CONVERT(VARCHAR,STURMA.DTINICIAL,103)[DATA INICIAL],
        CONVERT(VARCHAR,STURMA.DTFINAL,103)[DATA FINAL],
        SCURSO.NOME AS CURSO,
        SMOD.DESCRICAO[MODALIDADE],
        STURMA.CODTURMA AS CODTURMA
    FROM    dbo.STURMA (NOLOCK)
            INNER JOIN dbo.SPLETIVO (NOLOCK) 
            ON dbo.SPLETIVO.CODCOLIGADA = dbo.STURMA.CODCOLIGADA
            AND dbo.SPLETIVO.IDPERLET = dbo.STURMA.IDPERLET
            INNER JOIN dbo.GFILIAL (NOLOCK) 
            ON dbo.GFILIAL.CODCOLIGADA = dbo.SPLETIVO.CODCOLIGADA
            AND dbo.GFILIAL.CODFILIAL = dbo.STURMA.CODFILIAL
            INNER JOIN dbo.SHABILITACAOFILIAL (NOLOCK) 
            ON dbo.SHABILITACAOFILIAL.IDHABILITACAOFILIAL = dbo.STURMA.IDHABILITACAOFILIAL
            AND dbo.SHABILITACAOFILIAL.CODCOLIGADA = dbo.STURMA.CODCOLIGADA
            INNER JOIN dbo.SCURSO (NOLOCK) 
            ON dbo.SCURSO.CODCOLIGADA = dbo.SHABILITACAOFILIAL.CODCOLIGADA
            AND dbo.SCURSO.CODCURSO = dbo.SHABILITACAOFILIAL.CODCURSO
            INNER JOIN SMODALIDADECURSO SMOD
            ON SMOD.CODCOLIGADA = SCURSO.CODCOLIGADA
            AND SMOD.CODMODALIDADECURSO = SCURSO.CODMODALIDADECURSO
    WHERE   dbo.STURMA.CODCOLIGADA = 3
            AND DTINICIAL <= CONVERT(DATE, GETDATE())
            AND DTFINAL = @setedias
            AND SMOD.DESCRICAO NOT IN ('Aperfeiçoamento/Especialização Profissional')
            AND SMOD.DESCRICAO NOT IN ('Iniciação Profissional')'''), conn)
    data = pd.DataFrame(df)
    data.to_excel(r"C:\Users\antoniolucas\Documents\SAPES PYTHON\GT.xlsx",sheet_name='Sapes')
print(pyodbc.drivers())