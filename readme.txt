Após concluir as alterações, execute o comando 'pyinstaller --onefile main.py'
Abra o arquivo main.spec
Localize o paramentro hiddenimports=[], 
preencha com :hiddenimports=['sqlalchemy.sql.default_comparator']
excute novamente o camando 'pyinstaller --one....'