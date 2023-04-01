import gspread
import pandas as pd
CODE = "1stZuHZP9yFWdoIDO2SIPzSM-VD3Vyng4qjV6cJxtjLI"
gc = gspread.service_account(filename='chave.json')
sh = gc.open_by_key(CODE)
ws = sh.worksheet('pagina1')
ws.get_all_records()
lista = ['email enviado por gspread']
cont = 6
tabela = pd.DataFrame(ws.get_all_records())
ws.append_row(lista,table_range=f"H{cont}")
