import tkinter as tk
from tkinter import filedialog
import pandas as pd
from sqlalchemy import create_engine
from sqlalchemy.engine import URL

App = tk.Tk()
fichero=filedialog.askopenfile(title="Abrir", defaultextension="xlsx")
print(fichero.name)

#MSSQL
connection_string = "DRIVER={SQL Server Native Client 11.0};SERVER=DELL-LHP;DATABASE=LUIS;UID=sa;PWD=123456"
connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": connection_string})
engine = create_engine(connection_url)


#Mete un excel en el DataFrame y lo manda a una tabla en MSSQL
datos = pd.read_excel(fichero.name,header=0)
df=pd.DataFrame(data=datos)
df.to_sql(name="datos", con=engine, if_exists="replace", index=False)


sql = "select * from datos"
df = pd.read_sql_query(sql, con=engine)

fichero=filedialog.asksaveasfile(title="Guardar")

#newer xlsx file
df.to_excel(fichero.name, index=False, index_label=None)


App.mainloop()