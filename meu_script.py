import mysql.connector

# Conectando ao banco de dados
mydb = mysql.connector.connect(
  host="localhost",
  user="root",
  password="admin2804" # Substitua pela sua senha
)

print(mydb)

# Você pode então criar um cursor para executar consultas
mycursor = mydb.cursor()
mycursor.execute("SHOW DATABASES")

for x in mycursor:
  print(x)
