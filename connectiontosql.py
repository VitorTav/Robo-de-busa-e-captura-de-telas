import mysql.connector
from mysql.connector import Error
from datetime import date
import sys



def create_server_connection_mysql(host_name, user_name, user_password):
    connection = None
    try:
        connection = mysql.connector.connect(
            host=host_name,
            user=user_name,
            passwd=user_password
        )
        print("MySQL Database connection successful")
    except Error as err:
        print(f"Error: '{err}'")
        sys.exit() 
    return connection

def read_query(query):
    global connection
    
    cursor = connection.cursor()
    result = None
    try: 
        cursor.execute(query)
        result = cursor.fetchall()
        return result
    except Error as err:
        print(f"Error: '{err}'")
        sys.exit()

def insert_mysql(conexao,sql,valores):
    cursor = conexao.cursor()
    # Executa a inserção
    cursor.execute(sql, valores)
    # Faz o commit da transação
    conexao.commit()
    # Fecha o cursor e a conexão
    cursor.close()

def update_mysql(conexao, sql, valores):
    cursor = conexao.cursor()
    # Executa a atualização
    cursor.execute(sql, valores)
    # Faz o commit da transação
    conexao.commit()
    # Fecha o cursor e a conexão
    cursor.close()


def setMonitoramentoHoje(nome):
    global existeMonitoramentoHoje,roboFuncionouHoje
    data_Hoje = date.today().strftime("%Y-%m-%d")
    sql =  "select count(*),"+nome+" from robos.monitoramento_importacao_conect where data = '"+data_Hoje+"' group by "+nome
    consulta = read_query(sql) 
    
    existeMonitoramentoHoje = consulta[0][0] if consulta and consulta[0] and consulta[0][0] != 0 else False
    roboFuncionouHoje = consulta[0][1] if consulta and consulta[0] and consulta[0][1] != 0 else False

def atualiza_log_monitoramento(nome):
    setMonitoramentoHoje(nome)

    data_Hoje = date.today().strftime("%Y-%m-%d")
    if(existeMonitoramentoHoje):

        if(roboFuncionouHoje in ("S","N")):
            sql = "update robos.monitoramento_importacao_conect set "+nome+" = %s where data = '"+data_Hoje+"'"
            valores = ['S']
            update_mysql(connection,sql,valores)
    else:
        sql = "insert into robos.monitoramento_importacao_conect (data,"+nome+") values (%s,%s)"
        data_Hoje = date.today().strftime("%Y-%m-%d")
        valores = [data_Hoje,'S']
        insert_mysql(connection,sql,valores)


connection = create_server_connection_mysql('sua porta', 'seu usuario', 'sua senha')