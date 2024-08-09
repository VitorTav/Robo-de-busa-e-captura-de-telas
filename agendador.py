import schedule
import time
from scriptFiscalteste import fiscal
from datetime import datetime


# Executa a função fiscal imediatamente




# Agendamento da função job
# schedule.every().day.at("17:08").do(fiscal)

schedule.every().day.at("17:32").do(fiscal)
# Loop para manter o agendador rodando
while True:
    schedule.run_pending()
    print("Esperando pela próxima execução...")
    time.sleep(1)  # Reduzido para 1 segundo para melhor responsividade
