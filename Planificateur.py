# Planificateur.py

import schedule
import time
from actions import action

def planifier_action(callback_func, interval_minutes):
    # Planifier l'action avec la fonction de rappel toutes les X minutes
    schedule.every(interval_minutes).minutes.do(callback_func)

# Appeler la fonction pour planifier l'action avec la fonction spécifique
planifier_action(action, 5)  # Planifier l'action toutes les 5 minutes

while True:
    schedule.run_pending()
    time.sleep(1)