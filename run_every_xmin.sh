#!/bin/bash

while true; do
    # Execute your Python script
    echo ""
    echo "Inicio"
    echo "Inicio de otro ciclo" >> logfile.log 2>&1
    date
    date >> logfile.log 2>&1
    python main.py >> logfile.log 2>&1
    date
    date >> logfile.log 2>&1
    echo "FIN"
    echo "FIN de un ciclo" >> logfile.log 2>&1

    echo " " >> logfile.log 2>&1

    # Sleep for 30 minutes (1800 seconds)
    # Sleep for 20 minutes (1200 seconds)
    # Sleep for 15 minutes (900 seconds)
    # Sleep for 08 minutes (450 seconds)
    # Sleep for 3.5minutes (200 seconds)
    sleep 200
done
