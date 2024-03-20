#!/bin/bash

echo 'Serviço de monitoramento de diretorio iniciado em '

inotifywait -m -r -e 'close_write,moved_to' --format '%w%f' /home/pdfconvert/Input/ | while read file

do

/usr/bin/python3 /opt/pdfcut.py

rm -rf /home/pdfconvert/Input/*
done
