#!/usr/bin/env bash

# Instala a dependência externa 'poppler-utils' necessária para o pdf2image
sudo apt-get update && sudo apt-get install -y poppler-utils

# Instala as dependências do Python
pip install -r requirements.txt

# Inicia a aplicação
python main.py
