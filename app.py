import tabula
from functions.functions import *

# Leia o arquivo PDF
tables = tabula.read_pdf("files/testeocr.pdf", pages="all")

# Chamando a função que trata as colunas
format_cols(tables)
