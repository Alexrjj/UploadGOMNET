"""
Exporta o nome de arquivos em pastas / subpastas para "sobs.txt"
Obs.: Determinar o caminho da pasta a serem extraídos os nomes ou manter ('./') para extrair da pasta atual onde está o script.
"""

import os

with open("sobs.txt", "w") as a:
    for path, subdirs, files in os.walk('./'):
       for filename in files:
         f = os.path.join(filename.partition(".")[0])
         a.write(str(f) + "\n")
