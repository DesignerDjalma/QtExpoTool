import random
import pandas as pd

# Definindo as colunas da planilha
columns = ["ID", "MUNICIPIO", "SITUAÇÃO", "ANO", "NUMERO", "INTERESSADO", "IMOVEL", "PARCELA", "GEORREF"]

# Dados possíveis para preencher a planilha
municipios = ["Acará", "Afuá", "Altamira", "Ananindeua", "Belém", "Bragança", "Breves", "Cametá", "Capanema", "Castanhal"]
situacoes = ["Aprovado", "Pendente", "Negado", "Revisão", "Concluído"]
anos = [2018, 2019, 2020, 2021, 2022, 2023]
interessados = ["João Silva", "Maria Oliveira", "Pedro Sousa", "Ana Costa", "Lucas Lima", "Fernanda Almeida", "Carlos Pereira", "Juliana Santos", "Ricardo Gonçalves", "Sofia Machado"]
imoveis = ["Chácara", "Sítio", "Fazenda", "Terreno Urbano", "Apartamento", "Casa"]
parcelas = ["Parcela 1", "Parcela 2", "Parcela 3", "Parcela 4"]
georreferenciamento = ["S 01° 55' 44\" W 048° 29' 55\"", "S 02° 49' 12\" W 047° 28' 10\"", "S 01° 17' 32\" W 049° 12' 11\"", "S 03° 14' 22\" W 048° 55' 19\""]

# Gerando dados aleatórios
data = []

for i in range(30):
    row = [
        i + 1,  # ID
        random.choice(municipios),  # MUNICIPIO
        random.choice(situacoes),  # SITUAÇÃO
        random.choice(anos),  # ANO
        f"{random.randint(1000, 9999)}/{random.randint(2000, 2099)}",  # NUMERO
        random.choice(interessados),  # INTERESSADO
        random.choice(imoveis),  # IMOVEL
        random.choice(parcelas),  # PARCELA
        random.choice(georreferenciamento)  # GEORREF
    ]
    data.append(row)

# Criando o DataFrame
df = pd.DataFrame(data, columns=columns)

print(df)
