import pandas as pd

# Supongamos que df es tu DataFrame
df = pd.DataFrame({'Nombre': ['Juan', 'María', 'Carlos'],
                   'Edad': [25, 30, 22],
                   'Ciudad': ['México', 'Madrid', 'Buenos Aires']})

# Convertir el DataFrame a JSON y guardarlo en un archivo
df.to_json('output.json', orient='records', lines=True)

# Imprimir mensaje indicando que se ha guardado el archivo
print("DataFrame convertido a JSON y guardado en 'output.json'")