import pandas as pd

df = pd.read_excel("sales.xlsx")

unpaid_clients = df[df['pay by date'] < pd.to_datetime('today')]
