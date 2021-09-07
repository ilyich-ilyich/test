import pandas as pd
cols = [0, 2, 3]

top_players = pd.read_excel('./top_players.xlsx', usecols=cols)
top_players.head()