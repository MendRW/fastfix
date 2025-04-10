import pandas as pd
colfilter = [] # put the index of the columns you want to pull
resolvedby = ["dean.lynch"] #filter out other teams by choosing which person resolved ticket
df = pd.read_csv() #reads the csv file

def main():
    df = (df.iloc[:, colfilter]) & (df [resolvedby])
    df.to_csv('new_file.csv', index=False)

