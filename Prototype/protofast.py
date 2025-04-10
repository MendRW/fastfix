import pandas as pd

def main():
    # Define filters
    colfilter = []  # put the index of the columns you want to pull
    resolvedby = ["dean.lynch"]  # filter out other teams by choosing which person resolved ticket
    

    df = pd.read_csv('your_file.csv')
     
    filtered_df = df.iloc[:, colfilter]
    filtered_df = filtered_df[filtered_df['ResolvedBy'].isin(resolvedby)]
    
    # Save to new CSV
    filtered_df.to_csv('new_file.csv', index=False)


if __name__ == "__main__":
    main()