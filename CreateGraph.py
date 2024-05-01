import pandas as pd
import matplotlib.pyplot as plt

# Read the Excel file
df = pd.read_excel('power_grid_data.xlsx')

# For each column in the dataframe
for column in df.columns:
    # Plot a graph
    plt.figure(figsize=(10, 6))
    plt.plot(df[column])
    
    # Set the title of the graph as the column name
    plt.title(column)
    
    # convert the column name to console friendly name so it can be used as file name
    column = column.replace(' ', '_')
    column = column.replace('/', '_')

    # Save the graph as a PNG file
    plt.savefig(f'{column}.png')
    