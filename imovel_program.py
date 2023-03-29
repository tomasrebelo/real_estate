import pandas as pd
import openpyxl
import os
import csv

# --------------------------------------- PART I - WRITING DATA TO A CSV FILE ----------------------------------------

# open the CSV file for appending
with open('C:/Users/trebelo/Python_vs_code/imovel_project.csv', 'a', newline='') as file:
    writer = csv.writer(file) # create a CSV writer object

    # add an empty row before writing the new data
    writer.writerow([])
    
    # read the last row in the file
    with open('C:/Users/trebelo/Python_vs_code/imovel_project.csv', 'r') as f:
        last_row = list(csv.reader(f))[-1]

    # calculate the new values for the next row
    valor_divida = int(last_row[0]) - int(last_row[2])
    mes_ordinal = int(last_row[1]) + 1
    prestacao_total = int(last_row[2])
    prestacao_metade = valor_divida/2

    # write the new row to the file
    writer.writerow([valor_divida, mes_ordinal, prestacao_total, prestacao_metade])


# --------------------------------------- PART II - WRITE CSV FILE DATA TO AN EXCEL FILE ----------------------------------------

# code to input from csv file to excel
def write_csv_to_excel(filepath:str, df:pd.DataFrame):
    filename = os.path.basename(filepath)
    df.to_excel(filename, index=False)

def main(outfile:str):
    df = pd.read_csv('C:/Users/trebelo/Python_vs_code/imovel_project.csv') # engine='openpyxl'
    write_csv_to_excel(outfile, df)

main(outfile='imovel.xlsx')
