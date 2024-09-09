import openpyxl

# Define variable to load the dataframe
dataframe = openpyxl.load_workbook("./DPT KECAMATAN-KELURAHAN_TPS 6.pdf.xlsx")

# Define variable to read sheet
dataframe1 = dataframe.active

# Iterate the loop to read the cell values
read = False
for row in range(0, dataframe1.max_row):
    data = []
    for col in dataframe1.iter_cols(1, dataframe1.max_column):
        if col[row].value != None and str(col[row].value).strip() == "NAMA":
            read = True
        elif "Rekapitulasi" in str(col[row].value).strip():
            read = False
        elif read and col[row].value != None and str(col[row].value).strip() != "JENIS KELAMIN":
            data.append(col[row].value)


    if read:
        print(data)

