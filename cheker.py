import pandas as pd
import os
import re


def clean_filename(filename):
    # Regex pattern to remove characters that are not allowed in filenames
    pattern = r'[\\/*?:"<>|!]+'
    return re.sub(pattern, "", filename)


# list all file names in the directory artikel
filenames = os.listdir("artikel")
# create a list of filenames without the extension
filenames = [os.path.splitext(filename)[0] for filename in filenames]
# print(filenames)

# open the xlsx file
df = pd.read_excel("input.xlsx")
print(df)

# Clean the "Judul" column before comparing
df["Clean_Judul"] = df["Judul"].apply(clean_filename)

# compare the filenames in the directory with the filenames in the xlsx file
# jika nama file di folder artikel ada di kolom "Clean_Judul", maka di kolom "generate" diisi 1, jika tidak ada maka kosongkan cellnya
df["generate"] = df["Clean_Judul"].apply(lambda x: 1 if x in filenames else "")

# hitung jumlah 1 di kolom generate
print(df["generate"].eq(1).sum())

# Remove the temporary 'Clean_Judul' column
df.drop("Clean_Judul", axis=1, inplace=True)

# Save the modified DataFrame to a new Excel file
df.to_excel("input.xlsx", index=False)
