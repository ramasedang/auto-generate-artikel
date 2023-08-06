from docx2pdf import convert
import os

# convert all .docx files in the artikel folder to .pdf but dont delete the .docx files bedakan folder output dan input
# buat folder output dan input
# laukan pengecekan apakah nama file sudah ada di folder output jika sudah ada maka tidak perlu di convert
# jika belum ada maka lakukan convert

input_dir = "artikel"
output_dir = "artikel_pdf"
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

while True:
    for filename in os.listdir(input_dir):
        if filename.endswith(".docx"):
            if not os.path.exists(output_dir + "/" + filename.replace(".docx", ".pdf")):
                convert(
                    input_dir + "/" + filename,
                    output_dir + "/" + filename.replace(".docx", ".pdf"),
                )
                print("Convert " + filename + " to pdf")
            else:
                print("File " + filename + " sudah ada")
    print("Selesai")
    break
