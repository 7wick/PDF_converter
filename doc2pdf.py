import comtypes.client

in_file = input("Enter input file name along with path: ")  # C:/Users/wick7/Desktop/test/DOCX.docx
try:
    if int(input("Press 1 if you want to save file with same name and path? Else press any other number : ")) == 1:
        out_file = in_file.replace((in_file.rsplit('.', 1)[1]), "pdf")
    elif int(input(
            "Press 1 if you want to save file with same path but different name? Else press any other number : ")) == 1:
        out_file = in_file.rsplit('/', 1)[0] + "/" + input("Name of output file: ") + ".pdf"
    else:
        out_file = input("Path of output file: ") + "/" + input("Name of output file: ") + ".pdf"
except:
    print("Invalid entry!")
try:
    word = comtypes.client.CreateObject('Word.Application')
except ImportError:
    print("Word.Application missing on your system")
doc = word.Documents.Open(in_file)
try:
    export_no = int(input("Press 0 to export the entire document or Press 2 to export the current page or "
                          "\nPress any other number to export within a range:\n"))
    if export_no == 0 or export_no == 2:
        doc.ExportAsFixedFormat(out_file, ExportFormat=17, Range=export_no)
    else:
        start = int(input("Enter start page: "))
        end = int(input("Enter end page: "))
        doc.ExportAsFixedFormat(out_file, ExportFormat=17, Range=3, From=start, To=end)
except Exception as e:
    print("Invalid entry!")
doc.Close()
