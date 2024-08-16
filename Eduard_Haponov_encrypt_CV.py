import aspose.words as aw

password = "qwerty123"
file_name = "cv.docx"

try:
    doc = aw.Document(file_name)
    options = aw.saving.OoxmlSaveOptions(aw.SaveFormat.DOCX)
    options.password = password
    doc.save(file_name, options)
    print(f"The {file_name} has been successfully encrypted with the password and saved.")
except Exception as e:
    print(f"An error occurred: {e}")
