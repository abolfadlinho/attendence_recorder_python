import tabula


def PDFtoXLSX():
    filename = input('Input PDF filename with extension:')
    df = tabula.read_pdf(filename, pages='all')[0]
    newfilename = (filename[:-3]) + "xlsx"
    df.to_excel(newfilename)
    return newfilename
