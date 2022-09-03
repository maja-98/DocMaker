import docx
def docReader(doc_name):
    doc=docx.Document(doc_name)
    headings = list(map(lambda x:x.text,doc.paragraphs))
    return (headings)

