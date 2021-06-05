from docx import Document
from docx2pdf import convert
import pandas as pd

df = pd.read_csv('list.csv')
path = 'Your Path'
for name in df.Names:
    document = Document('Template.docx')
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        if 'Here' in run.text:
                            run.text = run.text.replace("Here", name)

    document.save(path + '/certificates-word/' + name + '.docx')
    convert(path + '/certificates-word/' + name + '.docx',
            path + '/certificates-pdf/' + name + '.pdf')
