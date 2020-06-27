import numpy as np
import pylatex as pl
import pandas as pd
import subprocess

####################### Arguments #######################
file_name = 'report'
csv_file = 'data.csv'
threshold = 0
latex_path = 'C:/Users/ryan.behdad/AppData/Local/Programs/MiKTeX 2.9/miktex/bin/x64/'

####################### Main #######################
df = pd.read_csv(csv_file)
df['Double CF'] = df['CF'] * 2
if df['Double CF'].sum() > threshold:
    sentence = 'The project has a profit of ' + str(df['Double CF'].sum())
else:
    sentence = 'The project has a loss of ' + str(-1 * df['Double CF'].sum())

doc = pl.Document()
doc.packages.append(pl.Package('booktabs'))

with doc.create(pl.Section('Results')):
    with doc.create(pl.Table(position='htbp')) as table:
        table.add_caption('Test')
        table.append(pl.Command('centering'))
        table.append(pl.NoEscape(df.to_latex(escape=False, index=False)))
    
    doc.append(sentence)

def create_pdf(filename):
    input_filename = file_name + '.tex'
    output_filename = file_name
    process = subprocess.Popen([
        latex_path + 'latex',
        '-output-format=pdf',
        '-job-name=' + output_filename,
        input_filename])
    process.wait()

doc.generate_tex(file_name)
create_pdf(file_name)
