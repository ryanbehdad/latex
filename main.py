import numpy as np
import pylatex as pl
import pandas as pd
import subprocess

####################### Arguments #######################
file_name = 'report'
csv_file = 'data.csv'
latex_path = 'C:/Users/ryan.behdad/AppData/Local/Programs/MiKTeX 2.9/miktex/bin/x64/'

####################### Main #######################
df = pd.read_csv(csv_file)
# Cash_Flow[‘DCF’] = Cash_Flow[‘CF’]/(1.1**Cash_Flow[‘Months_Passed’])
# Calculate Cash_Flow[‘DCF’].sum() and name it SDCF
df['DCF'] = df.apply(lambda row: row['CF'] / (1.1**row['Months Passed']), axis = 1)
SDCF = df['DCF'].sum()
if SDCF>0:
    sentence = f'The project has a profit of ${SDCF:,.0f}'
else:
    sentence = f'The project has a loss of ${-1 * SDCF:,.0f}'

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
