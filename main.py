import pylatex as pl
import pandas as pd
import subprocess

####################### Arguments #######################
file_name = 'report'
csv_file = 'data.csv'
report_title = 'Economic Analysis'
repot_author = 'Ryan Behdad'
latex_path = 'C:/Users/ryan.behdad/AppData/Local/Programs/MiKTeX 2.9/miktex/bin/x64/'

####################### functions #######################
def create_pdf(filename):
    input_filename = file_name + '.tex'
    output_filename = file_name
    process = subprocess.Popen([
        latex_path + 'latex',
        '-output-format=pdf',
        '-job-name=' + output_filename,
        input_filename])
    process.wait()

####################### Main #######################
df_original = pd.read_csv(csv_file)
df = df_original.copy()
df['DCF'] = df.apply(lambda row: row['CF'] / (1.1**row['Months Passed']), axis = 1)
SDCF = df['DCF'].sum()
if SDCF>0:
    sentence = f'The project has a profit of ${SDCF:,.0f}'
else:
    sentence = f'The project has a loss of ${-1 * SDCF:,.0f}'

# Create document
doc = pl.Document()
doc.packages.append(pl.Package('booktabs'))

# Add preamble
doc.preamble.append(pl.Command('title', report_title))
doc.preamble.append(pl.Command('author', repot_author))
doc.preamble.append(pl.Command('date', pl.NoEscape(r'\today')))
doc.append(pl.NoEscape(r'\maketitle'))

# Create section 1
with doc.create(pl.Section('Data')):
    doc.append('The cash flow data is as follows.')
    with doc.create(pl.Table(position='htbp')) as table:
        table.add_caption('Cash Flow')
        table.append(pl.Command('centering'))
        table.append(pl.NoEscape(df_original.to_latex(escape=False, index=False)))

# Create section 2
with doc.create(pl.Section('Conclusion')):
    doc.append(sentence)

# Create section 3
with doc.create(pl.Section('Appendix')):
    doc.append('The discounted cash flow is shown in the table below.')
    with doc.create(pl.Table(position='htbp')) as table:
        table.add_caption('Discounted Cash Flow')
        table.append(pl.Command('centering'))
        table.append(pl.NoEscape(df.to_latex(escape=False, index=False)))

doc.generate_tex(file_name)
create_pdf(file_name)
