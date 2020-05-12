# datawarrior_pdf_export
An alternative pdf export of an array of computed properties by DataWarrior

DataWarrior (http://www.openmolecules.org/datawarrior/index.html) is a freely
available program to compute and analyze data relevant to medicinal chemistry.
One of its internal representations of data is a data table, which may be print.
With an appropriate pdf printer, e.g. in Linux cups, it is possible to print
this table into a postscript file then converted (e.g., `pstopdf`) into a much
smaller .pdf print.

Some of the data stored in DataWarrior's own file format `.dwar` are readable
directly.  With the structures DataWarrior accessed available in a list of
SMILES (`.smi` file), it is possible to use information of both to create with
Python module [openpyxl](https://openpyxl.readthedocs.io/en/stable/) a `.xlsx`
file equally containing small `.png` about these molecules.  This may be edited
in LibreOffice Calc (as `.ods` file) and subsequently exported as `.pdf` which
is both smaller in size than the `.pdf` print from DataWarrior with cups (which
then basically was a container of pictures only), and additionally contains a
searchable text layer.

The script, deposit in the same folder as the `.dwar` and  `.smi` to work with
has requires `openpyxl` and `pil` as non-standard Python modules.  The molecular
structures are plot by [openbabel](www.openbabel.org).  It has shown to work
successfully both with current Python 3.8.3, as well as with the legacy of
Python 2.7.17 from the command line by

`python spreadsheet_test.py`
