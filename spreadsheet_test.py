# name:    spreadsheet_test.py
# author:  nbehrnd@yahoo.com
# license: MIT
# date:    2020-05-12 (YYYY-MM-DD)
# edit:
""" Write a .xlsx for .pdf export for corresponding .smi and .dwar file

A direct print of DataWarrior's table into a .pdf with cups yiels a
file larger in size, than anticipated.  As all content is represented
as an image, there is no searchable text layer either.

Started by

python spreadsheet_test.py

the script asks both for DataWarrior's .dwar file and a list of the SMILES
to instruct openbabel to plot the structures eventually put into the .xlsx
file.  Worked fine in both Python 3.8.3 and legacy branch 2.7.17.

"""
import os
import subprocess as sub
import sys

# non-standard Python modules, installed earlier via pip
from openpyxl.drawing.image import Image as XLIMG
# from openpyxl.worksheet import Worksheet
from openpyxl import Workbook

# initate spread sheet:
WB = Workbook()
WS = WB.active
# to accomodate _better_ the initial deposit of openbabels 300 x 300 px .png:
WS.column_dimensions['A'].width = 60


def read_dwar_lines():
    """ Retrieve relevant lines of information in DW's .dwar file. """
    if sys.version_info[0] == 2:
        dwar_source = str(raw_input("DataWarrior file to consider: "))
    if sys.version_info[0] == 3:
        dwar_source = str(input("DataWarrior file to consider: "))

    print("Considered input: {}".format(dwar_source))

    # identify lines with content of interest:
    read = False
    pre_register = []
    with open(dwar_source, mode="r") as source:
        for line in source:
            if line.startswith("idcoordinates2D"):
                read = True
            if line.startswith("<datawarrior properties>"):
                read = False
                break

            if read:
                pre_register.append(str(line).strip())

    del pre_register[0]  # the table caption will be restored later
    return pre_register


def read_dwar_information(read_dwar_lines):
    """ Transfer of .dwar columns past idcoordinates / FragFp """
    # restore the header of the spread sheet:
    WS['A1'] = "structure"

    WS['B1'] = "mol_name"
    WS['C1'] = "drug_like"
    WS['D1'] = "mutagenic"
    WS['E1'] = "tumorigenic"

    WS['F1'] = "reproductive_effect"
    WS['G1'] = "irritant"

    # transfer, reading:
    counter = 2  # because line #1 is already used by the heading
    for line in read_dwar_lines():
        entry_line = str(line).strip()

        mol_name = entry_line.split()[2]
        druglike = entry_line.split()[3]
        mutagenic = entry_line.split()[4]
        tumorigenic = entry_line.split()[5]

        reproductive_effect = entry_line.split()[6]
        irritant = entry_line.split()[7]

        # transfer, writing:
        WS.row_dimensions[counter].height = 100

        WS['B{}'.format(counter)] = mol_name
        WS['C{}'.format(counter)] = druglike
        WS['D{}'.format(counter)] = mutagenic
        WS['E{}'.format(counter)] = tumorigenic

        WS['F{}'.format(counter)] = reproductive_effect
        WS['G{}'.format(counter)] = irritant

        counter += 1


def openbabel_drawing():
    """ Ask openbabel to translate the SMILES into .png """
    if sys.version_info[0] == 2:
        smiles_input = str(raw_input("\nFile with the SMILES strings: "))
    if sys.version_info[0] == 3:
        smiles_input = str(input("\nFile with the SMILES strings: "))

    print("SMILES strings are read from: {}".format(smiles_input))

    with open(smiles_input, mode="r") as source2:
        counter = 2

        for line in source2:
            # work with openbabel:
            obabel_input = line.strip()  #.split()[0]
            obabel_output = ''.join(['entry', str(counter), '.png'])

            draw = str("obabel -:'{}' -opng -O {}".format(
                obabel_input, obabel_output))
            sub.call(draw, shell=True)

            # addition of the openbabel's .png images into the spreadsheet
            img = XLIMG(obabel_output)
            WS.add_image(img, 'A{}'.format(counter))
            counter += 1

    # eventually save the worksheet as a permanent record.
    WB.save("test.xlsx")


def space_cleaning():
    """ Remove the intermediate .png """
    for file in os.listdir("."):
        if file.endswith(".png"):
            os.remove(file)


def main():
    """ Joining function calls. """
    read_dwar_information(read_dwar_lines)
    openbabel_drawing()
    space_cleaning()


main()
