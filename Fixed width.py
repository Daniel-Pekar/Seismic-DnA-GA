import os
import win32com.client
import openpyxl
from openpyxl import Workbook
######
#Create variable class


class Variable:
    Dependent = False
    LowerBound = 0
    UpperBound = 1
    Value = 3


wb = Workbook()
ws = wb.active

Genome = []

NumberOfVariables = 23
i = 1
while i <= NumberOfVariables:
    Genome.append(Variable())
    i = i + 1

#Assign constraints to the variables that don't depend on other variables
#ASSUMES columns go: variable name, dependency, variable to subtract, number to subtract from, less than, greater than
i = 1
VariablesCol = 10
VariablesRow = 5
while i <= NumberOfVariables:
    if ws.cell(row = VariablesRow, col = VariablesCol + 1) = 0
        Genome(i).LowerBound = VariablesCol + 4
        Genome(i).UpperBound = VariablesCol + 5
        Genome(i).Value =
    else
        Genome(i).Value = 'WAIT'

    i = i+1
    VariablesRow = VariablesRow + 1
