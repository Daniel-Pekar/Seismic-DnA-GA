import os
import win32com.client
import openpyxl
import random
from openpyxl import *

class Gene:
    def __init__(self,ws,name="",lower=0,upper=0,value=0,cell_col = 0):
        self.cell_col = cell_col
        self.name = ws['K' + str(self.cell_col)].value
        self.lower = ws['L' + str(self.cell_col)].value
        self.upper = ws['M' + str(self.cell_col)].value
        self.value = value

    def generate_initial_values(self,ws):
        self.value = random.uniform(self.lower,self.upper)

class Genome:
    def __init__(self,len,genes = [],fitness = 0):
        self.len = len
        self.genes = genes
        self.fitness = fitness

    def create_Genome(self,worksheet):
        self.genes = [Gene(ws = worksheet, cell_col = i + 4) for i in range(self.len)]
        for gene in self.genes:
            gene.generate_initial_values(ws)

    #Selection, crossover, mutation

class Population:
    def __init__(self, genomes = [],generation = 1,pop = 0,genomelen = 0):
        self.genomes = genomes
        self.generation = generation
        self.pop = pop
        self.genomelen = genomelen

    def create_initial_pop(self,ws):
        self.genomes = [Genome(len = self.genomelen) for i in range(self.pop)]
        for genome in self.genomes:
            genome.create_Genome(ws)

    def avg_fitness(self):
        total_fitness = 0
        for genome in self.genomes:
            total_fitness += genome.fitness
        return total_fitness/self.pop

    def max_fitness(self):
        maxf = 0
        for genome in self.genomes:
            if genome.fitness > maxf:
                maxf = genome.fitness
        return maxf

wb = load_workbook('Setup.xlsx')
ws = wb.active
glen = (ws['J2']).value

Popinit = Population(genomelen = glen,pop = 5)
Popinit.create_initial_pop(ws)
