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

class Chromosome:
    def __init__(self,len,genes = [],fitness = 0, chance = 0):
        self.len = len
        self.genes = genes
        self.fitness = fitness
        self.selectionchance = chance

    def create_Chromosome(self,worksheet):
        self.genes = [Gene(ws = worksheet, cell_col = i + 4) for i in range(self.len)]
        for gene in self.genes:
            gene.generate_initial_values(ws)

    #Selection, crossover, mutation

class Population:
    def __init__(self, chromes = [],generation = 1,pop = 0,chromelen = 0):
        self.chromes = chromes
        self.generation = generation
        self.pop = pop
        self.chromelen = chromelen

    def create_initial_pop(self,ws):
        self.chromes = [Chromosome(len = self.chromelen) for i in range(self.pop)]
        for chrome in self.chromes:
            chrome.create_Chromosome(ws)

    def avg_fitness(self):
        total_fitness = 0
        for chrome in self.chromes:
            total_fitness += chrome.fitness
        return total_fitness/self.pop

    def max_fitness(self):
        maxf = 0
        for chrome in self.chromes:
            if chrome.fitness > maxf:
                maxf = chrome.fitness
        return maxf

    #def elitism(self,num):

    def fitprop_roulette(self):
        totalfit = 0
        for chrome in self.chromes:
            totalfit += chrome.fitness
        NewPop = Population(generation = self.generation + 1, chromelen = self.chromelen)


    #def stochastic_universal

    #def tournament

    #def rank selection

    #def crossover(self):

    #def mutation(self):

wb = load_workbook('Setup.xlsx')
ws = wb.active
clen = (ws['J2']).value

Popinit = Population(chromelen = clen,pop = 5)
Popinit.create_initial_pop(ws)

for i in Popinit.chromes:
    for j in i.genes:
        print(j.value)


