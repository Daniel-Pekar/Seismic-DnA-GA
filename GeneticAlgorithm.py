<<<<<<< HEAD
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


=======
import os
import win32com.client
import openpyxl
import random
from openpyxl import *
import bisect
from operator import attrgetter

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
    def __init__(self,len,genes = [],fitness = 0):
        self.len = len
        self.genes = genes
        self.fitness = random.random() #lol fix

    def create_initial_chromosome(self,worksheet):
        self.genes = [Gene(ws = worksheet, cell_col = i + 4) for i in range(self.len)]
        for gene in self.genes:
            gene.generate_initial_values(ws)

    def mutation_uniform(self,prob):
        for i in range(self.len):
            if random.random()<prob:
                self.genes[i].value = random.uniform(self.genes[i].lower,self.genes[i].upper)

    def mutation_triangular(self,prob):
        for i in range(self.len):
            if random.random() < prob:
                self.genes[i].value = random.triangular(self.genes[i].lower, self.genes[i].upper)

    def mutation_min(self,prob):
        for i in range(self.len):
            if random.random() < prob:
                self.genes[i].value = random.uniform(self.genes[i].lower)

    def mutation_max(self,prob):
        for i in range(self.len):
            if random.random() < prob:
                self.genes[i].value = random.uniform(self.genes[i].upper)

class Population:
    def __init__(self, chromosomes = [],generation = 1,pop = 0,chromosomelen = 0):
        self.chromosomes = chromosomes
        self.generation = generation
        self.pop = pop
        self.chromlen = chromosomelen

    def create_initial_pop(self,ws):
        self.chromosomes = [Chromosome(len = self.chromlen) for i in range(self.pop)]
        for chromosome in self.chromosomes:
            chromosome.create_initial_chromosome(ws)

    def total_fit(self):
        return sum(self.chromosomes[i].fitness for i in range(self.pop))

    def avg_fitness(self):
        return self.total_fit()/self.pop

    def max_fitness(self):
        maxf = 0
        for chromosome in self.chromosomes:
            if chromosome.fitness > maxf:
                maxf = chromosome.fitness
        return maxf

    def selection_elitism(self,num):
        self.chromosomes.sort(key=lambda x: x.fitness, reverse=True)
        parents = []
        for i in range(num):
            parents.append(self.chromosomes.pop(0))
        self.pop -= num
        return parents

    def selection_roulette(self,num_parents):
        self.chromosomes.sort(key=lambda x: x.fitness, reverse=False)
        parents = []
        fitness_list = []
        for chromosome in self.chromosomes:
            fitness_list.append(chromosome.fitness)
        parents = random.choices(self.chromosomes,weights = fitness_list, k = num_parents)
        return parents

    def selection_stochastic(self,num_parents):
        self.chromosomes.sort(key=lambda x: x.fitness, reverse=True)
        parents = []
        cumfit = 0
        fitness_list = [[], self.chromosomes]
        for chromosome in self.chromosomes:
            cumfit += chromosome.fitness
            fitness_list[0].append(cumfit)
        randnum = random.uniform(0, cumfit/num_parents)
        for i in range(num_parents):
            ind = bisect.bisect_left(fitness_list[0], randnum+(cumfit/num_parents)*(i))
            parents.append(fitness_list[1][ind])
        return parents

    def selection_tournament(self,num_parents,num_fighters):
        parents = []
        for i in range(num_parents):
            fighters = []
            for j in range(num_fighters):
                ind = random.randint(0,self.pop)
                fighters.append(self.chromosomes[ind])
                self.chromosomes.pop(ind)
            parents.append(max(fighters, key=attrgetter('fitness')))
        return parents

    def selection_rank(self,num_parents):
        self.chromosomes.sort(key=lambda x: x.fitness, reverse=True)
        weightings = []
        for i in range(self.pop):
            weightings.append(1/(i+1)) #1/(x+1) weightings for rank x
        parents = random.choices(self.chromosomes, weights = weightings, k = num_parents)
        return parents

    def crossover_npoint(self,n,parents):
        crossover_points = random.sample(range(1, self.chromlen), n)
        crossover_points.sort(key=lambda x: x, reverse=False)
        crossover_points.insert(0, 0)
        crossover_points.append(self.chromlen)
        children = [[], []]
        origin = []
        for i in range(len(crossover_points) - 1):
            for j in range(crossover_points[i + 1] - crossover_points[i]):
                origin.append(i % 2)
        for i in range(len(origin)):
            children[0].append(parents[origin[i]][i])
            children[1].append(parents[1 - origin[i]][i])
        return children

    def crossover_randomflip(self,parents,num_children):
        children = []
        for i in range(num_children):
            children.append([])
            for j in range(self.chromlen):
                if random.randint(0,1) == 0:
                    children[-1].append(parents[0][j])
                else:
                    children[-1].append(parents[1][j])
        return children

    def crossover_flip(self,parents):
        children = [[],[]]
        for i in range(self.chromlen):
            children[0].append(parents[i%2][i])
            children[1].append(parents[1 - i%2][i])
        return children

    def crossover_triangle(self,parents,num_children):
        children = [[] for i in range(num_children)]
        for i in children:
            for j in range(self.chromlen):
                i.append(random.triangular(parents[0][j],parents[1][j]))
        return children

    def crossover_uniform(self,parents,num_children):
        children = [[] for i in range(num_children)]
        for i in children:
            for j in range(self.chromlen):
                i.append(random.uniform(parents[0][j],parents[1][j]))
        return children
>>>>>>> master
