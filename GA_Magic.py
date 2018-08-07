import GeneticAlgorithm
import Generate_Tower
from GeneticAlgorithm import *
from Generate_Tower import *

wb = load_workbook('Setup.xlsx')
ws = wb.active
glen = (ws['J2']).value
# pop = (ws['']).value
#chromlen = (ws['']).value
generations = 10
max_fit = []
avg_fit = []

Pop = Population(chromosomelen = glen,pop = 10)
Pop.create_initial_pop(ws)


def run_GA(generations,population,chromosomelen,num_elitism,type_selection,type_crossover,type_mutation,mutation_prob):

for i in range(generations):
    for chromosome in Pop.chromosomes:
        chromosome.run_sap_models()
    avg_fit.append(Pop.avg_fitness())
    max_fit.append(Pop.max_fitness())
    new_Pop = Population(chromosomelen = glen, pop = 10)
    temp_parent = Pop.selection_elitism(2)
    temp_parent += Pop.selection_roulette(2)
    temp_parents = temp_parent+temp_parent_2
    parents = []
    for j in temp_parents:
        parents.append(j)
    #use choices or something to get 5 pairs out of 4 parents
    #apply same crossover for all pairs of parents to get children
    #put all children into the population
    #for all choromosomes in population, mutate
    #create sap models
    #eval max and avg fitness
    #apply elitism for 2 best
    #apply roulette/stochastic/rank/tournament selection
    #choose 2 parents
    #make 5 pairs of children using crossovers (npoint/randomflip/flip/uniform/triangular)
    #mutate
