from GeneticAlgorithm import *
from Generate_Tower import *
import matplotlib.pyplot as plt

def create_new_population(old_population,excel_index):
    TimeHistory = r'C:\Users\kotab\Documents\Seismic\EQ1_acc.txt'
    CurrentIndex = 1
    for CurChromosome in oldPop.chromosomes:
        SaveLocation = r'C:\Users\kotab\Documents\Seismic\Models\SAP2000_model' + str(CurrentIndex) + '.sdb'
        StartTime = time.time()
        print('\n\nBuilding ' + str(CurrentIndex) + ' out of ' + str(PopulationSize) + ' in population...')
        SapObject = ga_CONSTRUCT(CurChromosome.genes, ws, ExcelIndex, TimeHistory, SaveLocation)
        max_acc = ga_ANALYZE(SapObject)
        TotalTime = time.time() - StartTime
        print('Time to run', TotalTime)
        CurChromosome.fitness = max_acc
        CurrentIndex = CurrentIndex + 1
    new_population = Population()
    new_population.generation = old_population.generation + 1
    new_population.pop = old_population.pop
    new_population.chromlen = old_population.chromlen
    selection_type = excel_index['Type of Selection']
    crossover_type = excel_index['Type of Crossover']
    population_size = excel_index['Population']
    num_parents = excel_index['Number Parents']
    mut_rate = excel_index['Mutation Rate']
    mutation_type = excel_index['Type of Mutation']
    parents = []
    children = []
    children.extend(old_population.selection_elitism(ExcelIndex['Elitism Number']))
    if selection_type == "Roulette":
        parents.extend(old_population.selection_roulette(num_parents))
    elif selection_type == "Stochastic":
        parents.extend(old_population.selection_stochastic(num_parents))
    elif selection_type == "Tournament":
        parents.extend(old_population.selection_tournament(num_parents,ExcelIndex['Number Fighters']))
    elif selection_type == "Rank":
        parents.extend(old_population.selection_rank(num_parents))
    for i in range((population_size-ExcelIndex['Elitism Number'])//2):
        temp_parents = random.sample(parents,2)
        if crossover_type == "Npoint":
            children.extend(old_population.crossover_npoint(ExcelIndex['Number Points'],temp_parents))
        elif crossover_type == "Randomflip":
            children.extend(old_population.crossover_randomflip(temp_parents,2))
        elif crossover_type == "Flip":
            children.extend(old_population.crossover_flip(temp_parents))
        elif crossover_type == "Triangle":
            children.extend(old_population.crossover_triangle(temp_parents,2))
        elif crossover_type == "Uniform":
            children.extend(old_population.crossover_uniform(temp_parents,2))
    for child in children:
        if mutation_type == "Uniform":
            child.mutation_uniform(mut_rate)
        elif mutation_type == "Triangular":
            child.mutation_triangular(mut_rate)
        elif mutation_type == "Min":
            child.mutation_min(mut_rate)
        elif mutation_type == "Max":
            child.mutation_max(mut_rate)
    new_population.chromosomes = children
    return new_population

wb = load_workbook('Setup.xlsx')
ws = wb.active


ExcelIndex = Generate_Tower.get_excel_indices(ws, 'A', 'B', 4)
NumGenerations = ExcelIndex['Generations']
PopulationSize = ExcelIndex['Population']
ChromosomeLen = ExcelIndex['Chromosome Length']
NumElitism = ExcelIndex['Elitism Number']
SelectionType = ExcelIndex['Type of Selection']
CrossoverType = ExcelIndex['Type of Crossover']
MutationType = ExcelIndex['Type of Mutation']
MutationRate = ExcelIndex['Mutation Rate']
NumberFighters = ExcelIndex['Number Fighters'] #For tournament selection
NumberPoints = ExcelIndex['Number Points'] #For N point crossover

#Create initial population
max_fit = []
avg_fit = []

oldPop = Population(chromosomelen = ChromosomeLen, pop = PopulationSize)
oldPop.create_initial_pop(ws)

for i in range(NumGenerations):
    newPop = create_new_population(oldPop,ExcelIndex)
    max_fit.append(newPop.max_fitness())
    avg_fit.append(newPop.avg_fitness())
    oldPop = newPop
    print(max_fit)
    print(avg_fit)

x = [i for i in range(len(max_fit))]

plt.plot(x,max_fit,'r--',x,avg_fit,'b--')
plt.show()


#

'''
def run_GA(generations,population,chromosomelen,num_elitism,type_selection,type_crossover,type_mutation,mutation_prob):


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
'''