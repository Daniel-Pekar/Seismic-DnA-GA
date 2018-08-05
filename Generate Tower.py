import os
import win32com.client
import openpyxl
import random
from openpyxl import *

#from geneticalgorithm.py###############################
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

####################################################

class Node:
    def __init__(self, name, x, y, z, x_gene_name = 'na', y_gene_name = 'na', z_gene_name = 'na'):
        self.name = name
        self.x = x
        self.y = y
        self.z = z
        self.x_gene_name = x_gene_name
        self.y_gene_name = y_gene_name
        self.z_gene_name = z_gene_name



        




#note: put in the row/column indices as strings


def get_nodes(nodes_name_col, nodes_x_col, nodes_y_col, nodes_z_col, nodes_start_row, range_to_mult_by_gene, gene_to_mult_by, chromosome, ws):

    def find_value_dep_on_gene(gene_name, range_to_mult_by_gene, gene_to_mult_by, chromosome):
        print(gene_name)
        if '-' in gene_name:
            gene_name_pos = gene_name.replace('-', '')
            for i in range(len(chromosome)):
                if chromosome[i].name == gene_name_pos:
                    value = chromosome[i].value * -1
                    if multiply_by_gene(gene_name, range_to_mult_by_gene):
                        value = value * chromosome[gene_to_mult_by - 1].value
                    break
        else:
            for i in range(len(chromosome)):
                if chromosome[i].name == gene_name:
                    value = chromosome[i].value
                    if multiply_by_gene(gene_name, range_to_mult_by_gene):
                        value = value * chromosome[gene_to_mult_by - 1].value
                    break
        return value

    def multiply_by_gene(var_to_check, range_to_mult):
        for current_gene_num in range_to_mult:
            if str(current_gene_num) in var_to_check:
                return True
        return False

    all_nodes = []
    #assign values to nodes
    current_row = nodes_start_row
    while ws[nodes_name_col + str(current_row)].value is not None:
        node_name = ws[nodes_name_col + str(current_row)].value
        node_x = ws[nodes_x_col + str(current_row)].value
        node_y = ws[nodes_y_col + str(current_row)].value
        node_z = ws[nodes_z_col + str(current_row)].value
        #check if the x or y locations are dependent on the value of a gene
        node_x_loc = 'na'
        node_y_loc = 'na'
        node_z_loc = 'na'
        if type(node_x) != int:
            node_x_loc = node_x
            node_x = find_value_dep_on_gene(node_x, range_to_mult_by_gene, gene_to_mult_by, chromosome)
        if type(node_y) != int:
            node_y_loc = node_y
            node_y = find_value_dep_on_gene(node_y, range_to_mult_by_gene, gene_to_mult_by, chromosome)
        if type(node_z) != int:
            node_z_loc = node_z
            node_z = find_value_dep_on_gene(node_z, range_to_mult_by_gene, gene_to_mult_by, chromosome)
        all_nodes.append(Node(node_name, node_x, node_y, node_z, node_x_loc, node_y_loc, node_z_loc))
        current_row = current_row + 1
    return all_nodes

def build_tower(nodes, start_node_col, end_node_col, member_prop_col, mat_props_cols, section_props_cols, start_row, ws):
    #create SAP2000 object
    SapObject = win32com.client.Dispatch('SAP2000v15.SapObject')
    #start SAP2000
    SapObject.ApplicationStart()
    #create SapModel Object
    SapModel = SapObject.SapModel
    #initiaize model
    SapModel.InitializeNewModel()
    #create new blank model
    ret = SapModel.File.NewBlank()
    #set units
    kN_m_C = 6
    SapModel.SetPresentUnits(kN_m_C)

    #define material properties
    for mat_prop_counter in range(len(mat_props_cols)):
        current_col = mat_props_cols[mat_prop_counter]
        mat_type = ws[current_col + str(start_row)].value
        mat_name = ws[current_col + str(start_row + 1)].value
        mat_E = ws[current_col + str(start_row + 2)].value
        mat_poisson = ws[current_col + str(start_row + 3)].value
        mat_thermal = ws[current_col + str(start_row + 4)].value
        mat_unit_wt = ws[current_col + str(start_row + 5)].value
        #create material type
        SapModel.PropMaterial.SetMaterial(mat_name, mat_type)
        #set isotropic material properties
        SapModel.PropMaterial.SetMPIsotropic(mat_name, mat_E, mat_poisson, mat_thermal)
        #set unit weight
        SapModel.PropMaterial.SetWeightAndMass(mat_name, 1, mat_unit_wt)

    #define section properties
    for sec_prop_counter in range(len(section_props_cols)):
        current_col = section_props_cols[sec_prop_counter]
        sec_name = ws[current_col + str(start_row)].value
        sec_mat = ws[current_col + str(start_row + 1)].value
        sec_shape = ws[current_col + str(start_row + 2)].value
        if 'SQUARE' == sec_shape:
            sec_width = ws[current_col + str(start_row + 3)].value
            sec_height = ws[current_col + str(start_row + 4)].value
            SapModel.PropFrame.SetRectangle(sec_name, sec_mat, sec_height, sec_width)
        elif 'CIRCULAR' == sec_shape:
            sec_dia = ws[current_col +str(start_row + 3)].value
            SapModel.PropFrame.SetCircle(sec_name, sec_mat, sec_dia)
        else:
            print('ERROR: Define a section shape in the setup workbook!')
    #create nodes
    for current_node_counter in range(len(nodes)):
        current_node_name = nodes[current_node_counter].name
        current_node_x = nodes[current_node_counter].x
        current_node_y = nodes[current_node_counter].y
        current_node_z = nodes[current_node_counter].z
        print('Creating node at ' + str(current_node_x) + ', ' + str(current_node_y) + ', ' + str(current_node_z))
        assigned_node_name = ' '
        [ret, assigned_node_name] = SapModel.PointObj.AddCartesian(current_node_x, current_node_y, current_node_z, assigned_node_name, current_node_name)
        print(assigned_node_name)
        print(type(assigned_node_name))
        if ret != 0:
            print('Error creating node')
    print('Done creating nodes')
    #create members
    current_row = start_row
    while ws[start_node_col + str(current_row)].value is not None:
        print(current_row)
        start_node_name = ws[start_node_col + str(current_row)].value
        end_node_name = ws[end_node_col + str(current_row)].value
        member_prop = ws[member_prop_col + str(current_row)].value
        member_name = start_node_name + ' to ' + end_node_name
        print('Creating member ' + member_name)
        assigned_member_name = ' '
        ret = SapModel.FrameObj.AddByPoint(start_node_name, end_node_name, assigned_member_name, member_prop, member_name)
        print(ret)
        print(assigned_member_name)
        if ret != 0:
            print('Error creating member ' + member_name)
        current_row = current_row + 1
    print('Done creating members')
    #return the SapModel object


wb = load_workbook('Setup.xlsx')
ws = wb.active

TestChromosome = Chromosome(len=23)
TestChromosome.create_Chromosome(ws)

AllNodes = get_nodes('B', 'F', 'G', 'H', 4, range(2,12), 1, TestChromosome.genes, ws)
build_tower(AllNodes, 'Q', 'R', 'T', ['X', 'AA'], ['AD', 'AG'], 4, ws)
for i in range(len(AllNodes)):
    print(AllNodes[i].name)
    print(AllNodes[i].x)
    print(AllNodes[i].y)
    print(AllNodes[i].z)