import gurobipy as grb
import numpy as np
import xlrd 
import math
import csv

# Reading number of covid cases per state
wb = xlrd.open_workbook("covid.xlsx") 
sheet = wb.sheet_by_index(0) 

State_covid = {}
for i in range(1,sheet.nrows):
    State_covid[sheet.cell_value(i, 0)] = sheet.cell_value(i, 1)

# Read population per city
wb = xlrd.open_workbook("Census.xlsx") 

# Calculating distance between two cities
def haversine_distance(lat1, lon1, lat2, lon2):
    a = (lat1-lat2)*(lat1-lat2)
    b = (lon1-lon2)*(lon1-lon2)
    distance = math.sqrt(a+b)
    return distance

sheet = wb.sheet_by_index(0) 

Cities = []
Cities_Name = {}
Lat = {}
Long = {}
State = {}
Demand = {}
Cost = {}
Facilities = []
Fixed_Cost = {}
Capacity = {}

print("Importing Data")

for i in range(1,250):
    zip_code = int(sheet.cell_value(i, 0))
    Cities.append(zip_code)
    Facilities.append(zip_code)
    Fixed_Cost[zip_code] = 1000
    Capacity[zip_code] = 10000000
    Cities_Name[zip_code] = sheet.cell_value(i, 3)
    Lat[zip_code] = float(sheet.cell_value(i, 1))
    Long[zip_code] = float(sheet.cell_value(i, 2))
    State[zip_code] = sheet.cell_value(i, 4)
    Demand[zip_code] = int(sheet.cell_value(i, 5))

print("Number of cities ", len(Cities))
print("Number of possible hubs ", len(Facilities))

print("Finish Importing Data")
print("Calculating distances")

Distances = {}

# Calculating distances for every pair of cities
for i in Cities:
    for j in Cities:
        dist = haversine_distance(Lat[i], Long[i], Lat[j], Long[j])
        Distances[(i,j)] = dist
        Cost[(i,j)] = dist

print("Finish Calculating distances")

# Calling the optimization Model
opt_model = grb.Model()

# Variables: flow between two cities
x  = {(i,j): opt_model.addVar(vtype=grb.GRB.CONTINUOUS, 
                        lb=0,
                        name="x_{0}_{1}".format(i,j)) 
for i in Cities for j in Facilities}

# Binary variable: 1 if it's used 0 otherwise
y  = {(j): opt_model.addVar(vtype=grb.GRB.BINARY, 
                        name="y_{0}".format(j)) 
for j in Facilities}

# All demand should be satisfied
Cons1 = {i : 
opt_model.addConstr(
        lhs=grb.quicksum(x[i,j] for j in Facilities),
        sense=grb.GRB.EQUAL,
        rhs=Demand[i], 
        name="Cons1_{0}".format(i))
    for i in Cities}

# Capacity Constraint
Cons2 = {j : 
opt_model.addConstr(
        lhs=grb.quicksum(x[i,j] for i in Cities),
        sense=grb.GRB.LESS_EQUAL,
        rhs=Capacity[j]*y[j], 
        name="Cons2_{0}".format(j))
	for j in Facilities}

# Link between x and y variable
Cons3 = {(i,j) : 
opt_model.addConstr(
        lhs=x[i,j],
        sense=grb.GRB.LESS_EQUAL,
        rhs=Demand[i]*y[j], 
        name="Cons3_{0}_{1}".format(i,j))
	for i in Cities for j in Facilities}

# Number of hub points built
Cons4 = opt_model.addConstr(
        lhs=grb.quicksum(y[j] for j in Facilities),
        sense=grb.GRB.EQUAL,
        rhs=8, 
        name="Cons4")



objective = grb.quicksum(Fixed_Cost[j]*y[j] for j in Facilities) + grb.quicksum(Cost[i,j]*x[i,j] for i in Cities for j in Facilities)
opt_model.ModelSense = grb.GRB.MINIMIZE
opt_model.setObjective(objective)

# Run model
opt_model.optimize()

f = open('used.csv', 'w', newline = '')
with f:
    fnames = ['Hub', 'Used']
    writer = csv.DictWriter(f, fieldnames=fnames)    
    writer.writeheader()
    Used_Facilities = []
    for j in Facilities:
        writer.writerow({'Hub' : j, 'Used': int(y[j].x)})
        if(y[j].x > 0.1):
            Used_Facilities.append(int(y[j].x))

print("Number of used hubs ", len(Used_Facilities))
f = open('flow.csv', 'w', newline = '')
with f:
    fnames = ['Origin', 'Destination', 'Flow']
    writer = csv.DictWriter(f, fieldnames=fnames)    
    writer.writeheader()

    Flow = {}
    for i in Cities:
        for j in Facilities:
            if(x[i,j].x > 0.1):
                writer.writerow({'Origin' : j, 'Destination' : i, 'Flow': int(x[i,j].x)})
                Flow[i,j]=x[i,j].x

