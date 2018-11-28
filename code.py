import pandas as pd
import pulp
import time
import cplex
from xlsxwriter.utility import xl_rowcol_to_cell

supplier = pd.read_excel('C:\\Users\Swadha\\Desktop\\MHM\\Model\\supplier.xlsx', sheet_name="Sheet1")
warehouse = pd.read_excel('C:\\Users\Swadha\\Desktop\\MHM\\Model\\warehouse.xlsx', sheet_name="Sheet1")
storage = pd.read_excel('C:\\Users\Swadha\\Desktop\\MHM\\Model\\storage.xlsx', sheet_name="Sheet1")
TransCostij = pd.read_excel('C:\\Users\\Swadha\\Desktop\\MHM\\Model\\SupplyCostij.xlsx', index_col=[0,1],sheet_name="Sheet1")
TransCostjk = pd.read_excel('C:\\Users\\Swadha\\Desktop\\MHM\\Model\\SupplyCostjk.xlsx', index_col=[0,1],sheet_name="Sheet1")
SupplyCap= pd.read_excel('C:\\Users\\Swadha\\Desktop\\MHM\\Model\\capacity.xlsx', index_col=[0], sheet_name="Sheet1")
dk =  pd.read_excel('C:\\Users\\Swadha\\Desktop\\MHM\\Model\\dk.xlsx',sheet_name="Sheet1" )

ts = time.gmtime()
print(time.strftime("%Y-%m-%d %H:%M:%S", ts))

FacilityCost = 200000
ProcCost = 7.5 #PcC_ij : Cost of procurement from supplier i for potential warehouse j
QualCost = 0.1* ProcCost #tcoq_yi : Total cost of quality for supplier i as a function of percentage defectives by supplier
MinPctDemand = 0.1
#arb = pd.read_excel('/media/swadha/Windows/Users/Swadha/Desktop/MHM/Model/arb.xlsx', index_col=[0])
zmin = 3
zmax = 35
defective = 0.1
sumdk = 17903192.30
M = 10000000

"""Xij = pulp.LpVariable.dicts("Xij", ((i, j) for i in supplier for j in warehouse),lowBound=0, cat='Continuous') #S_ij
netXij = pulp.LpVariable.dicts("netXij",((i, j) for i in supplier for j in warehouse),lowBound=0, cat='Continuous') #S_ij*
zj =  pulp.LpVariable.dicts("Zj", ((j) for j in warehouse), lowBound=0, upBound=1,  cat='Continuous')#binary variable
Xjk = pulp.LpVariable.dicts("Xjk", ((j, k) for j in warehouse for k in storage),lowBound=0, cat='Continuous') #Xjk """

Xij = pulp.LpVariable.dicts("Xij", ((i, j) for i in supplier for j in warehouse),lowBound=0, cat='Integer') #S_ij
netXij = pulp.LpVariable.dicts("netXij",((i, j) for i in supplier for j in warehouse),lowBound=0, cat='Integer') #S_ij*
zj =  pulp.LpVariable.dicts("Zj", ((j) for j in warehouse), lowBound=0, upBound=1,  cat='Integer')#binary variable
Xjk = pulp.LpVariable.dicts("Xjk", ((j, k) for j in warehouse for k in storage),lowBound=0, cat='Integer') #Xjk

model = pulp.LpProblem("Cost minimising problem", pulp.LpMinimize)

model += pulp.lpSum([Xij[(i,j)]*ProcCost for i in supplier for j in warehouse] +
                    [QualCost*(1-defective)*SupplyCap.ix[i,0] for i in SupplyCap.index]+
                    [FacilityCost * zj[j] for j in warehouse]+
                    [TransCostij.ix[(i,j),0]*(1/23500)*Xij[i,j] for i in supplier for j in warehouse]+
                    [TransCostij.ix[(i,j),0]*(1/23500)*(Xij[i,j]-netXij[i,j]) for i in supplier for j in warehouse]+
                    [TransCostjk.ix[(j,k),0]*(1/23500)*Xjk[j,k] for j in warehouse for k in storage]
                    )

#Supply constraint
for i in supplier:
    model += pulp.lpSum(Xij[(i,j)] for j in warehouse) <= SupplyCap.ix[i, 0]

#Demand Constraint
for k in storage:
    if (dk.loc[k,'Storage']==k):
        model += pulp.lpSum([Xjk[(j,k)] for j in warehouse]) >= dk.loc[k,'Demandk']

#Facility opening : linking constraint
for j in warehouse:
    for i in supplier:
        for k in storage:
            model += netXij[(i,j)]+Xjk[(j,k)]-zj[j]*M <= 0

#Flow conservation:
for j in warehouse:
    model += pulp.lpSum([Xjk[(j,k)] for k in storage]) == pulp.lpSum([netXij[(i,j)] for i in supplier])

#Level of Service constraint
model += pulp.lpSum(Xjk[j,k]*TransCostjk.ix[(j,k), 1] for j in warehouse for k in storage) >= MinPctDemand*sumdk
model += pulp.lpSum(Xjk[j,k]*TransCostjk.ix[(j,k), 2] for j in warehouse for k in storage) >= MinPctDemand*sumdk
model += pulp.lpSum(Xjk[j,k]*TransCostjk.ix[(j,k), 3] for j in warehouse for k in storage) >= MinPctDemand*sumdk

#Quantity - Quality linking constraint
for i in supplier:
    for j in warehouse:
        model += Xij[(i,j)]*(1-defective) >= netXij[(i,j)]

#Number of facilities
model += pulp.lpSum([zj[j] for j in warehouse]) >= zmin
model += pulp.lpSum([zj[j] for j in warehouse]) <= zmax

solver = pulp.CPLEX()
model.setSolver(solver)
model.solve()

print(pulp.LpStatus[model.status])
print(pulp.value(model.objective))

count1=0
count2=0

for j in warehouse:
    if (zj[j].varValue == 1):
            print pulp.lpSum(Xjk[j,k].varValue for k in storage)
            if(pulp.lpSum(Xjk[j,k].varValue for k in storage)>=1000000):
                FacilityCost = 300000
                count1 = count1+1
            else:
                FacilityCost = 100000
                count2 = count2+1

print count1
print count2
print pulp.lpSum(zj[j].varValue for j in warehouse)

NewCost = pulp.value(model.objective) - 200000*pulp.lpSum(zj[j].varValue for j in warehouse)+pulp.lpSum(zj[j].varValue*FacilityCost for j in warehouse)
print NewCost

ts = time.gmtime()
print(time.strftime("%Y-%m-%d %H:%M:%S", ts))

output = []
for j in warehouse:
    var_output={
    'Warehouse' : j,
    'Zj': zj[j].varValue
    }
    if (zj[j].varValue > 0):
        output.append(var_output)
output_df = pd.DataFrame.from_records(output).sort_values(['Warehouse', 'Zj'])
output_df.set_index(['Warehouse','Zj'], inplace=True)
print(output_df)

writer_orig = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')
output_df.to_excel(writer_orig, index=True, sheet_name='Zj')
writer_orig.save()


output = []
for i in supplier:
    for j in warehouse:
        var_output={
        'supplier' : i,
        'warehouse' : j,
        'Xij': Xij[i, j].varValue
        }
        if (Xij[i,j].varValue > 0):
            output.append(var_output)
output_df = pd.DataFrame.from_records(output).sort_values(['supplier','warehouse', 'Xij'])
output_df.set_index(['supplier', 'warehouse','Xij'], inplace=True)
print(output_df)

writer_orig = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')
output_df.to_excel(writer_orig, index=True, sheet_name='Xij')
writer_orig.save()


output = []
for j in warehouse:
    for k in storage:
        var_output={
        'warehouse' : j,
        'storage' : k,
        'Xjk': Xjk[j, k].varValue
        }
        if (Xjk[j,k].varValue > 0):
            output.append(var_output)
output_df = pd.DataFrame.from_records(output).sort_values(['warehouse','storage', 'Xjk'])
output_df.set_index(['warehouse', 'storage','Xjk'], inplace=True)
print(output_df)

writer_orig = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')
output_df.to_excel(writer_orig, index=True, sheet_name='Xjk')
writer_orig.save()
