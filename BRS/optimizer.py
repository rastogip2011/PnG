import pandas as pd
import numpy as np
from pulp import *
from variables import *

def make_prob(n):
    prob = LpProblem('BRS',LpMinimize)

    x = []
    for i in range(n):
        a = 'x'+str(i)
        a = LpVariable(a, lowBound = 0, cat = 'Integer')
        x.append(a)
    
    return prob, x

print("\nStep 3:  Running optimizer")

Cost = pd.read_excel(filename_master, sheet_cost).values
Truck = pd.read_excel(filename_master, sheet_truck_cap)
t=Truck.columns[1:].tolist()
num_trucks = len(t)
truck=Truck.values
tvol = truck[1,1:]
tweight = truck[2,1:]
cost = dict()
site = dict()
for r in Cost:
    br = r[1].split(' + ')
    b=''
    for j in br:
        b+= j.split(' ')[0] + ' + '
    b=b[:-3]
    site[b]=r[0]
    cost[b]=r[3:]

sheet = pd.read_excel(output_pivot_table)
data = sheet.values
n=len(data)
m=len(data[0])

writer = pd.ExcelWriter(optimizer_output, engine='openpyxl')

lis = ['Site','Branch'] + t + ['Cost','VU']

branches = ['Branch']
sites = ['Site']
for i in range(2,n):
    br = data[i][0].split(' + ')
    b=''
    for j in br:
        b+= j.split(' ')[0] + ' + '
    b=b[:-3]
    branches.append(data[i][0])
    sites.append(site[b])


cc = ['Monthly Cost']
vv = ['Average Vehicle Utilisation']

for i in range(1,len(branches)):
    cc.append(0)
    vv.append(0)

cost_dw = [sites, branches, cc]        
vu_dw = [sites, branches, vv]

for j in range(1,m,2):
    date = data[0][j].split('/')
    date = date[0]+'-'+date[1]+'-'+date[2]
    print('         '+date)
    table = []
    table.append(lis)
    cost_list = [date]
    vu_list = [date]
    
    for i in range(2,n):
        br=data[i][0].upper().split(' + ')
        b=''
        for k in br:
            b+= k.split(' ')[0] + ' + '
        b=b[:-3]
        weight = data[i][j]
        vol = data[i][j+1]
        c=cost[b]
        
        prob, x = make_prob(num_trucks)
        
        prob+= sum([cc*xx for (cc,xx) in zip(c,x)])
        prob+= sum([tw*xx for (tw,xx) in zip(tweight,x)]) >= weight
        prob+= sum([tv*xx for (tv,xx) in zip(tvol,x)]) >= vol
        
        status = prob.solve()
        
        x = [value(xx) for xx in x]
        for i in range(len(x)):
            if x[i]==None:
                x[i]=0
        
        if weight == 0:
            vu = 0
        else:
            vuw = weight/(sum([tw*xx for (tw,xx) in zip(tweight,x)]))
            vuv = vol/(sum([tv*xx for (tv,xx) in zip(tvol,x)]))
            vu = round(max(vuw,vuv) *100 ,3)
        
        cost_list.append(value(prob.objective))
        vu_list.append(vu)
        
        table.append([site[b], b+' Branch'] + x + [value(prob.objective),vu])
    
    cost_dw.append(cost_list)
    vu_dw.append(vu_list) 
    
    df = pd.DataFrame(table)
    df.to_excel(writer, date, index=False, header=None)


cost_dw = np.array(cost_dw).T
vu_dw = np.array(vu_dw).T
    
cost_dw=cost_dw.tolist()
vu_dw=vu_dw.tolist()

n = len(cost_dw)
m = len(cost_dw[0])

for i in range(1,n):
    s = 0
    for j in range(2,m):
        k = cost_dw[i][j].split('.')
        if len(k) > 1:
            k = int(k[0])+pow(10,-1*len(k[1]))*int(k[1])
        else:
            k = int(k[0])
        cost_dw[i][j] = k
        s+= k
    cost_dw[i][2]=round(s,3)

for i in range(1,n):
    s=0
    c=0
    for j in range(2,m):
        try:
            k = vu_dw[i][j].split('.')
            if len(k) > 1:
                k = int(k[0])+pow(10,-1*len(k[1]))*int(k[1])
            else:
                k=int(k[0])
            vu_dw[i][j] = k
            if k != 0:
                c+=1
                s+=k
        except:
            a=1
    if c != 0:
        vu_dw[i][2] = round(s/c,3)


cost_df = pd.DataFrame(cost_dw)
vu_df = pd.DataFrame(vu_dw)

cost_df.to_excel(writer, 'BRS cost Daywise', index=False, header=None)
vu_df.to_excel(writer, 'Vehicle Utilisation Daywise', index=False, header=None)


writer.save()

print("Optimizer output saved!")
    
        