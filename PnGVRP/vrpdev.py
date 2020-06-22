import os
import pandas as pd
import numpy as np
from pulp import *

# n = number of branches
# m = #types of trucks
# Nt[t] = number of trucks of type t
# K = total number of trucks (including all types) 
# Kt[k] gives the type of kth truck
# cost[i] = cost matrix for truck[i]
# demand_w[i] = weight demand of ith branch
# demand_v[i] = volume demand of ith branch 

Dataset = 'DP10'
maxSeconds = 15
output = 'output10.xlsx'
outputtxt = 'output10.txt'
demandFile = Dataset+'/demand.xlsx'
truckFile = Dataset+'/TruckCap.xlsx'
costFile = Dataset+'/cost.xlsx'

def make_prob(n, K):
    prob = LpProblem('VRP',LpMinimize)
    #    x_k_i_j    where, k for #truck,    i,j for index (i,j)
    x = []
    for k in range(K):
        xi = []
        for i in range(n):
            xij = []
            for j in range(n):
                a = 'x_'+str(k)+'_'+str(i)+'_'+str(j)
                a = LpVariable(a, lowBound = 0, upBound = 1, cat = 'Integer')
                xij.append(a)
            xi.append(xij)
        x.append(xi)
    
    return prob, x

def get_values(y, n, K):
    x = []
    for k in range(K):
        xi = []
        for i in range(n):
            xij = []
            for j in range(n):
                a = value(y[k][i][j])
                xij.append(a)
            xi.append(xij)
        x.append(xi)
    return x

def false(n):
    vis = []
    for i in range(n):
        vis.append(False)
    return vis


if __name__ == "__main__":

    Truck = pd.read_excel(truckFile)
    trucks = Truck.columns[1:]
    m = trucks.shape[0]
    t_w = Truck.values[2][1:]
    t_v = Truck.values[1][1:]

    # branches 
    branches = pd.read_excel(costFile, trucks[0]).columns
    n = len(branches)

    # cost[i] = cost matrix for truck[i]
    cost = []
    for t in trucks:
        Cost = pd.read_excel(costFile, t).values
        cost.append(Cost)
        
    Demand = pd.read_excel(demandFile)
    cnames = Demand.columns

    br = [i.split(' ')[0] for i in Demand[cnames[0]][1:].tolist()]
    Demand = Demand.values
    d_w = {}
    d_v = {}
    for i in range(len(br)):
        d_w[br[i]] = Demand[i+1][1]
        d_v[br[i]] = Demand[i+1][2]

    # demand_w[i] = weight demand, demand_v[i] = volume demand of ith branch 
    demand_w = [0]
    demand_v = [0]

    for i in branches[1:]:
        
        demand_w.append(d_w[i])
        demand_v.append(d_v[i])


    Nt = [n for t in trucks];        # Nt[k] = number of trucks of type k
    K = sum(Nt)
    Kt = []                          # Kt[i] gives type of kth truck
    for i in range(m):
        for j in range(Nt[i]):
            Kt.append(i)
    mat = [(i,j) for i in range(n) for j in range(n)]

    prob, x = make_prob(n,K)

    # Objective function
    s = 0
    for k in range(K):
        s += sum([cost[Kt[k]][i][j]*x[k][i][j] for i,j in mat])

    prob+= s

    #Constraint 1
    for i in range(1,n):
        s = sum([x[k][i][j] for k in range(K) for j in range(n)])
        prob+= s==1


    #Constraint 2
    for j in range(1,n):
        s = sum([x[k][i][j] for k in range(K) for i in range(n)])
        prob+= s==1


    #Constraint 3
    for k in range(K):
        s = sum([x[k][0][j] for j in range(n)])
        prob+= s==1

    #Constraint 4
    for k in range(K):
        s = sum([x[k][i][0] for i in range(n)])
        prob+= s==1

    #Constraint 5 
    for k in range(K):
        s_w=0
        s_v=0
        for i in range(n):
            s_w = s_w + sum([x[k][i][j]*demand_w[j] for j in range(n)])
            s_v = s_v + sum([x[k][i][j]*demand_v[j] for j in range(n)])
        prob+= s_w <= t_w[Kt[k]]
        prob+= s_v <= t_v[Kt[k]]
        
    #Constraint 6
    for k in range(K):
        s_w_l=0
        s_v_l=0
        s_w_r=0
        s_v_r=0
        for i in range(n):
            s_w_l = s_w_l + sum([x[k][i][j]*demand_w[j] for j in range(n)])
            s_v_l = s_v_l + sum([x[k][i][j]*demand_v[j] for j in range(n)])
        for j in range(n):
            s_w_r = s_w_r + sum([x[k][i][j]*demand_w[i] for i in range(n)])
            s_v_r = s_v_r + sum([x[k][i][j]*demand_v[i] for i in range(n)])
        prob+= s_w_l + (-1*s_w_r) == 0
        prob+= s_v_l + (-1*s_v_r) == 0    

    #Constraint 7
    for k in range(K):
        for i in range(1, n):
            for j in range(i,n):
                prob+= x[k][i][j]+x[k][j][i]<=1
                
    #Constraint 8
    for k in range(K):
        s = sum([x[k][i][j] for i in range(1,n) for j in range(1,n)])
        prob+= s<=2
                
        

    status = prob.solve(PULP_CBC_CMD(maxSeconds=maxSeconds , msg=1, fracGap=0))
    print(' STATUS : '+str(LpStatus[status]))
    x = get_values(x, n, K)

    X = []
    for i in x:
        for j in i:
            X.append(j)
        X.append([])
    df=pd.DataFrame(X)
    df.to_excel(output)


    ans = []
    anstype = []

    for k in range(len(x)):
        a = x[k]
        i = 0
        l = []
        vis = false(n)
        while (not vis[i]):
            vis[i]=True
            l.append(branches[i])
            j=i
            for k in range(n):
                if a[i][k]==1:
                    j=k
            i=j
        if len(l) > 1:
            ans.append(l)
            anstype.append(trucks[Kt[k]])


    f = open(outputtxt, 'w')
    L = []
    for i in range(len(ans)):
        l = anstype[i]+' : '
        for j in ans[i]:
            l = l+j+' -> '
        l=l[:-4]+'\n'
        L.append(l)
    f.writelines(L)
    f.close()