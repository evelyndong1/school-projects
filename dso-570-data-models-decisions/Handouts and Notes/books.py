#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Mar 26 15:42:05 2019

@author: pengshi
"""

import pandas as pd
from gurobipy import Model, GRB

def optimize(inputFile,outputFile):
    '''Function which takes in two input arguments:
        - inputFile: the path to the input data. (.xlsx format)
        - outputFile: the path to the output data. (.xlsx format)
        The outputFile gives the optimal set of books to carry as well as 
        the minimum number of books needed.'''
    genres=pd.read_excel(inputFile,sheet_name='genres',index_col=0).fillna(0)
    requirements=pd.read_excel(inputFile,sheet_name='requirements',index_col=0)
    
    mod=Model()
    I=genres.index
    J=genres.columns

    x=mod.addVars(I,vtype=GRB.BINARY)
    mod.setObjective(sum(x[i] for i in I))
    for j in J:
        mod.addConstr(sum(genres.loc[i,j]*x[i] for i in I)>=requirements.loc[j])
    mod.setParam('outputflag',False)
    mod.optimize()

    writer=pd.ExcelWriter(outputFile)
    carry=[]
    for i in I:
        if x[i].x:
            carry.append(i)
    pd.DataFrame(carry,columns=['books'])\
        .to_excel(writer,sheet_name='optimal_decision')
    pd.DataFrame([mod.objVal],columns=['books_needed'])\
        .to_excel(writer,sheet_name='objective',index=False)
    writer.save()
    
    
if __name__=='__main__':
    import sys, os
    if len(sys.argv)!=3:
        print('Correct syntax: python books.py inputFile outputFile')
    else:
        inputFile=sys.argv[1]
        outputFile=sys.argv[2]
        if os.path.exists(inputFile):
            optimize(inputFile,outputFile)
            print(f'Successfully optimized. Results in "{outputFile}"')
        else:
            print(f'File "{inputFile}" not found!')
