#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Jan 25 20:46:38 2019

@author: clement
"""

import numpy as np  
import math as m
import xlwings as xl
import pandas as pd

# Opening Excel Books/Sheets
# opens Excel Book "Test"
xl.Book(r"C:\Users\cleme/Desktop\PYTHON\Test.xlsx") 
xl.Book(r"/Users/clement/Desktop/ABE5663/PYTHON programing/Test.xlsx")           
# activates "Result" Sheet
xl.Sheet('Result').activate()                                                  
# Input parameter and guess varible values


w = np.array([0.027,0,0.132])                                                                                  
muMAX = 0.79
mud = 0.24
Ks = 0.0257
Ki = 3.97
a = 0.0249
b = 0.027
YCO2X = 200000000
Kla = 0.99
Hl = 29.14


# Input real data              
RAW = pd.read_excel (r"/Users/clement/Desktop/ABE5663/PYTHON programing/algae growth data.xlsx")
Time = list(RAW["Time"])
Cell = list(RAW["Cell concentration"])
CO2 = list(RAW["simulated CO2"])
daily_cell = list(RAW["Daily cell concentration"])
EPS = list(RAW["EPS"])

# Mass balance differenctial equation
# E1 = algae biomass
# E2 = EPS
# E3 = dissolved CO2
# A = the paramter we want to optimize
# p, V ,F and A need to be edited*********************

def Function(t,w,A):                                                          
    X,P,S = w
    E1 = ((muMAX*S*1000)/(Ks*(X*1000/Ki + 1)+S*1000)-mud)*X 
    E2 = a*E1 + b*X
    E3 = - muMAX*S*1000*X/YCO2X*(Ks + S*1000) + Kla*(0.004*44/Hl - S)
    return np.array([E1,E2,E3])

# RK4 function_1: to calculate the varible values
def RK4(Function,t0,tf,h,A):                                                      
    Y = []                                                                     
    T = []                                                                     
    n = (tf-t0)/h                                                         
    t = t0
    Y1 = []
    Y2 = []
    Y3 = []
    w = np.array([0.027,0,0.132])     
    while t <= tf:
        X,P,S = w
        Y.append(w)
        Y1.append(X)
        Y2.append(P)
        Y3.append(S)
        k1 = Function(t,w,A)
        k2 = Function(t + 0.5*n, w + 0.5*n*k1,A)
        k3 = Function(t + 0.5*n, w + 0.5*n*k2,A)
        k4 = Function(t + n, w + n*k3,A)
        w = w + (k1 + 2*k2 + 2*k3 + k4)*n/6
        T.append(t)
        t = t + n
        if w[2] < 0:
            w[2] = 0
  
    
    xl.Range('B2').options(transpose=True).value = T
    xl.Range('C2').options(transpose=True).value = Y1
    xl.Range('D2').options(transpose=True).value = Y2
    xl.Range('E2').options(transpose=True).value = Y3
    return 1 

# RK4 function_2: to calculate sum of square error, SSE
def RK4s(Function,t0,tf,h,A):                                                         
    Y = []                                                                 
    T = []                                                                     
    n = (tf-t0)/h                                                     
    t = t0
    Y1 = []
    Y2 = []
    Y3 = []
    SSE1 = 0
    SSE2 = 0
    SSE3 = 0
    w = np.array([0.027,0,0.132])     
    while t <= tf:
        X,P,S = w
        Y.append(w)
        Y1.append(X)
        Y2.append(P)
        Y3.append(S)
        k1 = Function(t,w,A)
        k2 = Function(t + 0.5*n, w + 0.5*n*k1,A)
        k3 = Function(t + 0.5*n, w + 0.5*n*k2,A)
        k4 = Function(t + n, w + n*k3,A)
        w = w + (k1 + 2*k2 + 2*k3 + k4)*n/6
        
        if EPS[int(t)] != 0:
           d1 = (P-EPS[int(t)])**2 
#        d2 = (P-real[int(t)])**2 
#        d3 = (S-real[int(t)])**2
           SSE1 += d1
#        SSE2 += d2
#        SSE3 += d3
        T.append(t)
        t = t + n
#    SSE = SSE1*SSE2*SSE3
    return SSE1   


# Parameter optimization based on the minimum SSE
# A = the initial value of parameter estimated this time
# d = step size
def NEWSSE(RK4s,A,d):
    f = []
    result = []
    SSE0 = 1000
    SSE = 100
    while SSE < SSE0:
        SSE0 = SSE
        SSE = RK4s(Function,0,57,57,A)
        A -= d
        f.append(A)
        result.append(SSE0)
   # xl.Range('B2').options(transpose=True).value = f
   # xl.Range('C2').options(transpose=True).value = result
      
    return A+d,SSE0    



RK4s(Function,0,57,57,0.8)
RK4(Function,0,57,57,200000000)
NEWSSE(RK4s,200000000,1)












 