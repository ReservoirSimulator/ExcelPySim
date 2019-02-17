from __future__ import division
import sys
import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D
from matplotlib import cm
from matplotlib.ticker import LinearLocator, FormatStrFormatter

import xlwings as xw

import datetime

class CWell:
    wellCount = 0
    
    def __init__(self, name, hI,hJ):
        self.name = name
        self.head_I = hI
        self.head_J = hJ
        CWell.wellCount += 1
        self.id   = CWell.wellCount


def getKeyWordBeginRow(sht,iRow,strKW):
    index = 'A'
    while (sht.range(index+str(iRow)).value != strKW):
        iRow = iRow+1

    return iRow

def getKeyWordEndRow(sht,irow):
    irow = irow+1
    index = 'B'
    while (sht.range(index+str(irow)).value != None):        
        irow = irow+1

    return irow-1

def draw2YCurve(ax,title,xlabel,leftMaps,rightMaps):
    lineColor = ['b+-','ro-','gs-','h--']
    lmaxV = 0
    lminV = 10000
    rmaxV = 0
    rminV = 10000

    for lMap in leftMaps:
        x = lMap['x']
        y = lMap['y']
        i = lMap['id']
        ax.plot(x,y,lineColor[i],lw=2,label=lMap['ylabel'])

        maxNum = max(y)
        minNum = min(y)
        if(maxNum>lmaxV):
            lmaxV = maxNum
        if(minNum<lminV):
            lminV = minNum

    for rMap in rightMaps:
        x = rMap['x']
        y = rMap['y']
        i = rMap['id']
        ax2 = ax.twinx()

        ax2.plot(x,y,lineColor[i],lw=2,label=rMap['ylabel'])

        maxNum = max(y)
        minNum = min(y)
        if(maxNum>rmaxV):
            rmaxV = maxNum
        if(minNum<rminV):
            rminV = minNum

    ax.grid()
    ax.set_title(title)
    ax.set_xlabel(xlabel, fontweight='bold', fontsize=15)
    ax.legend(loc='upper left')
    ax.set_ylim(lminV,lmaxV)
    if(len(rightMaps)>0):
        ax2.legend(loc='upper right')
        ax2.set_ylim(rminV,rmaxV)

def drawWells(wells):
    fig = plt.figure('WELL PRODUCTION')

    wellNUm = len(wells)
    subRow  = int(wellNUm/2)
    if(wellNUm%2>0):
        subRow += 1
    wellId = 1
    for well in wells.values():
        wellName = well.name
        df_prod = well.df_prod
        df_prod['Date'] = df_prod['Date'].astype('int64')
        df_prod['Date'] = pd.to_datetime(df_prod['Date'],format="%Y%m")
        x = df_prod['Date']

        y1 = df_prod[u'Monthly Oil Production']
        y2 = df_prod[u'Monthly Water Production']
    
        ax = fig.add_subplot(subRow,2,wellId)

        lMap1      = {'x':x,'y':y1,'id':0,'ylabel':'Monthly Oil Production Rate'}
        lMap2      = {'x':x,'y':y2,'id':1,'ylabel':'Monthly Water Production Rate'}
        leftMaps   = [lMap1,lMap2]
        rightMaps  = []

        draw2YCurve(ax,wellName,'Date',leftMaps,rightMaps)

        wellId += 1

    plt.show()

def readDynamicSection(sht):
    iRow = 1
    index = 'A'
    wells = {}
    while (sht.range(index+str(iRow)).value!='END'):
        iRow = iRow+1
        kw =  sht.range(index+str(iRow)).value
        if(kw==u'Well Name'):
            wellName = sht.range('B'+str(iRow)).value
            iwellDefRow = getKeyWordBeginRow(sht,iRow,u'well type')
            head_I   = sht.range('D'+str(iwellDefRow)).options(numbers=int).value
            head_J   = sht.range('F'+str(iwellDefRow)).options(numbers=int).value
            well = CWell(wellName,head_I,head_J)
            wells[wellName] = well
        
        elif(kw == u'Well Completion'):
            comp = []
            wName = sht.range('B'+str(iRow)).value
            well = wells.get(wName)
            iRow = iRow + 3
            iEndComp = getKeyWordEndRow(sht,iRow)
            for i_layer in range(iRow,iEndComp+1):
                nLayer = sht.range('D'+str(i_layer)).options(numbers=int).value
                comp.append(nLayer)
            
            iRow = iEndComp
            well.comp = comp
        
        elif(kw == u'Production'):
            wName = sht.range('A'+str(iRow+1)).value
            iEndProdRow = getKeyWordEndRow(sht,iRow)
            prod_date = sht.range((iRow+1,2),(iEndProdRow,2)).options(numbers=str).value
            df_prod = sht.range((iRow,2),(iEndProdRow,7)).options(pd.DataFrame,index=False).value
            iRow = iEndProdRow
        
        
            well = wells.get(wName)
            well.df_prod = df_prod
        

            np_prodDate = np.array(prod_date)
            asciiList = [n.encode("ascii", "ignore") for n in np_prodDate]
            well.prod_date = asciiList

    drawWells(wells)

def DynamicStep0():

    sht = xw.Book.caller().sheets[1]

    readDynamicSection(sht)

'''
if __name__ == '__main__':
    # Expects the Excel file next to this source file, adjust accordingly.
    xw.Book('tutor2.xlsm').set_mock_caller()
    DynamicStep0()
'''