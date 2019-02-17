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

def drawCurve(ax,xlabel,curvelegend,x,yArray):

    lineColor = ['b+-','ro-','gs-','h--']

    for i,y in enumerate(yArray):
        ax.plot(x, y, lineColor[i], lw=2, zorder=30)

    ax.grid()
    ax.set_xlabel(xlabel, fontweight='bold', fontsize=15)
    ax.set_xticklabels([])
    ax.set_yticklabels([])

    ax.legend(ax.lines, curvelegend, fancybox=True, shadow=True, prop={'size': 15})

def drawPvtCurve(pvt):
    fig = plt.figure('PVT')

    # draw pvt_oil
    df_PVTo = pvt[0]
    o_x = df_PVTo['P']
    o_y0 = df_PVTo['MUO']
    o_y1 = df_PVTo['BO']
    o_y2 = df_PVTo['RSO']

    ax = fig.add_subplot(2,1,1)

    lMap1      = {'x':o_x,'y':o_y0,'id':0,'ylabel':'MUO'}
    lMap2      = {'x':o_x,'y':o_y1,'id':1,'ylabel':'BO'}
    rMap       = {'x':o_x,'y':o_y2,'id':2,'ylabel':'RSO'}
    leftMaps   = [lMap1,lMap2]
    rightMaps  = [rMap]

    draw2YCurve(ax,'Oil PVT','P',leftMaps,rightMaps)

    # draw pvt_gas
    df_PVTg = pvt[2]
    g_x = df_PVTg['P']
    g_y0 = df_PVTg['MUG']
    g_y1 = df_PVTg['BG']

    ax = fig.add_subplot(2,1,2)

    lMap1      = {'x':g_x,'y':g_y0,'id':0,'ylabel':'MUG'}
    lMap2      = {'x':g_x,'y':g_y1,'id':1,'ylabel':'BG'}
    leftMaps   = [lMap1,lMap2]
    rightMaps  = []

    draw2YCurve(ax,'Gas PVT','P',leftMaps,rightMaps)

    plt.show()

def readPvt(sht):
    iRowPvtSec= getKeyWordBeginRow(sht,1,'PVT')

    iRowPbo   = getKeyWordBeginRow(sht,iRowPvtSec,'PBO')
    fPbo      = sht.range((iRowPbo,2),(iRowPbo,2)).value
    iRowMuo_s   = getKeyWordBeginRow(sht,iRowPvtSec,'Viso_SLOPE')
    fMuo_s      = sht.range((iRowMuo_s,2),(iRowMuo_s,2)).value
    iRowBo_s   = getKeyWordBeginRow(sht,iRowPvtSec,'Bo_SLOPE')
    fBo_s      = sht.range((iRowBo_s,2),(iRowBo_s,2)).value

    iRowBegin = getKeyWordBeginRow(sht,iRowPvtSec,'OIL-PVT')
    iRowEnd   = getKeyWordEndRow(sht,iRowBegin)

    df_PVTo = sht.range((iRowBegin,2),(iRowEnd,5)).options(pd.DataFrame,index=False).value


    # pvtw
    iRowBegin = getKeyWordBeginRow(sht,iRowPvtSec,'WATER-PVT')
    iRowEnd   = getKeyWordEndRow(sht,iRowBegin)
    df_PVTw = sht.range((iRowBegin,2),(iRowEnd,4)).options(pd.DataFrame,index=False).value
    df_to_nparray = df_PVTw.to_records(index=True)

    # pvtg
    iRowBegin = getKeyWordBeginRow(sht,iRowPvtSec,'GAS-PVT')
    iRowEnd   = getKeyWordEndRow(sht,iRowBegin)
    df_PVTg = sht.range((iRowBegin,2),(iRowEnd,4)).options(pd.DataFrame,index=False).value

    # density
    iDenRow = getKeyWordBeginRow(sht,iRowPvtSec,u'DENSITY')
    df_density = sht.range((iDenRow,2),(iDenRow+1,4)).options(pd.DataFrame).value
    
    # rock
    iRockRow = getKeyWordBeginRow(sht,iRowPvtSec,u'CROCK')
    fRock = sht.range((iRockRow,3),(iRockRow,3)).value

    return (df_PVTo,df_PVTw,df_PVTg)

def drawRelPermCurve(relPerm):
    fig = plt.figure('Rel. Perm')

    # OIL-WATER
    df_SWOF = relPerm[0]

    wo_x = df_SWOF['SAT']
    wo_y0 = df_SWOF['KRW']
    wo_y1 = df_SWOF['KROW']

    wo_yArray = [wo_y0,wo_y1]
    ax = fig.add_subplot(2,1,1)

    drawCurve(ax,'SAT',['KRW','KROW'],wo_x,wo_yArray)

    # GAS-LIQUID
    df_SGLF = relPerm[1]

    x = df_SGLF['SLIQ']
    y0 = df_SGLF['KRG']
    y1 = df_SGLF['KROG']

    yArray = [y0,y1]
    ax = fig.add_subplot(2,1,2)

    drawCurve(ax,'SLIQ',['KRG','KROG'],x,yArray)

    plt.show()


def readRelativePerm(sht):
    iRowPvtSec= getKeyWordBeginRow(sht,1,'PVT')
    # water-oil
    iRowBegin = getKeyWordBeginRow(sht,iRowPvtSec,'WATER-OIL')
    iRowEnd   = getKeyWordEndRow(sht,iRowBegin)

    df_SWOF = sht.range((iRowBegin,2),(iRowEnd,5)).options(pd.DataFrame,index=False).value

    # GAS-LIQUID
    iRowBegin = getKeyWordBeginRow(sht,iRowPvtSec,'GAS-LIQUID')
    iRowEnd   = getKeyWordEndRow(sht,iRowBegin)

    df_SGLF = sht.range((iRowBegin,2),(iRowEnd,5)).options(pd.DataFrame,index=False).value

    return (df_SWOF,df_SGLF)

def Pvt():
    sht = xw.Book.caller().sheets[0]
    pvt = readPvt(sht)
    drawPvtCurve(pvt)


def RelPerm():
    sht = xw.Book.caller().sheets[0]
    relPerm = readRelativePerm(sht) 
    drawRelPermCurve(relPerm)   

'''
def readPropSection(sht):
    pvt = readPvt(sht)
    drawPvtCurve(pvt)

    relPerm = readRelativePerm(sht) 
    drawRelPermCurve(relPerm)   

if __name__ == '__main__':
    # Expects the Excel file next to this source file, adjust accordingly.
    xw.Book('tutor3.xlsm').set_mock_caller()

    sht = xw.Book.caller().sheets[0]

    readPropSection(sht)
'''