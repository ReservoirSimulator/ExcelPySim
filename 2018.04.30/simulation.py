from __future__ import division
import sys
import os
import h5py
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import xlwings as xw

from mpl_toolkits.mplot3d import Axes3D
from matplotlib import cm
from matplotlib.ticker import LinearLocator, FormatStrFormatter
import datetime
import subprocess

realPath = os.path.dirname(os.path.realpath(__file__))
hdf5File = realPath+'\\simModel\\simModel.hdf5'
filename = realPath+'\\simModel\\model.dxx'
sht = xw.Book.caller().sheets[0]
gridVaryOpt={ }
defaultDim = 200
kwMap = {'dx':'DX','dy':'DY','dz':'DZ','por':'POROSITY','permx':'KX','permy':'PERMY',
        'permz':'PERMZ'}
class CWell:
    wellCount = 0
    
    def __init__(self, name, hI,hJ):
        self.name = name
        self.head_I = hI
        self.head_J = hJ
        CWell.wellCount += 1
        self.id   = CWell.wellCount


def getKeyWordBeginRow(iRow,strKW):
    index = 'A'
    while (sht.range(index+str(iRow)).value != strKW):
        iRow = iRow+1

    return iRow

def getKeyWordEndRow(irow):
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
    
def readGeneralInfo():
    nx = sht.range('B3').options(numbers=int).value
    ny = sht.range('C3').options(numbers=int).value
    nz = sht.range('D3').options(numbers=int).value

    gridVaryOpt['xopt'] = sht.range('B4').value
    gridVaryOpt['yopt'] = sht.range('C4').value
    gridVaryOpt['zopt'] = sht.range('D4').value

    gridVaryOpt['por'] = sht.range('F3').value
    gridVaryOpt['kx'] = sht.range('G3').value
    gridVaryOpt['ky'] = sht.range('H3').value
    gridVaryOpt['kz'] = sht.range('I3').value

    return nx,ny,nz

nx,ny,nz = readGeneralInfo()

def createHDF5(nx,ny,nz):
    
    ni = nx
    nj = ny
    nk = nz

    fModel = h5py.File(hdf5File, "w")

    staticEntry = fModel.create_group(u"StaticGroup")
    staticEntry.attrs[u'nx'] = nx
    staticEntry.attrs[u'ny'] = ny
    staticEntry.attrs[u'nz'] = nz

    fModel.create_group(u"DynamicGroup")
    fModel.create_group(u"PvtGroup")
    fModel.create_group(u"InitGroup")
    fModel.create_dataset(u'StaticGroup/dx', (nk,nj,ni), dtype='f')
    fModel.create_dataset(u'StaticGroup/dy', (nk,nj,ni), dtype='f')
    fModel.create_dataset(u'StaticGroup/dz', (nk,nj,ni), dtype='f')
    fModel.create_dataset(u'StaticGroup/por', (nk,nj,ni), dtype='f')
    fModel.create_dataset(u'StaticGroup/permx', (nk,nj,ni), dtype='f')
    fModel.create_dataset(u'StaticGroup/permy', (nk,nj,ni), dtype='f')
    fModel.create_dataset(u'StaticGroup/permz', (nk,nj,ni), dtype='f')
    fModel.create_dataset(u'StaticGroup/tops', (nj,ni), dtype='f')

    fModel.close()

def init():

    sht.range((9,3),(9,defaultDim+3-1)).color = (255,255,255)
    sht.range((10,3),(10,defaultDim+3-1)).color = (255,255,255)
    sht.range((11,3),(11,defaultDim+3-1)).color = (255,255,255)
    sht.range((17,3),(17,defaultDim+3-1)).color = (255,255,255)
    sht.range((18,3),(18,defaultDim+3-1)).color = (255,255,255)
    sht.range((19,3),(19,defaultDim+3-1)).color = (255,255,255)
    sht.range((20,3),(20,defaultDim+3-1)).color = (255,255,255)


def defineGrid():
    init()

    createHDF5(nx,ny,nz)

#prepare gui    
    if gridVaryOpt['xopt'] == 'CON':
        sht.range('A9').value = 'DX'
        sht.range('B9').value = 'CON'
        sht.range('C9').color = (102,255,102)
    elif gridVaryOpt['xopt'] == 'XVAR':
        sht.range('A9').value = 'DX'
        sht.range('B9').value = 'XVAR'
        sht.range((9,3),(9,nx+3-1)).color = (102,255,102)

    if gridVaryOpt['yopt'] == 'CON':
        sht.range('A10').value = 'DY'
        sht.range('B10').value = 'CON'
        sht.range('C10').color = (102,255,102)
    elif gridVaryOpt['yopt'] == 'YVAR':
        sht.range('A10').value = 'DY'
        sht.range('B10').value = 'YVAR'
        sht.range((10,3),(10,ny+3-1)).color = (102,255,102)
        
    if gridVaryOpt['zopt'] == 'CON':
        sht.range('A11').value = 'DZ'
        sht.range('B11').value = 'CON'
        sht.range('C11').color = (102,255,102)
    elif gridVaryOpt['zopt'] == 'ZVAR':
        sht.range('A11').value = 'DZ'
        sht.range('B11').value = 'ZVAR'
        sht.range((11,3),(11,nz+3-1)).color = (102,255,102)

    if gridVaryOpt['por'] == 'CON':
        sht.range('A17').value = 'POR'
        sht.range('B17').value = 'CON'
        sht.range('C17').color = (102,255,102)
    elif gridVaryOpt['por'] == 'ZVAR':
        sht.range('A17').value = 'POR'
        sht.range('B17').value = 'ZVAR'
        sht.range((17,3),(17,nz+3-1)).color = (102,255,102)

    if gridVaryOpt['kx'] == 'CON':
        sht.range('A18').value = 'PERMX'
        sht.range('B18').value = 'CON'
        sht.range('C18').color = (102,255,102)
    elif gridVaryOpt['kx'] == 'ZVAR':
        sht.range('A18').value = 'PERMX'
        sht.range('B18').value = 'ZVAR'
        sht.range((18,3),(18,nz+3-1)).color = (102,255,102)
    
    if gridVaryOpt['ky'] == 'CON':
        sht.range('A19').value = 'PERMY'
        sht.range('B19').value = 'CON'
        sht.range('C19').color = (102,255,102)
    elif gridVaryOpt['ky'] == 'ZVAR':
        sht.range('A19').value = 'PERMY'
        sht.range('B19').value = 'ZVAR'
        sht.range((19,3),(19,nz+3-1)).color = (102,255,102)

    if gridVaryOpt['kz'] == 'CON':
        sht.range('A20').value = 'PERMZ'
        sht.range('B20').value = 'CON'
        sht.range('C20').color = (102,255,102)
    elif gridVaryOpt['kz'] == 'ZVAR':
        sht.range('A20').value = 'PERMZ'
        sht.range('B20').value = 'ZVAR'
        sht.range((20,3),(20,nz+3-1)).color = (102,255,102)

def readGrid():

    fModel = h5py.File(hdf5File, "a")
    dset_dx = fModel['StaticGroup/dx']
    dset_dy = fModel['StaticGroup/dy']
    dset_dz = fModel['StaticGroup/dz']
    dset_por = fModel['StaticGroup/por']
    dset_kx = fModel['StaticGroup/permx']
    dset_ky = fModel['StaticGroup/permy']
    dset_kz = fModel['StaticGroup/permz']
    dset_tops = fModel['StaticGroup/tops']

    iDxRow = getKeyWordBeginRow(1,'DX')
    iDyRow = getKeyWordBeginRow(1,'DY')
    iDzRow = getKeyWordBeginRow(1,'DZ')
    if gridVaryOpt['xopt'] == 'XVAR' :
        dxVar = sht.range((iDxRow,3),(iDxRow,nx+3-1)).options(np.array).value
        idArray = np.arange(nx)
        for id in idArray:
            dset_dx[:,:,id] = dxVar[id]
    elif gridVaryOpt['xopt'] == 'CON' :
        dset_dx[:,:,:] = sht.range((iDxRow,3)).options(numbers=float).value

    if gridVaryOpt['yopt'] == 'YVAR' :
        dyVar = sht.range((iDyRow,3),(iDyRow,ny+3-1)).options(np.array).value
        idArray = np.arange(ny)
        for id in idArray:
            dset_dy[:,id,:] = dyVar[id]
    elif gridVaryOpt['yopt'] == 'CON' :
        dset_dy[:,:,:] = sht.range((iDyRow,3)).options(numbers=float).value

    if gridVaryOpt['zopt'] == 'ZVAR' :
        dzVar = sht.range((iDzRow,3),(iDzRow,nz+3-1)).options(np.array).value
        idArray = np.arange(nz)
        for id in idArray:
            dset_dz[id,:,:] = dzVar[id]
    elif gridVaryOpt['zopt'] == 'CON' :
        dset_dz[:,:,:] = sht.range((iDzRow,3)).options(numbers=float).value

    itopRow = getKeyWordBeginRow(1,'TOPS')
    dset_tops[:,:] = sht.range((itopRow,3)).options(numbers=float).value
    
    iporRow = getKeyWordBeginRow(1,'POR')
    ikxRow  = getKeyWordBeginRow(1,'PERMX')
    ikyRow  = getKeyWordBeginRow(1,'PERMY')
    ikzRow  = getKeyWordBeginRow(1,'PERMZ')
    if gridVaryOpt['por'] == 'ZVAR' :
        porVar = sht.range((iporRow,3),(iporRow,nz+3-1)).options(np.array).value
        idArray = np.arange(nz)
        for id in idArray:
            dset_por[id,:,:] = porVar[id]
    elif gridVaryOpt['por'] == 'CON' :
        dset_por[:,:,:] = sht.range((iporRow,3)).options(numbers=float).value

    if gridVaryOpt['kx'] == 'ZVAR' :
        kxVar = sht.range((ikxRow,3),(ikxRow,nz+3-1)).options(np.array).value
        idArray = np.arange(nz)
        for id in idArray:
            dset_kx[id,:,:] = kxVar[id]
    elif gridVaryOpt['kx'] == 'CON' :
        dset_kx[:,:,:] = sht.range((ikxRow,3)).options(numbers=float).value

    if gridVaryOpt['ky'] == 'ZVAR' :
        kyVar = sht.range((ikyRow,3),(ikyRow,nz+3-1)).options(np.array).value
        idArray = np.arange(nz)
        for id in idArray:
            dset_ky[id,:,:] = kyVar[id]
    elif gridVaryOpt['ky'] == 'CON' :
        dset_ky[:,:,:] = sht.range((ikyRow,3)).options(numbers=float).value

    if gridVaryOpt['kz'] == 'ZVAR' :
        kzVar = sht.range((ikzRow,3),(ikzRow,nz+3-1)).options(np.array).value
        idArray = np.arange(nz)
        for id in idArray:
            dset_kz[id,:,:] = kzVar[id]
    elif gridVaryOpt['kz'] == 'CON' :
        dset_kz[:,:,:] = sht.range((ikzRow,3)).options(numbers=float).value

    fModel.flush()
    fModel.close()

def drawPvtCurve():
    fig = plt.figure('PVT')

    iRowPvtSec= getKeyWordBeginRow(1,'PVT')

    iRowBegin = getKeyWordBeginRow(iRowPvtSec,'OIL-PVT')
    iRowEnd   = getKeyWordEndRow(iRowBegin)

    df_PVTo = sht.range((iRowBegin,2),(iRowEnd,5)).options(pd.DataFrame,index=False).value

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

    # pvtg
    iRowBegin = getKeyWordBeginRow(iRowPvtSec,'GAS-PVT')
    iRowEnd   = getKeyWordEndRow(iRowBegin)
    df_PVTg = sht.range((iRowBegin,2),(iRowEnd,4)).options(pd.DataFrame,index=False).value

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

def readPvt(f):
    iRowPvtSec= getKeyWordBeginRow(1,'PVT')

    iRowPbo   = getKeyWordBeginRow(iRowPvtSec,'PBO')
    fPbo      = sht.range((iRowPbo,2),(iRowPbo,2)).value
    iRowMuo_s   = getKeyWordBeginRow(iRowPvtSec,'Viso_SLOPE')
    fMuo_s      = sht.range((iRowMuo_s,2),(iRowMuo_s,2)).value
    iRowBo_s   = getKeyWordBeginRow(iRowPvtSec,'Bo_SLOPE')
    fBo_s      = sht.range((iRowBo_s,2),(iRowBo_s,2)).value

    iRowBegin = getKeyWordBeginRow(iRowPvtSec,'OIL-PVT')
    iRowEnd   = getKeyWordEndRow(iRowBegin)

    df_PVTo = sht.range((iRowBegin,2),(iRowEnd,5)).options(pd.DataFrame).value

    # df to hdf5
    df_to_nparray = df_PVTo.to_records(index=True)

    #pvto_nparray = np.asarray(df_to_nparray)
    # create dataset
    f['PvtGroup/PVTO'] = df_to_nparray

    # pb
    f['PvtGroup'].attrs[u'PB']   = fPbo
    f['PvtGroup'].attrs[u'MUO_SLOPE'] = fMuo_s
    f['PvtGroup'].attrs[u'BO_SLOPE']  = fBo_s
    # pvtw
    iRowBegin = getKeyWordBeginRow(iRowPvtSec,'WATER-PVT')
    iRowEnd   = getKeyWordEndRow(iRowBegin)
    df_PVTw = sht.range((iRowBegin,2),(iRowEnd,4)).options(pd.DataFrame).value
    df_to_nparray = df_PVTw.to_records(index=True)
    f['PvtGroup/PVTW'] = df_to_nparray

    # pvtg
    iRowBegin = getKeyWordBeginRow(iRowPvtSec,'GAS-PVT')
    iRowEnd   = getKeyWordEndRow(iRowBegin)
    df_PVTg = sht.range((iRowBegin,2),(iRowEnd,4)).options(pd.DataFrame).value
    df_to_nparray = df_PVTg.to_records(index=True)
    f['PvtGroup/PVTG'] = df_to_nparray

    # density
    iDenRow = getKeyWordBeginRow(iRowPvtSec,u'DENSITY')
    df_density = sht.range((iDenRow,2),(iDenRow+1,4)).options(pd.DataFrame).value
    df_to_nparray = df_density.to_records(index=True)
    f['PvtGroup/DENSITY'] = df_to_nparray
    
    # rock
    iRockRow = getKeyWordBeginRow(iRowPvtSec,u'CROCK')
    fRock = sht.range((iRockRow,3),(iRockRow,3)).value
    f['PvtGroup'].attrs[u'ROCK'] = fRock

def drawRelPermCurve():
    fig = plt.figure('Rel. Perm')
    #ax = fig.add_subplot(111)

    iRowPvtSec= getKeyWordBeginRow(1,'PVT')
    # water-oil
    iRowBegin = getKeyWordBeginRow(iRowPvtSec,'WATER-OIL')
    iRowEnd   = getKeyWordEndRow(iRowBegin)
    
    df_SWOF = sht.range((iRowBegin,2),(iRowEnd,5)).options(pd.DataFrame,index=False).value

    wo_x = df_SWOF['SAT']
    wo_y0 = df_SWOF['KRW']
    wo_y1 = df_SWOF['KROW']

    wo_yArray = [wo_y0,wo_y1]
    ax = fig.add_subplot(2,1,1)

    drawCurve(ax,'SAT',['KRW','KROW'],wo_x,wo_yArray)

    # GAS-LIQUID
    iRowBegin = getKeyWordBeginRow(iRowPvtSec,'GAS-LIQUID')
    iRowEnd   = getKeyWordEndRow(iRowBegin)

    df_SGLF = sht.range((iRowBegin,2),(iRowEnd,5)).options(pd.DataFrame,index=False).value

    x = df_SGLF['SLIQ']
    y0 = df_SGLF['KRG']
    y1 = df_SGLF['KROG']

    yArray = [y0,y1]
    ax = fig.add_subplot(2,1,2)

    drawCurve(ax,'SLIQ',['KRG','KROG'],x,yArray)

    plt.show()


def readRelativePerm(f):
    iRowPvtSec= getKeyWordBeginRow(1,'PVT')
    # water-oil
    iRowBegin = getKeyWordBeginRow(iRowPvtSec,'WATER-OIL')
    iRowEnd   = getKeyWordEndRow(iRowBegin)

    df_SWOF = sht.range((iRowBegin,2),(iRowEnd,5)).options(pd.DataFrame,index=False).value
    #sht.range('G29:J39').options(pd.DataFrame).value = df_PVTo
    # df to hdf5
    df_to_nparray = df_SWOF.to_records(index=True)
    # create dataset
    f['PvtGroup/SWOF'] = df_to_nparray

    # GAS-LIQUID
    iRowBegin = getKeyWordBeginRow(iRowPvtSec,'GAS-LIQUID')
    iRowEnd   = getKeyWordEndRow(iRowBegin)

    df_SGLF = sht.range((iRowBegin,2),(iRowEnd,5)).options(pd.DataFrame,index=False).value
    #sht.range('G29:J39').options(pd.DataFrame).value = df_PVTo
    # df to hdf5
    df_to_nparray = df_SGLF.to_records(index=True)
    # create dataset
    f['PvtGroup/SGLF'] = df_to_nparray

def readPropSection():
    f = h5py.File(hdf5File,'a')
    readPvt(f)
    readRelativePerm(f)    

    f.close()

def readInitSection():
    f = h5py.File(hdf5File,'a')
    iRowInitSec= getKeyWordBeginRow(1,'INITIAL SECTION')
    iRowEquil   = getKeyWordBeginRow(iRowInitSec,'EQUIL')
    fRefP      = sht.range((iRowEquil+1,2),(iRowEquil+1,2)).value
    fRefDepth  = sht.range((iRowEquil+1,3),(iRowEquil+1,3)).value
    fOWC       = sht.range((iRowEquil+1,4),(iRowEquil+1,4)).value
    fOGC       = sht.range((iRowEquil+1,5),(iRowEquil+1,5)).value

    f[u'InitGroup'].attrs[u'RefPress'] = fRefP
    f[u'InitGroup'].attrs[u'RefDepth'] = fRefDepth
    f[u'InitGroup'].attrs[u'OWC']      = fOWC
    f[u'InitGroup'].attrs[u'OGC']      = fOGC

    f.close()

def drawWells():
    fig = plt.figure('WELL PRODUCTION')

    iRow = getKeyWordBeginRow(1,'SCHEDULE SECTION')
    index = 'A'
    wells = {}
    while (sht.range(index+str(iRow)).value!='END'):
        iRow = iRow+1
        kw =  sht.range(index+str(iRow)).value
        if(kw==u'General Info'):
            iwellDefRow = getKeyWordBeginRow(iRow,u'Well Name')
            wellName = sht.range('B'+str(iwellDefRow)).value
            head_I   = sht.range('D'+str(iwellDefRow)).options(numbers=int).value
            head_J   = sht.range('F'+str(iwellDefRow)).options(numbers=int).value
            well = CWell(wellName,head_I,head_J)
            wells[wellName] = well
                
        elif(kw == u'Production'):
            wName = sht.range('A'+str(iRow+1)).value
            iEndProdRow = getKeyWordEndRow(iRow)
            df_prod = sht.range((iRow,2),(iEndProdRow,7)).options(pd.DataFrame,index=False).value
            iRow = iEndProdRow
        
            well = wells.get(wName)
            well.df_prod = df_prod

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

        lMap1      = {'x':x,'y':y1,'id':0,'ylabel':'OPR'}
        lMap2      = {'x':x,'y':y2,'id':1,'ylabel':'WPR'}
        leftMaps   = [lMap1,lMap2]
        rightMaps  = []

        draw2YCurve(ax,wellName,'Date',leftMaps,rightMaps)

        wellId += 1

    plt.show()

    
def readScheduleSection():
    iRow = getKeyWordBeginRow(1,'SCHEDULE SECTION')
    index = 'A'
    wells = {}
    while (sht.range(index+str(iRow)).value!='END'):
        iRow = iRow+1
        kw =  sht.range(index+str(iRow)).value
        if(kw==u'General Info'):
            iwellDefRow = getKeyWordBeginRow(iRow,u'Well Name')
            wellName = sht.range('B'+str(iwellDefRow)).value
            head_I   = sht.range('D'+str(iwellDefRow)).options(numbers=int).value
            head_J   = sht.range('F'+str(iwellDefRow)).options(numbers=int).value
            well = CWell(wellName,head_I,head_J)
            wells[wellName] = well
        
        elif(kw == u'Well Completion'):
            comp = []
            wName = sht.range('B'+str(iRow)).value
            well = wells.get(wName)
            iRow = iRow + 3
            iEndComp = getKeyWordEndRow(iRow)
            for i_layer in range(iRow,iEndComp+1):
                nLayer = sht.range('D'+str(i_layer)).options(numbers=int).value
                comp.append(nLayer)
            
            iRow = iEndComp
            well.comp = comp
        
        elif(kw == u'Production'):
            wName = sht.range('A'+str(iRow+1)).value
            iEndProdRow = getKeyWordEndRow(iRow)
            prod_date = sht.range((iRow+1,2),(iEndProdRow,2)).options(numbers=str).value
            df_prod = sht.range((iRow,3),(iEndProdRow,7)).options(pd.DataFrame,index=False).value
            iRow = iEndProdRow
        
        
            well = wells.get(wName)
            well.df_prod = df_prod
        

            np_prodDate = np.array(prod_date)
            asciiList = [n.encode("ascii", "ignore") for n in np_prodDate]
            well.prod_date = asciiList

    fModel = h5py.File(hdf5File,'a')
    dt = h5py.special_dtype(vlen=str)

    for well in wells.values():
        groupName = "DynamicGroup/Field/"+well.name
        wellentry = fModel.create_group(groupName)
        wellentry.attrs[u'name'] = well.name
        wellentry.attrs[u'id']   = well.id
        wellentry.attrs[u'headI'] = well.head_I
        wellentry.attrs[u'headJ'] = well.head_J
        wellentry.attrs[u'comp']  = well.comp
    #wellentry.attrs[u'prodDate'] = well.prod_date
        df_to_nparray = well.df_prod.to_records(index=True)
    #wellentry.create_dataset(u'prod', data=df_to_nparray)
        fModel[groupName+'/'+'prodData'] = df_to_nparray
        wellentry.create_dataset('date', (len(well.prod_date),), data=well.prod_date,dtype=dt)
    
    fModel.close()          

def outputIOSection(f):
    f.write('TITLE1   GENERATED FROM EXCEL\nTITLE2  CASE  1\n')
    f.write('##################### I/O SECTION #######################\n')
    f.write('IO_CONTROL\nSUMMARY\nPRINTGRID  13*0\n')

def outputDimSection(f,fModel):
    f.write('############# DYNAMIC DIMENSIONING SECTION ##############\n')
    f.write('DIMENSION\n')
    datasetNames = [n for n in fModel.keys()]
    nx = 0
    ny = 0
    nz = 0
    for n in datasetNames:
        if(n=='StaticGroup'):
            staticG = fModel[n]
            modelShape = []
            #staticG.items()
            for item in staticG.attrs.keys():
                modelShape.append(staticG.attrs[item])
        elif(n=='DynamicGroup'):
            dynamicG = fModel[n]
            fieldG   = dynamicG['Field']

            dateArray    = []
            for wellG in fieldG.keys():
                well      = fieldG[wellG]
                wellDate  = well[u'date']
                dateArray.append(wellDate)
            
            date_tuple = tuple(dateArray)
            mergedDate = np.concatenate(date_tuple)
            uniqueDate = np.unique(mergedDate, return_index=True)[0]
            

        
    nx = modelShape[0]
    ny = modelShape[1]
    nz = modelShape[2]
    f.write('NX  {}  NY  {}  NZ  {}\n'.format(nx,ny,nz))

    f.write('NCMAX  3   NWMAX   2\nBLACKOIL\n')

    startDate  = uniqueDate[0]
    intDate = int(float(startDate))
    #formatedDate = datetime.datetime.strptime(str(intDate), '%Y%m').strftime('%Y %m %d')

    firstDate = datetime.datetime.strptime(str(intDate), '%Y%m')
    #delta = datetime.timedelta(days=1)
    prevDate = firstDate #-delta
    f.write('FIRSTDAY  {}\n'.format(prevDate.strftime('%Y %m %d')))

def dumpGridSection(f,fModel):
    f.write('################### GRID SECTION #######################')
    f.write('\nGRID\nCARTESIAN\n')
    datasetNames = [n for n in fModel.keys()]
    for n in datasetNames:
        if(n=='StaticGroup'):
            staticG = fModel[n]
            modelShape = []
            #staticG.items()
            for item in staticG.attrs.keys():
                modelShape.append(staticG.attrs[item])
            
            # keys() is used for iter
            for item in staticG.keys():
                x = staticG[item].value

                
                if(item!='tops'):
                    f.write('{}    XYZALL\n'.format(kwMap[item]))
                    for indexes, value in np.ndenumerate(x):
                        z,y,x = indexes
                        if(x==modelShape[0]-1):
                            f.write('{:.3f}\n'.format(value))
                        else:
                            f.write('{:.3f}    '.format(value))
                else:
                    f.write('TOPS    CON    {}'.format(x[0][0]))

                f.write('\n\n')

        elif(n=='PvtGroup'):
            pvtG = fModel[n]
            fCRock  = pvtG.attrs[u'ROCK']
            f.write('CROCK   CON   {}\n'.format(fCRock))
    

def dumpPvtSection(f,fModel):
    f.write('##################### PVT SECTION #######################\n')
    f.write('PVT\n#\n')
    f.write('COMPONENTS    H2O    HOIL      SGAS\n')
    f.write('#            ----    ----      ----\n')
    f.write('MW         18.015   210.0      22.9674\n')
    f.write('COMPRESS   2.88E-6  1.3687E-5  0\n')
    f.write('CTEXP      0        3.8E-5     0.00038\n')
    f.write('DENSITY    64.79    49.111     0.0605438\n')
    f.write('GRAVITY    1.0383   0.7411     0.7928\n')
    f.write('PHASE_ID   WATER    OIL       GAS\n')

    datasetNames = [n for n in fModel.keys()]
    for n in datasetNames:
        if(n=='PvtGroup'):
            pvtG = fModel[n]
            fPb  = pvtG.attrs[u'PB']
            fBo_s= pvtG.attrs[u'BO_SLOPE']

            pvto = pvtG[u'PVTO']
            a         = pvto[...]
            df_pvto   = pd.DataFrame(a)
            pvtg = pvtG[u'PVTG']
            a         = pvtg[...]
            df_pvtg   = pd.DataFrame(a)
            df_merged = pd.merge(df_pvto,df_pvtg,how='outer',on='P')
            df_merged = df_merged.sort_values(by = 'P',ascending = True)
            df_merged = df_merged.interpolate(method='linear')

            cols = df_merged.columns.tolist()
            cols.insert(1,cols.pop(cols.index('RSO')))
            cols.insert(2,cols.pop(cols.index('BO')))
            cols.insert(3,cols.pop(cols.index('BG')))
            cols.insert(4,cols.pop(cols.index('MUO')))
            cols.insert(5,cols.pop(cols.index('MUG')))

            df_merged = df_merged[cols]
            df_merged = df_merged.rename(columns={df_merged.columns[0]:'PTAB'
                ,df_merged.columns[2]:'BO'
                ,df_merged.columns[3]:'BG'
                ,df_merged.columns[4]:'VIS_OIL'
                ,df_merged.columns[5]:'VIS_GAS'})

            df_merged.to_csv(f,sep=' ',float_format='%.3f',index=False)
            
            f.write('\n#\nPB  {}\n'.format(fPb))
            pvtw = pvtG[u'PVTW']
            a    = pvtw[...]
            df_pvtw = pd.DataFrame(a)
            mx_pvtw = df_pvtw.as_matrix(columns=None)
            mx_pvto = df_merged.as_matrix(columns=None)
            pvtoRows= df_pvto.iloc[:,0].size
            fPusat  = mx_pvto[pvtoRows-1,0]
            #
            maxPress = fPusat+200
            pvtG.attrs[u'MAXPRESS'] = maxPress
            #
            BoPb    = 0
            if(df_pvto[df_pvto[u'P']==fPb].empty):
                #df_prodByDate = df_pvto[df_pvto[u'P'].isin([fPb])]
                df_pvto.loc[df_pvto.shape[0]+1] = {'P':fPb,'MUO':np.nan,'BO':np.nan,'RSO':np.nan}
                df_pvto = df_pvto.interpolate(method='linear')
                BoPb    = df_pvto[df_pvto[u'P']==fPb]['BO']
                mx_BoPb = BoPb.as_matrix(columns=None)
                fBoPb   = mx_BoPb[0]
            else:
                df_pb = df_pvto[df_pvto[u'P'].isin([fPb])]
                BoPb  = df_pb['BO']

                mx_BoPb = BoPb.as_matrix(columns=None)

                fBoPb   = mx_BoPb[0]

            fBusat = (1 + fBo_s * (fPusat - fPb))* fBoPb

            fBW     = mx_pvtw[0,2]
            fRefPw  = mx_pvtw[0,0]
            fVisW   = mx_pvtw[0,1]
            f.write('BWI  {}\nREFERPW  {}\nVIS_WAT  {}\n'.format(fBW,fRefPw,fVisW))
            f.write('CVISW   {}\nCVISO  {}\n'.format(0.0,0.0))
            f.write('PUSAT  {}\nBUSAT  {:.3f}\n'.format(fPusat,fBusat))


            f.write('#WATER-OIL RELATIVE PERMEABILITY TABLE\n')
            f.write('RELPERM    TABLE    1\n')
            f.write('STONE2\n')
            f.write('WATER-OIL\n')
            f.write('#     SW    KRW       KROW      PCWO  (PSIA)\n')

            swof = pvtG[u'SWOF']
            a         = swof[...]
            df_swof   = pd.DataFrame(a)
            df_swof.drop(u'index',axis=1,inplace=True)
            df_swof.to_csv(f,sep=' ',float_format='%.5f',index=False,header=False)

            f.write('#\n#GAS-OIL RELATIVE PERMEABILITY TABLE\n#\n')
            f.write('GAS-LIQUID\n')
            f.write('#SLIQ    KRG     KROG    PCOG (PSIA)\n')
            sglf = pvtG[u'SGLF']
            a         = sglf[...]
            df_sglf   = pd.DataFrame(a)
            df_sglf.drop(u'index',axis=1,inplace=True)
            df_sglf.to_csv(f,sep=' ',float_format='%.5f',index=False,header=False)

def dumpInitSection(f,fModel):
    f.write('############### EQUILLIBRATION SECTION #################\n')
    f.write('EQUIL   MODEL  1\n')
    datasetNames = [n for n in fModel.keys()]
    for n in datasetNames:
        if(n=='InitGroup'):
            initG = fModel[n]
            fRefPress = initG.attrs[u'RefPress']
            fRefDepth = initG.attrs[u'RefDepth']
            fwoc      = initG.attrs[u'OWC']
            fgoc      = initG.attrs[u'OGC']

    f.write('PRESI   CON    {}\n'.format(fRefPress))
    f.write('TRESI   CON    {}\n'.format(100.0))
    f.write('PCGRAVITY\n')
    f.write('DEPTH  {}  DWOC  {}  DGOC  {}  PCWOR  0.  PCGOC  0.\n'.format(fRefDepth,fwoc,fgoc))

def dumpSolutionSection(f,fModel):
    f.write('############### SOLUTION SECTION #######################\n')
    f.write('SOLUTION\n')
    
    datasetNames = [n for n in fModel.keys()]
    for n in datasetNames:
        if(n=='PvtGroup'):
            pvtG = fModel[n]
            maxpress = pvtG.attrs[u'MAXPRESS']
            f.write('MAXPRES  {}\n'.format(maxpress))

    f.write('RELTOL  0.0005\n')

def dumpSchduleSection(f,fModel):
    f.write('############### RECURRENT SECTION #######################\n')
    f.write('RUN\n')

    datasetNames = [n for n in fModel.keys()]
    for n in datasetNames:
        if(n=='DynamicGroup'):
            dynamicG = fModel[n]
            fieldG   = dynamicG['Field']

            wellProd_dfs = []
            dateArray    = []
            for wellG in fieldG.keys():
                well      = fieldG[wellG]
                wellDate  = well[u'date']

                prod      = well[u'prodData']
                a         = prod[...]
                df_prod   = pd.DataFrame(a)
                df_prod[u'wellName']   = well.attrs[u'name']
                df_prod[u'wellId']     = well.attrs[u'id']
                df_prod[u'Date'] = wellDate
                
                df_prod[u'wopr'] = df_prod[u'Monthly Oil Production']/df_prod[u'Production Days']

                wellProd_dfs.append(df_prod)
                dateArray.append(wellDate)
                
            df_mergedProd = pd.concat(wellProd_dfs)
            mergedDate    = df_mergedProd[u'Date']
            uniqueDate    = np.unique(mergedDate,return_index=True)[0]
            #
            #date_tuple = tuple(dateArray)
            #mergedDate = np.concatenate(date_tuple)
            #uniqueDate = np.unique(mergedDate, return_index=True)[0]
            date = uniqueDate[0]

            f.write('WELSPECS\n')

            for wellG in fieldG.keys():
                well     = fieldG[wellG]
                wellName = well.attrs[u'name']
                wellId   = well.attrs[u'id']
                headI    = well.attrs[u'headI']
                headJ    = well.attrs[u'headJ']
                f.write('WELL  {}  \'{}\'  PRODUCER  VERTICAL  {}  {}\n'.format(wellId,wellName,headI,headJ))

            f.write('COMPLETION\n')
            for wellG in fieldG.keys():
                well     = fieldG[wellG]
                wellName = well.attrs[u'name']
                wellId   = well.attrs[u'id']
                comps     = well.attrs[u'comp']
                f.write('WELL  {}  \'{}\'  RW   0.25\n'.format(wellId,wellName))
                for comp in comps:
                    f.write('    WPIV_KXYH  {}   -1.0   4*0\n'.format(comp))

            f.write('\nOPERATION\n')
            df_prodByDate = df_mergedProd[df_mergedProd[u'Date'].isin([date])]
            for index,prodItem in df_prodByDate.iterrows():
                wellName = prodItem[u'wellName']
                wellId   = prodItem[u'wellId']
                wopr     = prodItem[u'wopr']
                f.write('    WELL {}  \'{}\'  BHP   14.7   MAXOIL  {:.3f}\n'.format(wellId,wellName,wopr))

            for i, date in enumerate(uniqueDate):
                if i==0:
                    continue

                intDate = int(float(date))
                formatedDate = datetime.datetime.strptime(str(intDate), '%Y%m').strftime('%Y %m %d')
                f.write('DATE  {}'.format(formatedDate))
                f.write('\nOPERATION\n')
                df_prodByDate = df_mergedProd[df_mergedProd[u'Date'].isin([date])]
                for index,prodItem in df_prodByDate.iterrows():
                    wellName = prodItem[u'wellName']
                    wellId   = prodItem[u'wellId']
                    wopr     = prodItem[u'wopr']
                    f.write('    WELL {}  \'{}\'  BHP   14.7   MAXOIL  {:.3f}\n'.format(wellId,wellName,wopr))

def dumpModel2XXSim():
    f = open(filename,'w')

    fModel = h5py.File(hdf5File,'r+')

    outputIOSection(f)
    outputDimSection(f,fModel)
    dumpGridSection(f,fModel)
    dumpPvtSection(f,fModel)
    dumpInitSection(f,fModel)
    dumpSolutionSection(f,fModel)
    dumpSchduleSection(f,fModel)

    f.write('\nSTOP')
    f.flush()
    f.close()
    fModel.close()


def runXXSim():
    defineGrid()
    readGrid()

    readPropSection()

    readInitSection()

    readScheduleSection()

    dumpModel2XXSim()

