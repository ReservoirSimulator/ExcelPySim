from __future__ import division
import sys
import os
import matplotlib.pyplot as plt
import xlwings as xw


sht = xw.Book.caller().sheets[0]
gridVaryOpt={ }
defaultDim = 200
kwMap = {'dx':'DX','dy':'DY','dz':'DZ','por':'POROSITY','permx':'KX','permy':'PERMY',
        'permz':'PERMZ'}
    
    
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

    return (nx,ny,nz)


def reset():

    whiteColor = (255,255,255)
    sht.range((9,3),(9,defaultDim+3-1)).color = whiteColor
    sht.range((10,3),(10,defaultDim+3-1)).color = whiteColor
    sht.range((11,3),(11,defaultDim+3-1)).color = whiteColor
    sht.range((17,3),(17,defaultDim+3-1)).color = whiteColor
    sht.range((18,3),(18,defaultDim+3-1)).color = whiteColor
    sht.range((19,3),(19,defaultDim+3-1)).color = whiteColor
    sht.range((20,3),(20,defaultDim+3-1)).color = whiteColor

    sht.range((9,3),(9,defaultDim+3-1)).value = ""
    sht.range((10,3),(10,defaultDim+3-1)).value = ""
    sht.range((11,3),(11,defaultDim+3-1)).value = ""
    sht.range((17,3),(17,defaultDim+3-1)).value = ""
    sht.range((18,3),(18,defaultDim+3-1)).value = ""
    sht.range((19,3),(19,defaultDim+3-1)).value = ""
    sht.range((20,3),(20,defaultDim+3-1)).value = ""

def DxInterface(gridDimn,backColor):
    if gridVaryOpt['xopt'] == 'CON':
        sht.range('A9').value = 'Grid Size: DX'
        sht.range('B9').value = 'Constant'
        sht.range('C9').color = backColor
    elif gridVaryOpt['xopt'] == 'XVAR':
        sht.range('A9').value = 'Grid Size: DX'
        sht.range('B9').value = 'Variable by X direction'
        sht.range((9,3),(9,gridDimn[0]+3-1)).color = backColor

def DyInterface(gridDimn,backColor):
    if gridVaryOpt['yopt'] == 'CON':
        sht.range('A10').value = 'Grid Size: DY'
        sht.range('B10').value = 'Constant'
        sht.range('C10').color = backColor
    elif gridVaryOpt['yopt'] == 'YVAR':
        sht.range('A10').value = 'Grid Size: DY'
        sht.range('B10').value = 'Variable by Y direction'
        sht.range((10,3),(10,gridDimn[1]+3-1)).color = backColor

def DzInterface(gridDimn,backColor):
    if gridVaryOpt['zopt'] == 'CON':
        sht.range('A11').value = 'Grid Size: DZ'
        sht.range('B11').value = 'Constant'
        sht.range('C11').color = backColor
    elif gridVaryOpt['zopt'] == 'ZVAR':
        sht.range('A11').value = 'Grid Size: DZ'
        sht.range('B11').value = 'Variable by Z direction'
        sht.range((11,3),(11,gridDimn[2]+3-1)).color = backColor

def PoroInterface(gridDimn,backColor):
    if gridVaryOpt['por'] == 'CON':
        sht.range('A17').value = 'Porosity'
        sht.range('B17').value = 'Constant'
        sht.range('C17').color = backColor
    elif gridVaryOpt['por'] == 'ZVAR':
        sht.range('A17').value = 'Porosity'
        sht.range('B17').value = 'Variable by Z direction'
        sht.range((17,3),(17,gridDimn[2]+3-1)).color = backColor

def PermxInterface(gridDimn,backColor):
    if gridVaryOpt['kx'] == 'CON':
        sht.range('A18').value = 'PERMX'
        sht.range('B18').value = 'Constant'
        sht.range('C18').color = backColor
    elif gridVaryOpt['kx'] == 'ZVAR':
        sht.range('A18').value = 'PERMX'
        sht.range('B18').value = 'Variable by Z direction'
        sht.range((18,3),(18,gridDimn[2]+3-1)).color = backColor

def PermyInterface(gridDimn,backColor):
    if gridVaryOpt['ky'] == 'CON':
        sht.range('A19').value = 'PERMY'
        sht.range('B19').value = 'Constant'
        sht.range('C19').color = backColor
    elif gridVaryOpt['ky'] == 'ZVAR':
        sht.range('A19').value = 'PERMY'
        sht.range('B19').value = 'Variable by Z direction'
        sht.range((19,3),(19,gridDimn[2]+3-1)).color = backColor

def PermzInterface(gridDimn,backColor):
    if gridVaryOpt['kz'] == 'CON':
        sht.range('A20').value = 'PERMZ'
        sht.range('B20').value = 'Constant'
        sht.range('C20').color = backColor
    elif gridVaryOpt['kz'] == 'ZVAR':
        sht.range('A20').value = 'PERMZ'
        sht.range('B20').value = 'Variable by Z direction'
        sht.range((20,3),(20,gridDimn[2]+3-1)).color = backColor

def GridStep0():

    greenColor = (102,255,102)

    reset()

    gridDimension = readGeneralInfo()

    DxInterface(gridDimension,greenColor)

    DyInterface(gridDimension,greenColor)
        
    DzInterface(gridDimension,greenColor)

    PoroInterface(gridDimension,greenColor)

    PermxInterface(gridDimension,greenColor)
    
    PermyInterface(gridDimension,greenColor)

    PermzInterface(gridDimension,greenColor)