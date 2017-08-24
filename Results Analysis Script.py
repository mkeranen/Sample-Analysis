# -*- coding: utf-8 -*-
"""
Created on Fri May 12 13:34:13 2017

@author: mkeranen
"""
"""
This script analyzes data for each sample.

"""
import matplotlib.pyplot as plt
from openpyxl import load_workbook

fileName = 'Analysis Results.xlsx'

wb = load_workbook(fileName, data_only=True) #data_only to pick data not formulas
sheet1 = wb.get_sheet_by_name('Raw Data')

#FUNCTION----------------------------------------------------------------------
# This function extracts the values from the cells passed into it and returns
# them in a list
def get_raw_data(dataRange):
    
    rangeVals = dataRange.split(':')
    data = sheet1[rangeVals[0] : rangeVals[1]]
    
    dataList = []
    
    for cell in data:
        dataList.append(cell[0].value)
        
    return dataList

#CLASS-------------------------------------------------------------------------
# This class outlines a framework for each oxygenator tested
class sample:
        
    def __init__(self, serialNum, dataRanges):
        self.serialNum = serialNum
        self.layer1A = get_raw_data(dataRanges['layer1A'])
        self.layer1B = get_raw_data(dataRanges['layer1B'])
        self.layer11A = get_raw_data(dataRanges['layer11A'])
        self.layer11B = get_raw_data(dataRanges['layer11B'])
        self.layer21A = get_raw_data(dataRanges['layer21A'])
        self.layer21B = get_raw_data(dataRanges['layer21B'])
        self.layer31A = get_raw_data(dataRanges['layer31A'])
        self.layer31B = get_raw_data(dataRanges['layer31B'])
        self.layer41A = get_raw_data(dataRanges['layer41A'])
        self.layer41B = get_raw_data(dataRanges['layer41B'])
        self.layer51A = get_raw_data(dataRanges['layer51A'])
        self.layer51B = get_raw_data(dataRanges['layer51B'])
        self.layer61A = get_raw_data(dataRanges['layer61A'])
        self.layer61B = get_raw_data(dataRanges['layer61B'])
        self.layer71A = get_raw_data(dataRanges['layer71A'])
        self.layer71B = get_raw_data(dataRanges['layer71B'])
        self.layer81A = get_raw_data(dataRanges['layer81A'])
        self.layer81B = get_raw_data(dataRanges['layer81B'])
        self.layer91A = get_raw_data(dataRanges['layer91A'])
        self.layer91B = get_raw_data(dataRanges['layer91B'])

#Data range designations and Oxygenator object creation            
s1DataRanges = {'layer1A':'B5:B8',
                  'layer1B':'C5:C8',
                  'layer11A':'D5:D8',
                  'layer11B':'E5:E8',
                  'layer21A':'F5:F8',
                  'layer21B':'G5:G8',
                  'layer31A':'H5:H8',
                  'layer31B':'I5:I8',
                  'layer41A':'J5:J8',
                  'layer41B':'K5:K8',
                  'layer51A':'L5:L8',
                  'layer51B':'M5:M8',
                  'layer61A':'N5:N8',
                  'layer61B':'O5:O8',
                  'layer71A':'P5:P8',
                  'layer71B':'Q5:Q8',
                  'layer81A':'R5:R8',
                  'layer81B':'S5:S8',
                  'layer91A':'T5:T8',
                  'layer91B':'U5:U8'}; s1 = sample('0740093', s1DataRanges)

s2DataRanges = {'layer1A':'B30:B33',
                  'layer1B':'C30:C33',
                  'layer11A':'D30:D33',
                  'layer11B':'E30:E33',
                  'layer21A':'F30:F33',
                  'layer21B':'G30:G33',
                  'layer31A':'H30:H33',
                  'layer31B':'I30:I33',
                  'layer41A':'J30:J33',
                  'layer41B':'K30:K33',
                  'layer51A':'L30:L33',
                  'layer51B':'M30:M33',
                  'layer61A':'N30:N33',
                  'layer61B':'O30:O33',
                  'layer71A':'P30:P33',
                  'layer71B':'Q30:Q33',
                  'layer81A':'R30:R33',
                  'layer81B':'S30:S33',
                  'layer91A':'T30:T33',
                  'layer91B':'U30:U33'}; s2 = sample('0740099', s2DataRanges)

s3DataRanges = {'layer1A':'B55:B58',
                  'layer1B':'C55:C58',
                  'layer11A':'D55:D58',
                  'layer11B':'E55:E58',
                  'layer21A':'F55:F58',
                  'layer21B':'G55:G58',
                  'layer31A':'H55:H58',
                  'layer31B':'I55:I58',
                  'layer41A':'J55:J58',
                  'layer41B':'K55:K58',
                  'layer51A':'L55:L58',
                  'layer51B':'M55:M58',
                  'layer61A':'N55:N58',
                  'layer61B':'O55:O58',
                  'layer71A':'P55:P58',
                  'layer71B':'Q55:Q58',
                  'layer81A':'R55:R58',
                  'layer81B':'S55:S58',
                  'layer91A':'T55:T58',
                  'layer91B':'U55:U58'}; s3 = sample('0740104', s3DataRanges)

s4DataRanges = {'layer1A':'B80:B83',
                  'layer1B':'C80:C83',
                  'layer11A':'D80:D83',
                  'layer11B':'E80:E83',
                  'layer21A':'F80:F83',
                  'layer21B':'G80:G83',
                  'layer31A':'H80:H83',
                  'layer31B':'I80:I83',
                  'layer41A':'J80:J83',
                  'layer41B':'K80:K83',
                  'layer51A':'L80:L83',
                  'layer51B':'M80:M83',
                  'layer61A':'N80:N83',
                  'layer61B':'O80:O83',
                  'layer71A':'P80:P83',
                  'layer71B':'Q80:Q83',
                  'layer81A':'R80:R83',
                  'layer81B':'S80:S83',
                  'layer91A':'T80:T83',
                  'layer91B':'U80:U83'}; s4 = sample('0740106', s4DataRanges)

s5DataRanges = {'layer1A':'B105:B108',
                  'layer1B':'C105:C108',
                  'layer11A':'D105:D108',
                  'layer11B':'E105:E108',
                  'layer21A':'F105:F108',
                  'layer21B':'G105:G108',
                  'layer31A':'H105:H108',
                  'layer31B':'I105:I108',
                  'layer41A':'J105:J108',
                  'layer41B':'K105:K108',
                  'layer51A':'L105:L108',
                  'layer51B':'M105:M108',
                  'layer61A':'N105:N108',
                  'layer61B':'O105:O108',
                  'layer71A':'P105:P108',
                  'layer71B':'Q105:Q108',
                  'layer81A':'R105:R108',
                  'layer81B':'S105:S108',
                  'layer91A':'T105:T108',
                  'layer91B':'U105:U108'}; s5 = sample('0740124', s5DataRanges)

s6DataRanges = {'layer1A':'B130:B134',
                  'layer1B':'C130:C134',
                  'layer11A':'D130:D134',
                  'layer11B':'E130:E134',
                  'layer21A':'F130:F134',
                  'layer21B':'G130:G134',
                  'layer31A':'H130:H134',
                  'layer31B':'I130:I134',
                  'layer41A':'J130:J134',
                  'layer41B':'K130:K134',
                  'layer51A':'L130:L134',
                  'layer51B':'M130:M134',
                  'layer61A':'N130:N134',
                  'layer61B':'O130:O134',
                  'layer71A':'P130:P134',
                  'layer71B':'Q130:Q134',
                  'layer81A':'R130:R134',
                  'layer81B':'S130:S134',
                  'layer91A':'T130:T134',
                  'layer91B':'U130:U134'}; s6 = sample('0740127', s6DataRanges)

s7DataRanges = {'layer1A':'B155:B158',
                  'layer1B':'C155:C158',
                  'layer11A':'D155:D158',
                  'layer11B':'E155:E158',
                  'layer21A':'F155:F158',
                  'layer21B':'G155:G158',
                  'layer31A':'H155:H158',
                  'layer31B':'I155:I158',
                  'layer41A':'J155:J158',
                  'layer41B':'K155:K158',
                  'layer51A':'L155:L158',
                  'layer51B':'M155:M158',
                  'layer61A':'N155:N158',
                  'layer61B':'O155:O158',
                  'layer71A':'P155:P158',
                  'layer71B':'Q155:Q158',
                  'layer81A':'R155:R158',
                  'layer81B':'S155:S158',
                  'layer91A':'T155:T158',
                  'layer91B':'U155:U158'}; s7 = sample('0740128', s7DataRanges)

s8DataRanges = {'layer1A':'B180:B183',
                  'layer1B':'C180:C183',
                  'layer11A':'D180:D183',
                  'layer11B':'E180:E183',
                  'layer21A':'F180:F183',
                  'layer21B':'G180:G183',
                  'layer31A':'H180:H183',
                  'layer31B':'I180:I183',
                  'layer41A':'J180:J183',
                  'layer41B':'K180:K183',
                  'layer51A':'L180:L183',
                  'layer51B':'M180:M183',
                  'layer61A':'N180:N183',
                  'layer61B':'O180:O183',
                  'layer71A':'P180:P183',
                  'layer71B':'Q180:Q183',
                  'layer81A':'R180:R183',
                  'layer81B':'S180:S183',
                  'layer91A':'T180:T183',
                  'layer91B':'U180:U183'}; s8 = sample('0740129', s8DataRanges)

#FUNCTION----------------------------------------------------------------------
#This function plots the results of FM analysis
x = (1,1,1,1)
def plot_FM_analysis_data(l1A):
    mean1A = sum(l1A)/float(len(l1A))
    plt.plot(x,l1A,'b.')
    plt.plot(1, mean1A, 'r.')
    plt.show()

#Function call to plot data
plot_FM_analysis_data(s1.layer1A)