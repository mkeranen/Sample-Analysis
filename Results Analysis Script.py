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
# This class outlines a framework for each sample tested
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

#Data range designations and sample object creation            
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

s6DataRanges = {'layer1A':'B130:B133',
                  'layer1B':'C130:C133',
                  'layer11A':'D130:D133',
                  'layer11B':'E130:E133',
                  'layer21A':'F130:F133',
                  'layer21B':'G130:G133',
                  'layer31A':'H130:H133',
                  'layer31B':'I130:I133',
                  'layer41A':'J130:J133',
                  'layer41B':'K130:K133',
                  'layer51A':'L130:L133',
                  'layer51B':'M130:M133',
                  'layer61A':'N130:N133',
                  'layer61B':'O130:O133',
                  'layer71A':'P130:P133',
                  'layer71B':'Q130:Q133',
                  'layer81A':'R130:R133',
                  'layer81B':'S130:S133',
                  'layer91A':'T130:T133',
                  'layer91B':'U130:U133'}; s6 = sample('0740127', s6DataRanges)

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
#This function calculates and returns the mean of a dataset parameter
def find_mean(dataset):
    return sum(dataset)/float(len(dataset))

#FUNCTION----------------------------------------------------------------------
#This function plots the results of analysis
def plot_data(sample, figNum):
    
    #Calculate the means of all the raw data points, store in list
    listOfMeans = [find_mean(sample.layer1A), find_mean(sample.layer1B),
                   find_mean(sample.layer11A), find_mean(sample.layer11B),
                   find_mean(sample.layer21A), find_mean(sample.layer21B),
                   find_mean(sample.layer31A), find_mean(sample.layer31B),
                   find_mean(sample.layer41A), find_mean(sample.layer41B),
                   find_mean(sample.layer51A), find_mean(sample.layer51B),
                   find_mean(sample.layer61A), find_mean(sample.layer61B),
                   find_mean(sample.layer71A), find_mean(sample.layer71B),
                   find_mean(sample.layer81A), find_mean(sample.layer81B),
                   find_mean(sample.layer91A), find_mean(sample.layer91B)]
    
    #Create new figure for each plot
    plt.figure(figNum)
    #Uncomment for to plot raw data
#    line1A = plt.plot([1]*(len(sample.layer1A)),sample.layer1A, 'ko', markerfacecolor='none')
#    line1B = plt.plot([2]*(len(sample.layer1B)),sample.layer1B, 'ko', markerfacecolor='none')
#    line11A = plt.plot([3]*(len(sample.layer11A)),sample.layer11A, 'ko', markerfacecolor='none')
#    line11B = plt.plot([4]*(len(sample.layer11B)),sample.layer11B, 'ko', markerfacecolor='none')
#    line21A = plt.plot([5]*(len(sample.layer21A)),sample.layer21A, 'ko', markerfacecolor='none')
#    line21B = plt.plot([6]*(len(sample.layer21B)),sample.layer21B, 'ko', markerfacecolor='none')
#    line31A = plt.plot([7]*(len(sample.layer31A)),sample.layer31A, 'ko', markerfacecolor='none')
#    line31B = plt.plot([8]*(len(sample.layer31B)),sample.layer31B, 'ko', markerfacecolor='none')
#    line41A = plt.plot([9]*(len(sample.layer41A)),sample.layer41A, 'ko', markerfacecolor='none')
#    line41B = plt.plot([10]*(len(sample.layer41B)),sample.layer41B, 'ko', markerfacecolor='none')
#    line51A = plt.plot([11]*(len(sample.layer51A)),sample.layer51A, 'ko', markerfacecolor='none')
#    line51B = plt.plot([12]*(len(sample.layer51B)),sample.layer51B, 'ko', markerfacecolor='none')
#    line61A = plt.plot([13]*(len(sample.layer61A)),sample.layer61A, 'ko', markerfacecolor='none')
#    line61B = plt.plot([14]*(len(sample.layer61B)),sample.layer61B, 'ko', markerfacecolor='none')
#    line71A = plt.plot([15]*(len(sample.layer71A)),sample.layer71A, 'ko', markerfacecolor='none')
#    line71B = plt.plot([16]*(len(sample.layer71B)),sample.layer71B, 'ko', markerfacecolor='none')
#    line81A = plt.plot([17]*(len(sample.layer81A)),sample.layer81A, 'ko', markerfacecolor='none')
#    line81B = plt.plot([18]*(len(sample.layer81B)),sample.layer81B, 'ko', markerfacecolor='none')
#    line91A = plt.plot([19]*(len(sample.layer91A)),sample.layer91A, 'ko', markerfacecolor='none')
#    line91B = plt.plot([20]*(len(sample.layer91B)),sample.layer91B, 'ko', markerfacecolor='none')
    
#   Plot the means
    avgLine = plt.plot(range(1,len(listOfMeans)+1), listOfMeans, 'r',linewidth=0.4)
    
    #Overlay boxplots on mean line. Comment out to remove boxplots
    bp = plt.boxplot([sample.layer1A,sample.layer1B,sample.layer11A,sample.layer11B,
                 sample.layer21A,sample.layer21B,sample.layer31A,sample.layer31B,
                 sample.layer41A,sample.layer41B,sample.layer51A,sample.layer51B,
                 sample.layer61A,sample.layer61B,sample.layer71A,sample.layer71B,
                 sample.layer81A,sample.layer81B,sample.layer91A,sample.layer91B],widths=.25)
    
    #Format all linewidths, median line color, and flier marker types
    for box in bp['boxes']:
        box.set(linewidth=.4)
    
    for whisker in bp['whiskers']:
        whisker.set(linewidth=.4)
    
    for median in bp['medians']:
        median.set(color='#000000', linewidth=.4)
        
    for cap in bp['caps']:
        cap.set(linewidth=.4)
        
    for flier in bp['fliers']:
        flier.set(markersize=2,marker='.')
        

    #Add figure labels, limits, titles, etc.
    plt.xlabel('Sampling Layer')
    plt.ylabel('Intensity')
    plt.ylim(800,3700)
    plt.title('Analysis Results for sample ' + str(sample.serialNum))
    plt.xticks(range(1,21),('1A', '1B', '11A', '11B', '21A', '21B', 
                            '31A', '31B', '41A', '41B', '51A', '51B', 
                            '61A', '61B', '71A', '71B', '81A', '81B', 
                            '91A', '91B',))
    plt.xticks(rotation=60)
    plt.tight_layout()
    plt.draw() #draws all plots without waiting for previous one to be dismissed
    plt.savefig(str(sample.serialNum)+'.pdf', bbox_inches='tight')
    
#Function call to plot data for all objects
plot_data(s1,1)
plot_data(s2,2)
plot_data(s3,3)
plot_data(s4,4)
plot_data(s5,5)
plot_data(s6,6)
plot_data(s7,7)
plot_data(s8,8)
plt.show() #keeps plots open until dismissed