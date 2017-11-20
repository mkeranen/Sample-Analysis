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
import scipy.stats as stats

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
# This class outlines a framework for each sample tested.
# Since it has no methods, probably better off as a dict.
class Sample:
        
    def __init__(self, serialNum, conditioning, dataRanges):
        self.serialNum = serialNum
        self.conditioning = conditioning
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
        listOfRawDataLists = [self.layer1A, self.layer1B, self.layer11A, self.layer11B,
                              self.layer21A, self.layer21B, self.layer31A, self.layer31B,
                              self.layer41A, self.layer41B, self.layer51A, self.layer51B,
                              self.layer61A, self.layer61B, self.layer71A, self.layer71B,
                              self.layer81A, self.layer81B, self.layer91A, self.layer91B]
        
        #Flatten list of lists
        self.allData = [item for sublist in listOfRawDataLists for item in sublist]

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
                  'layer91B':'U5:U8'}; s1 = Sample('0740093', 'SU', s1DataRanges)

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
                  'layer91B':'U30:U33'}; s2 = Sample('0740099', 'None', s2DataRanges)

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
                  'layer91B':'U55:U58'}; s3 = Sample('0740104', 'ETO,TC,SU', s3DataRanges)

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
                  'layer91B':'U80:U83'}; s4 = Sample('0740106', 'ETO,TC,SU', s4DataRanges)

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
                  'layer91B':'U105:U108'}; s5 = Sample('0740124', 'ETO,TC', s5DataRanges)

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
                  'layer91B':'U130:U133'}; s6 = Sample('0740127', 'None', s6DataRanges)

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
                  'layer91B':'U155:U158'}; s7 = Sample('0740128', 'SU', s7DataRanges)

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
                  'layer91B':'U180:U183'}; s8 = Sample('0740129', 'ETO,TC', s8DataRanges)

#FUNCTION----------------------------------------------------------------------
#This function calculates and returns the mean of a dataset parameter
def find_mean(dataset):
    return sum(dataset)/float(len(dataset))

#FUNCTION----------------------------------------------------------------------
#This function plots the results of analysis and returns a list of mean values
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
    
#   Plot the means
    plt.plot(range(1,len(listOfMeans)+1), listOfMeans, 'r',linewidth=0.4)
    
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
    plt.title('Analysis Results for sample ' + str(sample.serialNum) + ' ' + str(sample.conditioning))
    plt.xticks(range(1,21),('1A', '1B', '11A', '11B', '21A', '21B', 
                            '31A', '31B', '41A', '41B', '51A', '51B', 
                            '61A', '61B', '71A', '71B', '81A', '81B', 
                            '91A', '91B',))
    plt.xticks(rotation=60)
    plt.tight_layout()
    plt.draw() #draws all plots without waiting for previous one to be dismissed
    plt.savefig(str(sample.serialNum)+'.pdf', bbox_inches='tight')
    
    #List of mean intensities at each test point for the sample
    return listOfMeans
#------------------------------------------------------------------------------
#Function call to plot data for all objects. This plots the boxes and returns the mean intensity for all layers.
s1meanList = plot_data(s1,1)
s2meanList = plot_data(s2,2)
s3meanList = plot_data(s3,3)
s4meanList = plot_data(s4,4)
s5meanList = plot_data(s5,5)
s6meanList = plot_data(s6,6)
s7meanList = plot_data(s7,7)
s8meanList = plot_data(s8,8)

#Perform ANOVA to check if any difference between samples is detected
anovaResult = stats.f_oneway(s2.allData, s6.allData, s1.allData, s7.allData,
               s5.allData, s8.allData, s3.allData, s4.allData)

#2 Sample T-Test to check for differences between samples of the same type
noneTT = stats.ttest_ind(s2.allData, s6.allData)
suTT = stats.ttest_ind(s1.allData, s7.allData)
etotcTT = stats.ttest_ind(s5.allData, s8.allData)
etotcsuTT = stats.ttest_ind(s3.allData, s4.allData)

#Test for equal variances between all samples
eqVarTest = stats.levene(s2.allData, s6.allData, s1.allData, s7.allData,
               s5.allData, s8.allData, s3.allData, s4.allData)

#Create and format new figure for plotting the mean intensities against each other
plt.figure()
plt.xlabel('Sampling Layer')
plt.ylabel('Intensity')
plt.ylim(800,3700)
plt.title('Means of all Samples')
plt.xticks(range(1,21),('1A', '1B', '11A', '11B', '21A', '21B', 
                            '31A', '31B', '41A', '41B', '51A', '51B', 
                            '61A', '61B', '71A', '71B', '81A', '81B', 
                            '91A', '91B',))
plt.xticks(rotation=60)

#These plot commands plot all the mean intensities on a single figure
plt.plot((range(1,len(s2meanList)+1)), s2meanList, 'r-', label=(str(s2.serialNum) + ' ' + str(s2.conditioning)),linewidth=0.8)
plt.plot((range(1,len(s6meanList)+1)), s6meanList, 'r-', label=(str(s6.serialNum) + ' ' + str(s6.conditioning)),linewidth=0.8)
plt.plot((range(1,len(s1meanList)+1)), s1meanList, 'b-', label=(str(s1.serialNum) + ' ' + str(s1.conditioning)),linewidth=0.8)
plt.plot((range(1,len(s7meanList)+1)), s7meanList, 'b-', label=(str(s7.serialNum) + ' ' + str(s7.conditioning)),linewidth=0.8)
plt.plot((range(1,len(s5meanList)+1)), s5meanList, 'k-', label=(str(s5.serialNum) + ' ' + str(s5.conditioning)),linewidth=0.8)
plt.plot((range(1,len(s8meanList)+1)), s8meanList, 'k-', label=(str(s8.serialNum) + ' ' + str(s8.conditioning)),linewidth=0.8)
plt.plot((range(1,len(s3meanList)+1)), s3meanList, 'g-', label=(str(s3.serialNum) + ' ' + str(s3.conditioning)),linewidth=0.8)
plt.plot((range(1,len(s4meanList)+1)), s4meanList, 'g-', label=(str(s4.serialNum) + ' ' + str(s4.conditioning)),linewidth=0.8)

#Add ANOVA results to figure
plt.text(15, 1200, 'F - value: ' + str(format(anovaResult[0],'.4f')))
plt.text(15, 1000, 'P - value: ' + str(format(anovaResult[1],'.4f')))

plt.legend(bbox_to_anchor=(1, 1.02))
plt.tight_layout()
plt.draw()
plt.savefig('All Sample Means.pdf', bbox_inches='tight')

#Boxplots to compare raw data of each sample
plt.figure()
plt.boxplot([s2.allData, s6.allData, s1.allData, s7.allData,
               s5.allData, s8.allData, s3.allData, s4.allData])

plt.xlabel('Sample S/N')
plt.ylabel('Intensity')
plt.title('Mean intensity for individual test samples')
plt.xticks(range(1,9),((str(s2.serialNum)+' '+str(s2.conditioning)),
                 (str(s6.serialNum)+' '+str(s6.conditioning)),
                 (str(s1.serialNum)+' '+str(s1.conditioning)),
                 (str(s7.serialNum)+' '+str(s7.conditioning)),
                 (str(s5.serialNum)+' '+str(s5.conditioning)),
                 (str(s8.serialNum)+' '+str(s8.conditioning)),
                 (str(s3.serialNum)+' '+str(s3.conditioning)),
                 (str(s4.serialNum)+' '+str(s4.conditioning))))
plt.xticks(rotation=60)

#Add ANOVA results to boxplot figure
plt.text(6.3, 1300, 'F - value: ' + str(format(anovaResult[0],'.4f')))
plt.text(6.3, 1000, 'P - value: ' + str(format(anovaResult[1],'.4f')))

#Add t-test results to boxplot figure
plt.text(1.1, 1500, 'p=' + format(noneTT[1],'.3f'), fontsize=8, color='red')
plt.text(3.1, 1500, 'p=' + format(suTT[1],'.3f'), fontsize=8, color='red')
plt.text(5.1, 1500, 'p=' + format(etotcTT[1],'.3f'), fontsize=8, color='red')
plt.text(7.1, 2000, 'p=' + format(etotcsuTT[1],'.3f'), fontsize=8, color='red')

plt.tight_layout()
plt.draw()
plt.savefig('Boxplots of individual sample raw data.pdf', bbox_inches='tight')

#Print out statistical test information
print('ANOVA Results: ', 'F-value: ', anovaResult[0], ' P-value: ', anovaResult[1])
print('\nT-test P-values: \n')
print('None: ', noneTT[1], 'SU: ', suTT[1], 
      'ETO,TC: ', etotcTT[1], 'ETO,TC,SU: ', etotcsuTT[1])
print('\nLevene Results: ', 'P=', eqVarTest[1])

plt.show() #keeps plots open until dismissed
