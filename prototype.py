import openpyxl
import tkinter as tk
from tkinter import *
import pandas as pd
import matplotlib.pyplot as plt



global total, numbers, header

global df
global _6X,_6E,_9E,_12E,_15E

_6X = 0


book = openpyxl.load_workbook('write2cell.xlsx')

sheet = book.active

root = tk.Tk()
root.title('Input Data')


df = pd.read_excel (r'C:\Research\machine1.xlsx')
df1 = pd.read_excel(r'C:\Research\MachineOutput.xls', sheet_name = '6x Daily')

header = StringVar()
header.set("LINAC Monitoring Dashboard")

canvas1 = tk.Canvas(root, width = 1200, height = 650, bg = 'lightsteelblue')
canvas1.pack()


#def getMean(df):
    

def printtext():
    #global entry1
    string = entry1.get()
    a = list(entry1.get().split(","))
    print(string)
    sheet.append(a)
    book.save('write2cell.xlsx')

#e = Entry(root)

#e.pack()
#e.focus_set()


### THIS FUNCTION GETS THE AVERAGE CHANGE FOR PREDICTION ###


def getAverage():
    print("Inside getAverage function\n")
    #print(df[-3:])
    
    trialDF = df[::-1] #this iterates bottom to top
    resets = []

    for row in trialDF.index:
        if((df['Reset'][row]) == 1):
            resets.append(row)
            print('Being Reset at index: ', row)

    print('\n',resets)

    rlen = len(resets)

    a = df[(resets[0])+1:]        
    
    nextValue = 0
    currValue = 0
    length = len(df)

    startIndex = resets[rlen-1]
    stopIndex = length - 1  #we change this to different values to test
    print(a.index)
    #print("\nThe length of df is: ",len(df))

    total = 0
    values = 0
    
    for index in a.index:
        
        if(index == stopIndex):
            print("\nIndex is on last element thus breaking now")
            continue
        
        b = df['Value'][index], df['Reset'][index]
        
        currValue = b[_6X]
        nextValue = df['Value'][index+1]
        
        #print('\n',index, "Current Value:", currValue)
        #print("Next value is:",nextValue)
        
        if(b[_6X+1]== 1):
            print('Being reset')
            
        else:
            total = total + (nextValue - currValue)
            values += 1
            hi = 1
    avgChange = total/values
    #print(avgChange)

    predict = df['Value'][stopIndex]

    print('\n',predict)
    checker = avgChange

    days = 0
    
    while (predict < 1.55):

        #checker += checker
        days += 1

        predict += avgChange

    print('\n',days)
    print(predict)

    print("\nTRYING SECOND METHOD NOW\n")

    ### TRYING OUT DIFFERENT START/END VALUES

    secondDF = df[resets[4]+1:resets[3]]
    stopIndex1 = resets[3]

    print(secondDF.index, '\n')
    
    for index1 in secondDF.index:
        
        if(index1 == stopIndex1):
            print("\nIndex is on last element thus breaking now")
            continue
        
        b = df['Value'][index1], df['Reset'][index1]
        
        currValue = b[_6X]
        nextValue = df['Value'][index1+1]
        
        #print('\n',index, "Current Value:", currValue)
        #print("Next value is:",nextValue)
        
        if(b[_6X+1]== 1):
            print('Being reset')
            
        else:
            total = total + (nextValue - currValue)
            values += 1
            hi = 1

    avgChange1 = total/values

    print(avgChange1, '\n')

    
    predict = df['Value'][stopIndex1+1]

    print(predict)
    days = 0
    
    while (predict < 1.3):

        #checker += checker
        days += 1

        predict += avgChange

    print('\n',days)
    print(predict)
    
    #############################################
    ### CALCULATING AVERAGE USING PAST 5 DAYS ###
    #############################################

    
    onedf = df[-5:]

#print(len(df))

    total = 0
    count = 0
    avgChange = 0

    for row in onedf.index:
        value = onedf['Value'][row], df['Reset'][row]
        currValue = value[0]
        if(row == (len(df)-1)):
            #print("last element\n")
            continue

        nextValue = df['Value'][row+1]

    #print("Current Value is: ", currValue)
    #print("Next Value is: ", nextValue)

        if(value[1]==1):
            print("Being Reset\n")

        else:
            total = total + (nextValue - currValue)
            count += 1

    #print("The total is: ",total)
        avgChange = abs(total/count)

        lastValue = df['Value'][len(df)-1]

#print("\nThe last element is:",lastValue)
#print(avgChange)


    months = (1.55 - lastValue)/avgChange

    print(months)

    years = int(months / 12)

    months = months - (years*12)
    months = int(months)

    days = (months - int(months)) * 30

    days = int(days)

    print("\nPrediction: Years =",years,",months =",months, ",days =",days)



##    
##    count = 0
##    total = 0
##    for row in DF1.index:
##        a = DF1['Value'][row],df['Reset'][row]
##        if(a[1] == 1):
##            continue
##        total += a[0]
##        count += 1
##        if (count == 5):
##            break
##
##    print('\n',total/5)
        
        
    
    print("\nExiting getAverage function\n")
    
def getResetMean():
    total = 0
    numbers = 0
    for row in df.index:
        a = df['Value'][row],df['Reset'][row]
        
        if (a[1] == 1):
            total = total + abs(a[0])
            numbers = numbers + 1
            #print(a[0])
    #print(total/numbers)
    return (total/numbers)


def getExcel ():

    print(getResetMean())
    label2 = tk.Label(root, text= "Average reset value is: " + str(getResetMean()) , bg='lightsteelblue')
    canvas1.create_window(800, 600, window=label2)
    df.plot( y='Value', kind = 'line')
    
    ###LABEL AXES

    
    df1.plot( y='Dose', kind = 'line')

    
    ## Add title and axis names
    plt.title('Daily Dose for Pink Machine 6X')
    #plt.xlabel('categories')
    #plt.ylabel('values')
    
    plt.show()


def getCusum(compare):

    ###Add a message when reaching a threshold value

    ###Add a warning

    
    total = 0
    dat = []
    value = 0
    
    for row in df.index:
        a = df['Value'][row], df['Reset'][row]
        #print(a[0])
        
        initial = (compare - abs(a[0]))
        value = abs(initial)        
        
        if(a[1] == 1.0):
            total = 0
            #dat.append(0)
        else:
            total = total + value
            
        dat.append(total)
    
    plt.close('all')    
    plt.plot(dat)
    plt.show()

  

### ENTRY FIELDS ###
 
entry1 = tk.Entry(root, bd=5, width=50) 
canvas1.create_window(350,100, window=entry1)
#entry1.pack()

entry2 = tk.Entry(root, bd=5, width=50)
canvas1.create_window(350,200, window=entry2)

### LABELS ###

label1 = tk.Label(root, text= "Enter dose% value:", bg='lightsteelblue')
canvas1.create_window(120, 100, window=label1)

label3 = tk.Label(root, text= "Enter mean cusum value:", bg='lightsteelblue')
canvas1.create_window(120, 200, window=label3)

label4 = tk.Label(root, textvariable= header, bg='lightsteelblue',justify='center',font=("Helvetica", 24,'bold'))
canvas1.create_window(600, 30, window=label4)


### BUTTONS ###

browseButton_Excel = tk.Button(text='Input', command=printtext, bg='green', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(550, 100, window=browseButton_Excel)

cusumButton = tk.Button(text='Calculate Cusum', command=lambda: getCusum(float(entry2.get())), bg='green', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(595, 200, window=cusumButton)
    
browseButton_Excel = tk.Button(text='Plot Pink Machine data', command=getExcel, bg='green', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(550, 600, window=browseButton_Excel)


get_Average = tk.Button(text='Predict when out of tolerance', command=lambda: getAverage(), bg='green', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(200, 600, window=get_Average)


root.mainloop()
