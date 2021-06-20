import matplotlib.pyplot as plt
from numpy import mean
import numpy as np
import openpyxl
import random

# function to add value labels
def addlabels(y):
    for i in range(len(y)):
        plt.text(i,y[i],y[i],  ha = 'center',
                 Bbox = dict(facecolor = 'red', alpha =.8))

def yearBarChart(years, gpa, colors):
    
    averagedGpa = {}
    sizedGpa = {}
    for k, v in gpa.items():
        year = k
        if year == "200" or year == "201":
            year = "2000"
        averagedGpa[year] = mean(v)
        sizedGpa[year] = len(v)

    # Setup Year-Rating Bar Chart  
    plt.figure(3)

    plt.title("Average 'GPA' Of Different Time Periods")
    plt.xlabel('Time Periods')
    plt.ylabel('Average GPA')
    plt.ylim(min(averagedGpa.values()) - 0.5, 4.0)
    plt.bar(years, averagedGpa.values(), color='royalblue')
    addlabels(list(averagedGpa.values()))

    # Setup Year-NumberAlbumsRated Bar Chart
    plt.figure(4)
    plt.title('Number Of Rated Albums In Different Time Periods')
    plt.xlabel('Time Periods')
    plt.ylabel('Number Of Albums Rated')
    bars = plt.bar(years, sizedGpa.values(), color='royalblue')
    addlabels(list(sizedGpa.values()))

def gradedCharts(valueNames, values, colors):

    # Setup Pie Chart Figure
    plt.figure(1)

    # Explode biggest 1st slice
    explode = [0] * 10
    explode[0] = 0.2
    wedges, texts, autotexts = plt.pie(values, explode=explode, labels=valueNames, colors=colors, 
                                    autopct='%1.0f%%', shadow=True, startangle=140, 
                                    textprops=dict(color="w"))
    # Create legend to right and move off pie with 1-1.5 axes point width
    plt.legend(wedges, valueNames, loc='right', bbox_to_anchor=(1, 0, 0.5, 1))

    # Setup Bar Chart Figure
    plt.figure(2)

    plt.title('Album Ratings And Number Of Albums Rated')
    plt.xlabel('Ratings')
    plt.ylabel('Number Of Albums')
    plt.bar(valueNames, values, color='royalblue')

if __name__ == "__main__":

    xlsx = openpyxl.load_workbook('Vinyl_Collection_201022.xlsx')

    sheet = xlsx.active

    maxRow = sheet.max_row
    sheetValues = sheet['A1' : 'E' + str(maxRow)]

    valueNames = ['Not rated', 'A+', 'A', 'A-', 'B+', 'B', 'B-', 'C', 'D', 'F']
    values = [0] * 10

    years = ["1960s", "1970s", "1980s", "1990s", "2000s"]
    gpa = {
        "196": [],
        "197": [],
        "198": [],
        "199": [],
        "200": [],
        "201": []
    }

    counter = 0
    for c1, c2, c3, c4, c5 in sheetValues:	
        counter += 1
        # stop at end of sheet
        #if c3.value == "OTHER GENRES":
        if c2.value == "VINYL" and c3.value == "SLEEVE":
            print(counter)
            break
        # skip any intermediary rows without data
        elif c2.value == None or c3.value == None:
            continue
        
        try:

            if c1.value == None:
                values[0] += 1
            elif c1.value == "A+":
                values[1] += 1
                gpa[str(c5.value)[0:3]].append(4.0)
            elif c1.value == "A" or c1.value == "A ":
                values[2] += 1
                gpa[str(c5.value)[0:3]].append(4.0)
            elif c1.value == "A-":
                values[3] += 1
                gpa[str(c5.value)[0:3]].append(3.7)
            elif c1.value == "B+":
                values[4] += 1
                gpa[str(c5.value)[0:3]].append(3.3)
            elif c1.value == "B":
                values[5] += 1
                gpa[str(c5.value)[0:3]].append(3.0)
            elif c1.value == "B-":
                values[6] += 1
                gpa[str(c5.value)[0:3]].append(2.7)
            elif c1.value[0] == "C":
                values[7] += 1
                gpa[str(c5.value)[0:3]].append(2.0)
            elif c1.value[0] == "D":
                values[8] += 1
                gpa[str(c5.value)[0:3]].append(1.0)
            elif c1.value[0] == "F":
                values[9] += 1
                gpa[str(c5.value)[0:3]].append(0.0)
            else:
                print("unexpected data c1.value = '" + c1.value + "' at counter")

        except KeyError:
            print("year format exception at line " + str(counter))

    colors = [
        (0,0,0),
        (0,1,0),
        (0,153/255,102/255),
        (0,50/255,0),
        (0,180/255,220/255),
        (0,128/255,1),
        (0,0,1),
        (204/255,204/255,0),
        (1,128/255,0),
        (1,0,0)
    ]

    yearBarChart(years, gpa, colors)
    gradedCharts(valueNames, values, colors)

    plt.show()

