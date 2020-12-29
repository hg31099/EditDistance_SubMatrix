# Hardik Gandhi
# BT17CSE030

# Library import for reading excel sheet
import openpyxl

# Reading excel sheet
path='./matrix.xlsx'
workbook=openpyxl.load_workbook(path)
sheet=workbook.active

# input of correct and incorrect string
incorrectString="enformatica"
correctString="information"
incLen=len(incorrectString)
cLen=len(correctString)

# reading the substitution matrix
mat=[[] for i in range(26)]
for i in range(1,27):
    for j in range(1,27):
        cell = sheet.cell(row=i,column=j)
        mat[i-1].append(cell.value)

# declaring 2-d array between incorrect and correct
editDist=[[] for i in range(incLen+1)]

# intialising first row and coloumn
for i in range(0,incLen+1):
    editDist[i].append(int(i))

for i in range(1,cLen+1):
    editDist[0].append(int(i))

# calculating the minimum cost
for i in range (1,incLen+1):
    for j in range(1,cLen+1):
        cost=min(editDist[i-1][j]+1,editDist[i][j-1]+1,editDist[i-1][j-1]+mat[ord(incorrectString[i-1])-97][ord(correctString[j-1])-97])
        editDist[i].append(cost)

# printing
print("Edit distance between strings is: "+ str(editDist[incLen][cLen]))
print("Edit Operations performed are: ")

# row and coloumn for backtracking
r=incLen
c=cLen
count=1

# backtracking
while(r!=0 and c!=0):
    insert=editDist[r][c-1]+1
    delete=editDist[r-1][c]+1
    if(editDist[r][c]==insert):
        print(" Insert " + correctString[c-1]+ ":" + " Insert " + correctString[c-1]+" Cost is: " + str(1))
        c=c-1
    elif editDist[r][c]==delete:
        print(" Delete " + incorrectString[r-1] + ":" + " Delete " + incorrectString[r-1]+" Cost is: " + str(1))
        r=r-1
    else:
        print(" Change " + incorrectString[r-1] +" to "+ correctString[c-1]+ ":" + " Substitute " + correctString[c-1]+" Cost is:" +  str(mat[ord(incorrectString[r-1])-97][ord(correctString[c-1])-97]))
        c=c-1
        r=r-1
    count+=1
while(r!=0):
    print(" Delete " + incorrectString[r-1] + ":" + " Delete " + incorrectString[r-1]+" Cost is: " + str(1))
    r=r-1
while(c!=0):
    print(" Insert " + correctString[c-1]+ ":" + " Insert " + correctString[c-1]+" Cost is: " + str(1))
    c=c-1
