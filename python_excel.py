import xlsxwriter,xlrd



def write():
    #create workbook and worksheet
    workbook=xlsxwriter.Workbook("Student_result.xlsx")
    sheet=workbook.add_worksheet()

    #data
    global roll_numbers
    global names
    global marks

    roll_numbers=[11,12,13,14,15]
    names=['harshada','mansi','tanishka','isha','archana']
    marks=[90,87,95,92,89]

    #create header
    sheet.write("A1","Roll Number")
    sheet.write("B1","Name")
    sheet.write("C1","Marks")

    #insert all roll numbers
    for roll in range(len(roll_numbers)):
        sheet.write(roll+1,0,roll_numbers[roll])

    #insert all names
    for name in range(len(names)):
        sheet.write(name+1,1,names[name])

    #insert all marks
    for mark in range(len(marks)):
        sheet.write(mark+1,2,marks[mark])    

    workbook.close()

def read():
    #open workbook
    workbook=xlrd.open_workbook("Student_result.xlsx")
    #select sheet by it's index
    sheet=workbook.sheet_by_index(0)#there is only one sheet

    #print total rows and columns 
    total_rows=sheet.nrows
    #total_cols=sheet.ncols

    #read roll numbers
    roll_numbers=[]
    for roll in range(total_rows-1):
        roll_numbers.append(int(sheet.cell_value(roll+1,0)))
    
    print(roll_numbers)

    #read names
    names=[]
    for name in range(total_rows-1):
        names.append(sheet.cell_value(name+1,1))

    print(names)

    #read marks
    marks=[]
    for mark in range(total_rows-1):
        marks.append(int(sheet.cell_value(mark+1,2)))
    
    print(marks)

def choice():
    while True:
        print("1. Write Data ")
        print("2. Read Data")
        option=input("What would you like to do ? :")
        print(option)
        if option=='2':
            read()
        elif option == '1':
            write()
        answer=input("Would you like to go ahead ?(y/n) :")
        if answer.lower() == 'n':
            print("Good Bye")
            return(exit)
choice()