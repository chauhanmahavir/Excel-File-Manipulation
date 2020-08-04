import xlrd
import os
import xlwt

def search():
    search = float(input("\nEnter Enrollment number to find"))
    found = 0

    #x=sheet1.nrows

    for i in range(len(disc['En'])):
        if str(search) == disc['En'][i]:
            found = i
            break
    for i in disc:
        print(disc[i][found])

def add():
    enroll = float(input("Enter Enrollment number"))
    name = input("Enter Name of The student")
    contact = input("Enter Contact number")
    address = input("Enter address")
    gender = input("Enter Gender")
    marks = input("Enter marks")

    l2 = [enroll,name,contact,address,gender,marks]

    import xlwt
    import xlrd
    r=0
    wb = xlwt.Workbook()
    sheet1 = wb.add_sheet('sheet 1')
    for i in disc:
        c=0
        for j in range(len(disc[i])):
            #print(disc[i][j])
            sheet1.write(c,r,disc[i][j])
            c = c+1
        r = r+1
    for i in range(r):
        sheet1.write(c,i,l2[i])
    wb.save('1.xls')

    disc.clear()
    
    #disc1={}
    loc = ("1.xls")

    wb = xlrd.open_workbook(loc)

    sheet = wb.sheet_by_index(0)

    sheet.cell_value(0,0)
    
    for i in range(sheet.ncols):
        for j in range(sheet.nrows):
            #print(sheet.cell_value(0,i),sheet.cell_value(j,i))
            if j != len(range(sheet.nrows))-1:
                l.append(str(sheet.cell_value(j,i)))
            else:
                l.append(str(sheet.cell_value(j,i))) 
                disc[sheet.cell_value(0,i)]=list(l)
                while len(l)>0:
                   l.pop()
    
    
def remove():
    r=0
    wb = xlwt.Workbook()
    sheet1 = wb.add_sheet('sheet 1')
    
    re = float(input("Enter Enrollment Number"))
    for i,v in enumerate(disc[sheet.cell_value(0,0)]):
        if str(re) == v:
            print("Your Element is removed")
            re = i
     
    for j in disc:
        del (disc[j][re])
    
    for i in disc:
        c=0
        for j in range(len(disc[i])):
            #print(disc[i][j])
            sheet1.write(c,r,disc[i][j])
            c = c+1
        r = r+1
    wb.save('1.xls')
    
def update():
     enroll = float(input("Enter Enrollemnt number"))   
     choice1 = 0 
    
     while(choice1 < 6):
         print("\n1)Name")
         print("2)Contact")
         print("3)Address")
         print("4)Gender")
         print("5)Marks")
         print("6)Exit from Update")
         choice1 = int(input("Enter update value"))

         if choice1 == 1:
             c = sheet.cell_value(0,1)
             up = input("Enter Name")
         elif choice1 == 2:
             c = sheet.cell_value(0,2)
             up = input("Enter Contact")
         elif choice1 == 3:
             c = sheet.cell_value(0,3)
             up = input("Enter Address")
         elif choice1 == 4:
             c = sheet.cell_value(0,4)
             up = input("Enter Gender")
         elif choice1 == 5:
             c = sheet.cell_value(0,5)
             up = input("Enter Marks")
         else:
             print("\nThank You\n")
         #print(c,up)

         for i,v in enumerate(disc[sheet.cell_value(0,0)]):
             if v == str(enroll):
                 print(i)
                 break;
         for j in disc:
             if j == str(c):
                 disc[j][i] = up
                     
         #print(disc)
         r=0
         wb = xlwt.Workbook()
         sheet1 = wb.add_sheet('sheet 1')
         for i in disc:
            c=0
            for j in range(len(disc[i])):
            #print(disc[i][j])
                sheet1.write(c,r,disc[i][j])
                c = c+1
            r = r+1
         wb.save('1.xls')   
    

disc={}
if os.path.exists('1.xls') == True:
    loc = ("1.xls")
else:
    loc = ("1.xlsx")

wb = xlrd.open_workbook(loc)

sheet = wb.sheet_by_index(0)

sheet.cell_value(0,0)

l=[]

for i in range(sheet.ncols):
    for j in range(sheet.nrows):
        #print(sheet.cell_value(0,i),sheet.cell_value(j,i))
        if j != len(range(sheet.nrows))-1:
                l.append(str(sheet.cell_value(j,i)))
        else:
               l.append(str(sheet.cell_value(j,i))) 
               disc[sheet.cell_value(0,i)]=list(l)
               while len(l)>0:
                   l.pop()
#print(disc)

choice  = 0
while choice < 5:
    print("\n1) Search")
    print("2) Add")
    print("3) Remove")
    print("4) Update")
    print("5) Exit")
    choice = int(input("\nEnter Choice"))


    if choice == 1:
        search()
    elif choice == 2:
        add()
    elif choice == 3:
        remove()
    elif choice == 4:
        update()
    else:
        print("\nTHANK YOU")

