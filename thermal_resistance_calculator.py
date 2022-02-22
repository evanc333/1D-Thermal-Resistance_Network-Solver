import re
import datetime
import sys
from sympy.solvers import solve
from sympy.solvers import solveset
from sympy import Symbol,simplify
import csv
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from subprocess import Popen

#Various  separator lines below to cleanup code#################################################################################################################################################################################
def separator():
    print('---------------------------------------------')
    return('---------------------------------------------')
def finalsep():
    print('*********************************************')
    return('*********************************************')
def newrunsep():
    print('#############################################################################')
    return('#############################################################################')

#Called by any function#################################################################################################################################################################################

def do(inp):
    print(inp)
    output.append(inp+'\n\n')
def helpf():
    filen = 'Thermal_Resistance_Calculator_Help_File.txt'
    try:
        file = open('Thermal_Resistance_Calculator_Help_File.txt', "w")
        file.write("********************************************\n\n")
        file.write("Thermal Resistance Calculator Help File\n\n")
        file.write("********************************************\n\n\n")
        file.write("----------------------------Introduction--------------------------\n\n")
        file.write("The objective of this program is: \nA) To allow for fast  setup and calcualtion of thermal resitance networks of all shapes and sizes and\n")
        file.write('B) To allow modiciation and recalculation of these changes in real time \n\n')
        file.write("---------------------------Basic Operation-------------------------\n\n")
        file.write('The program really only does one thing, which is to solve Q=(T2-T1)/R and the sub components of R.\n')
        file.write('The operation of the program is fairly straight forward. All known values are entered into a setup file in the form of a csv. \n')
        file.write('The unknown value (rarely values) are replaced with an x. Additionally a thermal resistance network is entered using the following two rules:\n\n')
        file.write('1) Everything that is in series with something else must be separated by a "-" (without the quotes)\n')
        file.write('2) When values are in parallel they must be separated by a "," and bounded by "<" and ">"\n\n')
        file.write("------------------------------Examples-----------------------------\n\n")
        file.write('Based on those rules, a resistor network of three resistors in series would look something like this:\n\n')
        file.write('                      -----r1-----r2-----r3-----\n\n')
        file.write('And would be entered into the setup file as: r1-r2-r3\n\n')
        file.write('A resistor network of three resistors in parallel would look something like this:\n\n')
        file.write('                       |-----r1-----|\n')
        file.write('                    ---|-----r2-----|-----\n')
        file.write('                       |-----r3-----|\n\n')
        file.write('And would be entered into the setup file as: <r1,r2,r3>\n\n')
        file.write('Here is and example of a complex, but realistic, network:\n\n')
        file.write('                                 |---r6----r7----r8----|\n')
        file.write('         |---r2---r3---|         |                     |\n')
        file.write('  ---r1--|             |---r5----|     |---r9---|      |----r11----\n')
        file.write('         |------r4-----|         |-----|        |------|\n')
        file.write('                                       |---r10--|\n\n')
        file.write('And this would be entered as: r1-<r2-r3,r4>-r5-<r6-r7-r8,<r9,r10>>-r11\n\n')
        file.write('One other weird notation note is that two parallel resistor chains in series would be notated as like this:\n')
        file.write('<r1,r2>-<r3,r4>\n\n')
        file.write("------------------------Notes about CSV Setup File----------------\n\n")
        file.write('The csv setup was used to make input as flexible as possible while still being able to retain information for later retreival.\n')
        file.write('You can use the balnk space to the right of any of the columns as a "scratch space". I find this particularly helpful\n')
        file.write('to use when you are converting between imperial and metric. I will paste my list of values  in inches/inches^2 to the right of the existing table.\n')
        file.write('and use the excel function =CONVERT(x,"in2","m2"). Even if you have a formula in place of a variable, Excel turns it into a numeric value\n')
        file.write('when it saves in the csv format.\n\n')
        file.write('-------------------------Running List of Issues-------------------\n\n')
        file.close()
    except FileExistsError:
        pass
    p = Popen(filen, shell=True)
def initvari():
    fnlong = ''
def blankvar():             #blanks and initializes the variables
    global ml,l,k,a,v,xc,xt,typekey,finalpos,output,fnlong
    output = []
    ml = ["r1","r2","r3","r4","r5","r6","r7","r8",
          "r9","r10","r11","r12","r13","r14","r15",
          "r16", "r17", "r18", "r19", "r20", "r21",
          "r22","r23","r24","r25","r26","r27","r28",
          "r29","r30","r31","r32","r33","r34","r35",
          "r36","r37","r38","r39","r40","r41","r42"]
    l = []
    k = []
    a = []
    v = {'r1':0,'r2':0,'r3':0,'r4':0,'r5':0,'r6':0,'r7':0,'r8':0,
         'r9':0,'r10':0, 'r11':0, 'r12':0,'r13':0,'r14':0,'r15':0,
         'r16':0,'r17':0, 'r18':0, 'r19':0,'r20':0,'r21':0,'r22':0,
         'r23':0,'r24':0, 'r25':0, 'r26':0,'r27':0,'r28':0,'r29':0,
         'r30':0,'r31':0, 'r32':0, 'r33':0,'r34':0,'r35':0,'r36':0,
         'r37':0,'r38':0, 'r39':0, 'r40':0,'r41':0,'r42':0}
    xc = []
    xt = []
    typekey = 0 #determines what is solved for in the end |1=k|2=l|3=a|8=T1|9=T2|10=Power|0=Uknown|
    finalpos = 0

#Main Script Below#########################################################################################################################################################################################

def start():
    global filename,output,fnlong
    initvari()
    blankvar()
    output.append('File Generated on: {:%Y-%m-%d %H:%M:%S}'.format(datetime.datetime.now()))
    startin = input("Please make a selection: \n    1: Load Existing \
Setup File \n    2: Create New setup file \n    3: Help \nEntry:  ")
    if startin == "1":
        print("File Dialog Box Opening...")
        root = Tk()
        root.withdraw()
        root.deiconify()
        root.lift()
        root.focus_force()
        filename = askopenfilename(parent=root)
        root.destroy()
        fn = filename.split('.')
        fnlong = str(fn[0])
        fn2 = filename.rsplit('/',1)
        fnnopath = fn2[1]
        printme = "File Chosen: " + str(filename)
        do(printme)
        #try:
        readfile(filename)
        #except:
        print("Something has gone wrong.")
        runagain(0)
    elif startin =="2":
        fileopen = input("Please Enter Name for new CSV:\n")
        filename = fileopen+ '.csv'
        fnnopath = filename
        fnonlypath = ''
        printme = "File Created:" + str(filename)
        do(printme)
        csvgen()
        p = Popen(filename, shell=True)
        fnlong=filename
        cont = input("Fill out the file and press 1 to continue. Or press 2 to exit\nEntry: ")
        if cont == '1':
            blankvar()
            readfile(filename)
        elif cont == '2':
            sys.exit()
    elif startin == "3":
        helpf()
    else:
        print("That wasn't right! Try again \n \n")
        start()

#######################################################################################################################################################################################################

def csvgen():
    global filename
    with open(filename, 'w', newline='') as csvfile:
        spamwriter = csv.writer(csvfile, delimiter=',',
                                quotechar='"', quoting=csv.QUOTE_MINIMAL)
        spamwriter.writerow(['Power:'])
        spamwriter.writerow(['T1:'])
        spamwriter.writerow(['T2:'])
        spamwriter.writerow(['Resistance Formula:'])
        spamwriter.writerow([''])
        spamwriter.writerow(
            ['Designator (Must be Seqential Add/Delete as needed)', 'Name', 'R (C/W) (Optional)', 'k (W/m*C)',
             'Length (m)', 'Area(m^2)'])
        spamwriter.writerow(['r1'])
        spamwriter.writerow(['r2'])
        spamwriter.writerow(['r3'])

#########################################################################################################################################################################################################

def readfile(filename):                                                         #opens the file and fills variable lists with data
    fn = filename
    global xc,xt,typekey,v,l,k,a,output
    with open(fn, newline='') as csvfile:                                       #csv reader copy and pasted text
        spamreader = csv.reader(csvfile, delimiter=',', quotechar='"')          #excel saves items with a ',' in them wrapped in '"' to ensure they don't get split.
        rwc = 0                                                                 #count loops/rows
        for row in spamreader:
            if rwc == 0:                                                        #if this is the first row
                try:                                                            #attampts to turn the power into a number else saves as a string
                    q = float(row[1])
                except:
                    q = str(row[1])
                    typekey = 10                                               #sets typekey to solve for power
            if rwc == 1:                                                        #same as above but for second row
                try:
                    t1 = float(row[1])
                except:
                    t1 = str(row[1])
                    typekey = 8                                                 #sets
            if rwc == 2:                                                        #same as above but for third row
                try:
                    t2 = float(row[1])
                except:
                    t2 = str(row[1])
                    typekey = 9
            if rwc == 3:                                                        #this is where it grabs the actualy resistance network
                x = row[1]                                                      #this is the network
                separator()
                printme = "Raw Network Input:\n" + x                     #prints it out so the user can see it
                do(printme)
                separator()
                xt = re.split('(<|-|,|>)',x)                                    #splits the text on everything but the restiances
                #print(xt)
            if rwc >= 6:                                                        #same as above but for all rows afterward (should be resistors)
                if row[2] != '':                                                #checks to see if 3rd column is empty
                    v[row[0]] = str(row[2])                                     #if its not empty it grabs that resistance value
                if row[0] != '':                                                #checks to see if the second column in the row has an rx if so proceed. (This allows user to use the area below for sratch work)
                    kt = row[3]                                                 #sets the variables based on their columns
                    k.append(kt)
                    lt = row[4]
                    l.append(lt)
                    at = row[5]
                    a.append(at)
            rwc = rwc+1                                                         #next row
    prepare()                                                                   #starts prepare to process this stuff
    runcalc(xc,t2,t1,q)                                                         #runs calcs with prepared data

#######################################################################################################################################################################################################

def prepare():                                                      #calcualtes Rx values and determines how to proceed
    global v,k,a,l,ml,xt,xc,typekey,finalpos,output
    dummy = 0
    rlen = len(l)                                                   #grabs the length of l to see how long all of the resistor component lists are
    r = 0
    while r < rlen:                                                 #Goes through whole list of l and tries to solves
        if v[ml[r]] == 0:                                           #if this is true it means there was no R value on the sheet
            try:
                v[ml[r]] = float(l[r])/(float(k[r])*float(a[r]))    #attempts to math this, non numeric values will cause and error
            except ValueError:                                      #I only want to catch the above error and still break if someone does something stupid
                if a[r] == '' and k[r] == '' and l[r] == '':        #if theyre all blank, blank the whole variable. I think this is old
                    v[ml[r]] = ''
                else:                                               #set the value to x and then try to float each component to determine which one is non-numeric and set the typekey for final solve
                    v[ml[r]] = 'x'
                    try:
                        dummy = float(l[r])
                    except:
                        typekey = 2
                    try:
                        dummy= float(k[r])
                    except:
                        typekey = 1
                    try:
                        dummy = float(a[r])
                    except:
                        typekey = 3
                finalpos = r                                        #!!!!!!!!!!!!!!!(DOUBLE CHECK THIS)sets the finalpostition for later use
            except:
                runagain(0)
        r=r+1
    cn = 0
    for tempo in xt:                                                #this goes through the master list and replaces whatever R value is 'x' to actually be 'x' in the raw resistor network.
        if tempo in v:
            if v[tempo] == 'x':
                xt[cn] = (v[tempo])
        cn = cn+1
    xc = xt                                                         #sets global xc for use in run to the new proper one we just made

#######################################################################################################################################################################################################

def runcalc(xv,t2,t1,q):                                                    #main function that actually gets the math done and gives results
    global typekey,finalpos,v,output,fnnopath,fnonlypath,fnlong,filename             #calls globals
    strip(xv)                                                               #runs the possibly useless strip function
    d=0
    while d < 10:                                                           #runs this 10x which is hopefully enough times to simplify and solve
        parallel(xv)                                                        #calls parallel
        series(xv)                                                          #calls series
        d=d+1
    testsolve = numerify(xv)                                                #calls numerify, so now we have a useful equation
    separator()
    printme = "Substituting Variables into Equation:"                       #this print a list of resistor values so errors can be caught by the user
    do(printme)
    for key, value in v.items():                                            #loops through dictionary
        if value != '' and value != 0:                                      #if the value exists
            print(key, ' : ', value)                                        #prints the key and value
            output.append('\n' + str(key) + ' : ' + str(value))
    printme = "\n\nEquation:\n"+testsolve
    do(printme)
    separator()
    try:                                                                    #this is attempting to solve the left side temperature component of the conductioin quation
        dt=t2-t1                                                            #delta t, will fail if both aren't numbers
    except:                                                                 #if this fails we know that one of the tempeatures was 'x'
        dt = "(" + str(t2) + '-' + str(t1) + ")"                            #so we string it together for later math
    try:                                                                    #tries to solve the rest of the conduction equation
        testsolve2 = str(dt/q)                                              #tries to do math again, will fail if both aren't numbers
    except:                                                                 #when it fails we know that one was 'x'
        testsolve2 = "("+str(dt) + "/" + str(q)+")"                         #string them together to deal with later
    vari = Symbol('x')                                                      #from the sympy package, lets us solve for x
    testsolve = testsolve + "-" + testsolve2                                #moves left side of cond eq to the right side. Sympy needs an equation that equals 0
    separator()
    printme = "Simplified Equation to: \n" + str(simplify(testsolve)) + "=0"   #prints whole formula that is equal to 0
    do(printme)
    separator()
    result = solve(testsolve, vari)                                         #this actually solves the function above for 'x' and result is our answer
    if typekey == 0:                                                        #If some unknon error happens
        printme = "Unknown Result: " + str(result[0])                        #unknown results but still gives answer
        do(printme)
    if typekey == 1 or typekey == 2 or typekey == 3:                        #these typekeys correspond to a resistor being 'x'
        separator()
        printme = "Resistor Value Found:" + str(result[0]) + " degC/W"                   #gives resistor value
        do(printme)
        separator()
        finalsep()
        printme = "Final Results:"
        do(printme)
        finalsolve  = str(result[0]) + "-(" + str(l[finalpos])+ \
            "/(" + str(k[finalpos]) + "*" + str(a[finalpos]) + "))"         #we know the resitor value, but often we have at least two of the R=L/KA values, so we try to solve for this
        if typekey == 1:                                                    #sovle for k
            printme = "K: " + str(solve(finalsolve,vari)) + " W/m*degC"
            do(printme)
            finalsep()
        if typekey == 2:                                                    #sovle for l
            printme ="Length: " + str(solve(finalsolve,vari)) + " m"
            do(printme)
            finalsep()
        if typekey == 3:                                                    #sovle for a
            printme = "Area: " + str(solve(finalsolve,vari)) + " m^2"
            do(printme)
            finalsep()
    else:
        finalsep()
        printme = "Final Results"
        do(printme)
        if typekey == 8:                                                        #This means that the final value was t1
            printme = "T1: " + str(result[0]) + ' degC'
            do(printme)
            finalsep()
        if typekey == 9:                                                        #I think I need to fix this, idk if its called out anywhere
            finalsep()
            printme = "T2: " + str(result[0]) + ' degC'
            do(printme)
            finalsep()
        if typekey == 10:                                                       #This means the result was power
            finalsep()
            printme = "Power:  " + str(result[0]) + ' W'
            do(printme)
            finalsep()
    c2 = input("Make a Selection: \n1: Rename Output File \n2: Append Default Name \n\
3: Continue without Saving \nAny Key: Continue with Default Naming (Won't Overwrite)\nEntry: ")                                        #waits for the user and then continues
    if c2 == "1":
        c3 = input("New_Name: ")
        writeoutput(c3,'')
    elif c2 == "2":
        c3 = input("Append with: ")
        writeoutput(fnlong,c3)
    elif c2 == "3":
        pass
    else:
        writeoutput(fnlong,'')
    runagain(1)

###################################################################################################################################################################################################################
    
def cleanup(xv):                                        #simply removes blank spaces that my get put into the list either from csv imports or bad code
    while('' in xv) :                                   #check if there are any blank list items 
        try:
            xv.remove('')                               #removes them
        except:
            pass                                        #if there aren't any '' then this stops an error from closing the program

#######################################################################################################################################################################################################
           
def strip(xv):                                          #not sure if this is still needed
    m = 0                                               #looper to keep position in xv list
    for each in xv:                                     #iterate through xv to check if there are any extra brackets
        pos = m
        if xv[pos] == '<' and xv[pos+2] == '>':
            xv[pos] = ''
            xv[pos+2] = ''
        m=m+1
    cleanup(xv)                                         #calls cleanup function

################################################################################################################################################################################################################################       
def series(xv):                                                     # this will add together any resistors that are in series
    l = 0                                                           #looper to keep position in xv
    for ldash in xv:                                                #iterate through xv
        if ldash == '-':                                            #checks to see if the current item is a - which indicates series
            fser = l                                                #backs up l
            rser = xv[fser-1:fser+2]                                #this creates a mini list of all values near the - to so that it can go through it
            if '>' in rser or '<' in rser:                          #checks to see if the minilist has any parallel brackets, stops it from trying to add these to variables
                c = 0                                               #junk variable to satisfy the requirement of if statement. Does nothing, just exits
            else:                                                   #do this is it doesnt find any parallel brackets. Safe to proceed
                ser = str(xv[fser-1]) + '+' + str(xv[fser+1])       #Concatenates string before and after the "-" with a "+" in the middle
                z = fser-1                                          #sets z to be the first value in the minilist
                w = z                                               #allows you to start on the z position but iterate
                while w < fser+2:                                   #iterates throug main list (basically through the values we isolated in minilist)
                    xv.pop(z)                                       #pops out the item that is in the z postion
                    w=w+1                                           #goes to the next item in the list, this will always pop the z position but will "shift" through the list above
                xv.insert(z, ser)                                   #after all values are removed, it adds the new addition string where rser used to exist
        l=l+1                                                       #tracks position in list for fser use

#######################################################################################################################################################################################################    
        
def parallel(xv):                                                                           #Adds together resistors that are in parallel
    w=0                                                                                     #intilize value for keeping position
    try:                                                                                    #try this and if the next loop passes without finding a 'v' then whole loop breaks and this passes to keep doing series
        try:                                                                                #tries to find the last '<' -- the last one has the best change of simplifying because nesting and whatnot
            while True:                                                                     #keep doing this until it breaks 
                y = xv.index('<',w)                                                         #sets y to be the next instance of '<' -- this starts looking at position w
                w=y+1                                                                       #sets w to be the position of the current '<' so it can search from w next time
        except:                                                                             #when w is the last occurence of '<' the loop breaks because it can't find another '<' - this is janky as fuck
            pass
        z = xv.index('>',y)                                                                 #if the loop is entered correctly, there will always be a '>' after w
        mini = xv[y+1:z]                                                                    #miniloop that includes all values between '<' and '>'
        if '-' in mini:                                                                     #checks to see if anything in mini still has a '-' and needs simplification before proceeding
            pass
        else:                                                                               #if its simplified, start the porallel loop
            summedtop = ''                                                                  #initializes blank variables
            summedbottom = ''
            top1=''
            bottom1=''
            for each in mini:                                                               #goes through the items and parallyzes them
                if each != ',':                                                             #one last check to make sure you don't try to parallyze the comma
                    if '+' in each:                                                         #checks if the values have been simplified by the series() function
                        frac = '('+str(each)+')'                                            #adds an extra set of brackets for resistors i.e. r1+r2 => (r1+r2) for the loop
                    else:                                                                   #if its not a addition problem, you dont need the extra parathesis. 
                        frac = str(each)                                                    #same as above r1 => r1 
                    if bottom1 == '':                                                       #this checks to see if this is the first run though
                        bottom1 = frac                                                      #0/0+null/x = null/x so this just sets the summedbottom value to x ||||| everything below is similar to avoid extra 0's and 1's and ()
                    elif top1 == '':                                                        #seen above, when top1 is null but summedbottom is not, we know it is the second time through the loop
                        bottom2 = frac                                                      #null/x + null/y => sets y (x is problem above)
                        summedtop = '(' + bottom1+'+'+ bottom2 + ')'                        #because null/x+null/y = (x+y)/x*y --- this sets the summed summedtop to x+y
                        summedbottom = bottom1+'*'+bottom2                                  # like above sets the summedbottom
                        top1 = summedtop                                                    #sets top1 to be the new top, it's no longer null
                        bottom1 = summedbottom                                              #I think you get it at this point. Just iterate with new values
                    else:                                                                   #This gets called the 3rd -> Zth time it loops through.
                        bottom2 = frac                                                      #sets bottom 2 to the next value in mini
                        summedtop = '(('+top1+'*'+ bottom2+')'+ '+' + bottom1 + ')'         #x/y+null/z = (z*x+y)/(y*z) --- this is the top
                        summedbottom = '('+bottom1+'*'+bottom2+')'                          #this is the bottom of above equation
                        top1 = summedtop                                                    #sets top amd bottom to the values we just found for future iteratios
                        bottom1 = summedbottom
            final = "(" + '(' +summedbottom+ ')'+'/'+summedtop + ")"                        #Since 1/R = (x+y)/(x*y) we need to invert the top and bottom. Yes, I could have done this from the start
            i=y                                                                             #new variable for new loop
            while i < (z+1):                                                                #loop trough out master list from y to z (above)
                xv.pop(y)                                                                   #pops the value at y, shifting the list as we go 
                i = i+1                                                                     #keeps going a set number of times to avoid extra popping
            xv.insert(y, final)                                                             #inserts our final value into master list where this expression used to exist
    except:
        pass

#######################################################################################################################################################################################################
       
def numerify(xv):                                   #replaces all variables in simplified equation with their values
    xp1 = str(xv[0])                                #grabs string of formula (whole formula should be in position 1
    separator() 
    printme = "Proccessed Network to:\n" + xp1      #prints the processed network
    do(printme)
    separator()
    xp2 = re.split('([*]|[+]|[(]|[)])',xp1)         #re package - splits the network at all non variable functions, but retains them in list
    e = 0
    for each in xp2:                                #goes through the list of variables and math functions
        if each in v:                               #if it existing in the master list of variables, do this
            xp2[e] = str(v.get(each))               #finds the dicitonary value for the resistor name and substitutes it in the list
        e=e+1
    xp3 = ''.join(xp2)                              #after it is all subsituted, return the list to it's original form 
    return xp3                                      #lets this be used when the function is called1

########################################################################################################################################################################################################

def writeoutput(fillet,apnd):
    global output,fnlong
    fl=2
    cfn = False
    try:
        filn = fillet + apnd + '.txt'
        file = open(filn, "x")
    except:
        while cfn == False:
            try:
                filrep =  fillet + apnd + '_' + str(fl)
                filn = filrep + ".txt"
                file = open(filn, "x")
                cfn = True
            except:
                fl = fl + 1
    for each in output:
        file.write(str(each))
    file.close()
    #p = Popen(filn, shell=True)

#######################################################################################################################################################################################################

def runagain(ra):                                                       #this lets you make a modification the file but run again immediately for smooth workflow
    global filename
    print("\n")
    newrunsep()
    newrunsep()
    newrunsep()
    print("\n")
    if ra == 0:                                                         #error restart
        input("Please check your input file and press \
any button to try again")
        blankvar()
        readfile(filename)
    if ra == 1:                                                         #normal restart after completion
        runag = input("Please make a selection: \n    1: Reload and \
Run Again \n    2: Return to Main Menu \n    3: Help \n    4: Quit \nEntry:  ")                     #still need to add the functionality of the other objects, also this is ugly as fuck, but when you line it up it breaks
        if runag == '1': 
            blankvar()                                                  #clears variables
            output.append('\nRunning File Again:' + str(filename))
            readfile(filename)                                          #rereads the file, which also starts the calc
        elif runag == '2':
            start()
        if runag == '3':
            print("Opening Help File.....")
            helpf()
            runagain(1)
        if runag == '4':
            sys.exit()

#######################################################################################################################################################################################################

start()

            

