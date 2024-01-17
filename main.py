import sys
import pandas as pd
ex=pd.read_excel('example.xlsx')
into=input("How would you like the column to be sorted? (Alphabetically, numerically)")
if(into == 'alphabetically' ):
    sort=input("Ascending or descending?")
    if(sort == 'ascending' ):
        col=input("Which column?")
        ex=ex.sort_values(by='col')
    elif(sort == 'descending' ):
        col=input("Which column?")
        ex=ex.sort_values(by='col',ascending=False)
    else:
        print("Please input one of the offered options")
        sys.exit()
elif(into == 'numerically' ):
    sort=input("Ascending or descending?")
    if(sort == 'ascending' ):
       col=input("Which column?")
       ex=ex.sort_values(by='col')
    elif(sort == 'descending' ):
        col=input("Which column?")
        ex=ex.sort_values(by='col',ascending=False)
    else:
        print("Please input one of the offered options")
        sys.exit()
else:
    print("Please input one of the offered options")
    sys.exit()

print(ex)