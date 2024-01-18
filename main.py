import sys
import pandas as pd
try:
    ex=pd.read_excel('example.xlsx')
except Exception as e:
    print(f"Kļūda lasot failu: {e}")
    sys.exit()
into=input("Kā vēlaties sakārtot kolonnu? (Alfabētiski, skaitliski, datums) ")
if(into == 'alfabētiski' or  into == 'alfabetiski' ):
    sort=input("Augoši vai dilstoši? ")
    if(sort == 'augoši' or sort == 'augosi' ):
        col=input("Kura kolonna? (Ierakstiet kolonnas nosaukumu) ")
        try:
            ex=ex.sort_values(by=col, na_position='last')
        except KeyError:
            print(f"Kolonna {col} nav atrasta")
            sys.exit()
        except Exception as e:
            print(f"Kļūda: {e}")
            sys.exit()
        name=input("Ierakstiet nosaukumu lapai, kurā saglabāt sakārtoto: ")
        with pd.ExcelWriter('example.xlsx',mode='a') as writer:  
           ex.to_excel(writer, sheet_name=name, index=False)
    elif(sort == 'dilstoši' or sort == 'dilstosi' ):
        col=input("Kura kolonna? (Ierakstiet kolonnas nosaukumu) ")
        try:
            ex=ex.sort_values(by=col,ascending=False, na_position='last')
        except KeyError:
            print(f"Kolonna {col} nav atrasta")
            sys.exit()
        except Exception as e:
            print(f"Kļūda: {e}")
            sys.exit()
        name=input("Ierakstiet nosaukumu lapai, kurā saglabāt sakārtoto: ")
        with pd.ExcelWriter('example.xlsx',mode='a') as writer:  
           ex.to_excel(writer, sheet_name=name, index=False)
    else:
        print("Lūdzu ierakstiet vienu no piedāvātajām opcijām")
        sys.exit()
elif(into == 'skaitliski' ):
    sort=input("Augoši vai dilstoši? ")
    if(sort == 'augoši' or sort == 'augosi' ):
        col=input("Kura kolonna? (Ierakstiet kolonnas nosaukumu) ")
        try:
            ex=ex.sort_values(by=col, na_position='last')
        except KeyError:
            print(f"Kolonna {col} nav atrasta")
            sys.exit()
        except Exception as e:
            print(f"Kļūda: {e}")
            sys.exit()
        name=input("Ierakstiet nosaukumu lapai, kurā saglabāt sakārtoto: ")
        with pd.ExcelWriter('example.xlsx',mode='a') as writer:  
            ex.to_excel(writer, sheet_name=name, index=False)
    elif(sort == 'dilstoši' or sort == 'dilstosi' ):
        col=input("Kura kolonna? (Ierakstiet kolonnas nosaukumu) ")
        try:
            ex=ex.sort_values(by=col,ascending=False, na_position='last')
        except KeyError:
            print(f"Kolonna {col} nav atrasta")
            sys.exit()
        except Exception as e:
            print(f"Kļūda: {e}")
            sys.exit()
        name=input("Ierakstiet nosaukumu lapai, kurā saglabāt sakārtoto: ")
        with pd.ExcelWriter('example.xlsx',mode='a') as writer:  
            ex.to_excel(writer, sheet_name=name, index=False)
    else:
        print("Lūdzu ierakstiet vienu no piedāvātajām opcijām")
        sys.exit()
elif(into == 'datums' ):
    sort=input("Augoši vai dilstoši? ")
    if(sort == 'augoši' or sort == 'augosi' ):
        col=input("Kura kolonna? (Ierakstiet kolonnas nosaukumu) ")
        try:
            ex=ex.sort_values(by=col, na_position='last')
        except KeyError:
            print(f"Kolonna {col} nav atrasta")
            sys.exit()
        except Exception as e:
            print(f"Kļūda: {e}")
            sys.exit()
        name=input("Ierakstiet nosaukumu lapai, kurā saglabāt sakārtoto: ")
        with pd.ExcelWriter('example.xlsx',mode='a') as writer:  
           ex.to_excel(writer, sheet_name=name, index=False)
    elif(sort == 'dilstoši' or sort == 'dilstosi' ):
        col=input("Kura kolonna? (Ierakstiet kolonnas nosaukumu) ")
        try:
            ex=ex.sort_values(by=col,ascending=False, na_position='last')
        except KeyError:
            print(f"Kolonna {col} nav atrasta")
            sys.exit()
        except Exception as e:
            print(f"Kļūda: {e}")
            sys.exit()
        name=input("Ierakstiet nosaukumu lapai, kurā saglabāt sakārtoto: ")
        with pd.ExcelWriter('example.xlsx',mode='a') as writer:  
           ex.to_excel(writer, sheet_name=name, index=False)
    else:
        print("Lūdzu ierakstiet vienu no piedāvātajām opcijām")
        sys.exit()
else:
    print("Lūdzu ierakstiet vienu no piedāvātajām opcijām")
    sys.exit()
