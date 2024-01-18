import sys
import pandas as pd
nam=input("Ievadiet Excel faila nosaukumu: ")
sh=input("Ievadiet faila lapas nosaukumu: ")
try:
    ex=pd.read_excel(nam, sheet_name=sh)
except FileNotFoundError:
    print(f"Failu {nam} nevarēja atrast")
    sys.exit()
except Exception as e:
    print(f"Kļūda lasot failu: {e}")
    sys.exit()
def cs(value):
    try:
        return float(value)
    except ValueError:
        return float('inf')
into=input("Kā vēlaties sakārtot kolonnu? (Alfabētiski, skaitliski, datums) ")
if(into.lower() == 'alfabētiski' or  into.lower() == 'alfabetiski' ):
    sort=input("Augoši vai dilstoši? ")
    if(sort.lower() == 'augoši' or sort.lower() == 'augosi' ):
        col=input("Kura kolonna? (Ierakstiet kolonnas nosaukumu) ")
        try:
            if ex[col].dtype == 'object':
                ex['temp'] = ex[col].astype(str)
                ex['temp'] = ex['temp'].str.extract('(\D+)')
                ex['temp'] = ex['temp'].fillna('')
                ex = ex.sort_values(by='temp', na_position='last')
                ex = ex.drop(columns=['temp'])
        except KeyError:
            print(f"Kolonna {col} nav atrasta")
            sys.exit()
        except Exception as e:
            print(f"Kļūda: {e}")
            sys.exit()
        name=input("Ierakstiet nosaukumu lapai, kurā saglabāt sakārtoto: ")
        with pd.ExcelWriter(nam,mode='a') as writer:  
           ex.to_excel(writer, sheet_name=name, index=False)
    elif(sort.lower() == 'dilstoši' or sort.lower() == 'dilstosi' ):
        col=input("Kura kolonna? (Ierakstiet kolonnas nosaukumu) ")
        try:
            if ex[col].dtype == 'object':
                ex['temp'] = ex[col].astype(str)
                ex['temp'] = ex['temp'].str.extract('(\D+)')
                ex['temp'] = ex['temp'].fillna('')
                ex = ex.sort_values(by='temp',ascending=False, na_position='last')
                ex = ex.drop(columns=['temp'])
        except KeyError:
            print(f"Kolonna {col} nav atrasta")
            sys.exit()
        except Exception as e:
            print(f"Kļūda: {e}")
            sys.exit()
        name=input("Ierakstiet nosaukumu lapai, kurā saglabāt sakārtoto: ")
        with pd.ExcelWriter(nam,mode='a') as writer:  
           ex.to_excel(writer, sheet_name=name, index=False)
    else:
        print("Lūdzu ierakstiet vienu no piedāvātajām opcijām")
        sys.exit()
elif(into == 'skaitliski' ):
    sort=input("Augoši vai dilstoši? ")
    if(sort.lower() == 'augoši' or sort.lower() == 'augosi' ):
        col=input("Kura kolonna? (Ierakstiet kolonnas nosaukumu) ")
        try:
            ex = ex.sort_values(by=col,key=lambda x: x.map(cs), na_position='last')
        except KeyError:
            print(f"Kolonna {col} nav atrasta")
            sys.exit()
        except Exception as e:
            print(f"Kļūda: {e}")
            sys.exit()
        name=input("Ierakstiet nosaukumu lapai, kurā saglabāt sakārtoto: ")
        with pd.ExcelWriter(nam,mode='a') as writer:  
            ex.to_excel(writer, sheet_name=name, index=False)
    elif(sort.lower() == 'dilstoši' or sort.lower() == 'dilstosi' ):
        col=input("Kura kolonna? (Ierakstiet kolonnas nosaukumu) ")
        try:
            ex = ex.sort_values(by=col,ascending=False,key=lambda x: x.map(cs), na_position='last')
        except KeyError:
            print(f"Kolonna {col} nav atrasta")
            sys.exit()
        except Exception as e:
            print(f"Kļūda: {e}")
            sys.exit()
        name=input("Ierakstiet nosaukumu lapai, kurā saglabāt sakārtoto: ")
        with pd.ExcelWriter(nam,mode='a') as writer:  
            ex.to_excel(writer, sheet_name=name, index=False)
    else:
        print("Lūdzu ierakstiet vienu no piedāvātajām opcijām")
        sys.exit()
elif(into == 'datums' ):
    sort=input("Augoši vai dilstoši? ")
    if(sort.lower() == 'augoši' or sort.lower() == 'augosi' ):
        col=input("Kura kolonna? (Ierakstiet kolonnas nosaukumu) ")
        try:
            if ex[col].dtype == 'datetime64[ns]':
                ex = ex.sort_values(by=col, ascending=True, na_position='last')
            else:
                try:
                    ex['temp'] = pd.to_datetime(ex[col], errors='coerce')
                    ex = ex.sort_values(by='temp', ascending=True, na_position='last')
                    ex = ex.drop(columns=['temp'])
                except ValueError:
                    print(f"Kolonnas {col} vērtības nevar pārvērst par datumu")
                    sys.exit()
        except KeyError:
            print(f"Kolonna {col} nav atrasta")
            sys.exit()
        except Exception as e:
            print(f"Kļūda: {e}")
            sys.exit()
        name=input("Ierakstiet nosaukumu lapai, kurā saglabāt sakārtoto: ")
        with pd.ExcelWriter(nam,mode='a') as writer:  
           ex.to_excel(writer, sheet_name=name, index=False)
    elif(sort.lower() == 'dilstoši' or sort.lower() == 'dilstosi' ):
        col=input("Kura kolonna? (Ierakstiet kolonnas nosaukumu) ")
        try:
            if ex[col].dtype == 'datetime64[ns]':
                ex = ex.sort_values(by=col, ascending=False, na_position='last')
            else:
                try:
                    ex['temp'] = pd.to_datetime(ex[col], errors='coerce')
                    ex = ex.sort_values(by='temp', ascending=False, na_position='last')
                    ex = ex.drop(columns=['temp'])
                except ValueError:
                    print(f"Kolonnas {col} vērtības nevar pārvērst par datumu")
                    sys.exit()
        except KeyError:
            print(f"Kolonna {col} nav atrasta")
            sys.exit()
        except Exception as e:
            print(f"Kļūda: {e}")
            sys.exit()
        name=input("Ierakstiet nosaukumu lapai, kurā saglabāt sakārtoto: ")
        with pd.ExcelWriter(nam,mode='a') as writer:  
           ex.to_excel(writer, sheet_name=name, index=False)
    else:
        print("Lūdzu ierakstiet vienu no piedāvātajām opcijām")
        sys.exit()
else:
    print("Lūdzu ierakstiet vienu no piedāvātajām opcijām")
    sys.exit()
