from csv import DictReader
from csv import DictWriter
from openpyxl import Workbook
from openpyxl import styles
from openpyxl import load_workbook
from tabulate import tabulate
import pandas as pd
import numpy as np
from datetime import datetime
import os


def data_cleaning(val):
    if val is None:
        return 999
    elif isinstance(val, str) and val.strip() == '':
        return 999
    elif isinstance(val, list) and not val:  # Ha üres a lista
        return 999
    elif isinstance(val, float) and np.isnan(val):  # Ha NaN az érték
        return 999
    elif val == 'V':
        return 'W'
    elif val == 'R':
        return 'T'
    elif val == 'P':
        return 'L'
    elif val in ['OK', 'Ok', 'ok']:
        return 0
    elif val in ['outside', 'error', 'chyba', 'mimo', 'vběhl', 'run in']:
        return 999
    else:
        return val

def pandas_processor():
    directory = '/Users/radnar/Documents/Development/Projects/Data36_Flyball/Sources/Competitions/'

    # Lista az Excel fájlok tárolására
    excel_files = []
    # Magyar hónapnevek és azok angol megfelelői

    magyar_honapok = ['január', 'február', 'március', 'április', 'május', 'június', 'július', 'augusztus', 'szeptember',
                      'október', 'november', 'december']
    angol_honapok = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October',
                     'November', 'December']
    # Könyvtár tartalmának beolvasása
    for filename in os.listdir(directory):
        # Ellenőrzés, hogy az adott elem fájl-e és Excel fájl-e
        if filename.endswith('.xlsx') or filename.endswith('.xls'):
            excel_files.append(os.path.join(directory, filename))

    processed_data = []

    for excel_file in excel_files:
        xls = pd.ExcelFile(excel_file)
        workbook = load_workbook(filename=excel_file)

        # Munkalapok (sheetek) neveinek lekérdezése
        sheet_names = xls.sheet_names

        # Munkalapok (sheetek) neveinek kiíratása
        custom_headers = ['Nr', 'Division', 'when', 'who', 'whom', 'total time', 'W/L/T', 'Hurdles', 'name1', 'start1', 'time1', 'name2', 'start2', 'time2', 'name3', 'start3', 'time3', 'name4', 'start4', 'time4']

        for sheet_name in sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

            sheet = workbook[sheet_name]
        
            columns_to_check = df.columns[2:]  # A 4. oszloptól kezdve az összes oszlop
            duplicates_in_columns = df[columns_to_check][df[columns_to_check].duplicated()]

            # Szemre ellenőrzöm csak, hogy van-e duplikált érték
            if not duplicates_in_columns.empty:
                print("Vannak duplikált értékek a 4. oszloptól kezdve:")
                print(duplicates_in_columns)

            # Itt folytathatja a DataFrame df feldolgozását
            print(f"A(z) {sheet_name} munkalap adatai:")
            print(df.iloc[0,1])

            first_cell_value = df.iloc[0,1]
            info_in_parentheses = first_cell_value.split('(')[-1].split(')')[0].split()
            # A verseny neve a zárójel előtt
            race_name = first_cell_value.split('(')[0].strip()
            # A város az első elem a zárójelben lévő információkból
            city = info_in_parentheses[0]
            datum = first_cell_value.split(')')[1].split('-')[0]

            if datum[-1] != '.':
                datum = first_cell_value.split(')')[1].split('-')[0]+'.'

            for magyar_honap, angol_honap in zip(magyar_honapok, angol_honapok):
                datum = datum.replace(magyar_honap, angol_honap)

            # A dátum a többi elem a zárójelben lévő információkból
            date = datetime.strptime(datum.strip(), '%Y. %B %d.').replace(hour=0, minute=0, second=0, microsecond=0)
            # Verseny információk tárolása a dictionary-ben
            print(f'race_name: {race_name}, city: {city}, date: {date}')

            # Fejlécek beállítása
            df.columns = custom_headers
            #df = df.iloc[2:, :]

            row_num = 2

            for index, row in df.iloc[2:].iterrows():

                act_dogs = []
                if row.isnull().all():
                    continue
                elif pd.isnull(row['W/L/T']):
                    continue
                else:
                    dogs = [(row['name' + str(i)], row['start' + str(i)], row['time' + str(i)], i) for i in range(1, 5)]
                    division = row['Division']
                    when = row['when'].strftime('%H:%M:%S')
                    w_pair = False
                    current_when = row['when']

                    if index != 2:  # Ha nem az első sor
                        previous_when = df.loc[index - 1, 'when']
                        if index != len(df) - 1:  # Ha nem az utolsó sor
                            next_when = df.loc[index + 1, 'when']
                            if current_when != previous_when and current_when != next_when:
                                #print("Az", index,". sor when értéke nem egyezik az előtte lévő és az utána lévő sor when értékével.")
                                w_pair = True
                        else:
                            continue

                    who = row['who'].title()
                    total_time = row['total time']
                    wl = data_cleaning(row['W/L/T'])
                    hurdles = row['Hurdles']
                    
                    for dog, st, ti, i in dogs:
                        cell = sheet[f'D{index+1}']
                        print(cell)

                        if cell.font.color is not None:
                            try:
                                cell_color = cell.font.color.rgb
                                # Biztosítjuk, hogy a font_color sztring típusú legyen
                                if isinstance(cell_color, str):
                                    print("A betűszín RGB kódja:", cell.font.color)
                                else:
                                    color_index = cell.font.color.indexed
                                    #cell_color = cell.font.color.indexed
                                    color_map = styles.colors.COLOR_INDEX
                                    cell_color = color_map[color_index]
                            except AttributeError:
                                print("A cella betűszíne nem rendelkezik RGB kódjával.")
                        else:
                            print("A cellának nincs beállított betűszíne.")

                        act_dogs.append(dog)
                        t = data_cleaning(ti)
                        s = data_cleaning(st)
                        if (t == 999 or s == 999):
                            error = True
                        elif (float(t) < 0 or float(s) < 0):
                            error = True
                        else:
                            error = False
                        if (wl == 'W' and error == True) or (t < 3 and t >= 0) or (np.isnan(hurdles) or (isinstance(dog, float) and np.isnan(dog)) or (isinstance(dog, int) and np.isnan(dog)) ):
                            anomaly = True
                        else:
                            anomaly = False

                        processed_data.append({
                            'city' : city,
                            'race_name' : race_name,
                            'date' : date,
                            'when': when,
                            'Division': division,
                            'who': who,
                            'total_time': total_time,
                            'W/L/T': wl,
                            'Hurdles': hurdles,
                            'name': dog,
                            'start': s,
                            'time': t,
                            'error': error,
                            'anomaly': anomaly,
                            'w_pair': w_pair,
                            'ordinal_no': i,
                            'dogs' : act_dogs,
                            'szin' : cell_color
                        })

    df_csv = pd.DataFrame(processed_data)

    # A CSV fájl elérési útvonala
    csv_file_path = '/Users/radnar/Documents/Development/Projects/Data36_Flyball/Sources/FileResults/Competitions.csv'

    # A DataFrame adatok kiírása CSV fájlba
    result_headers = ['city', 'race_name', 'date', 'when', 'Division', 'who', 'total_time', 'W/L/T', 'Hurdles', 'name', 'start', 'time', 'error', 'anomaly', 'w_pair', 'ordinal_no', 'dogs', 'szin']
    df_csv.to_csv(csv_file_path, index=False, columns=result_headers)


def analyzis(team_name='WRTF'):
    # Szeretnék különböző információkat gyűjteni a csapat kutyáiról
    # Átlagos teljesítmény
    # Átlagos teljesítmény pozíciónként


    df = pd.read_csv('/Users/radnar/Documents/Development/Projects/Data36_Flyball/Sources/FileResults/Competitions.csv', sep=',')
    dogs = pd.read_csv('/Users/radnar/Documents/Development/Projects/Data36_Flyball/Sources/FileResults/Wild_Runners_Dogs.csv', sep=';')
    
    wild_runners_data = df[df['who'].str.contains('WildRunners', na=False, case=False)]
    print(tabulate(wild_runners_data.head()))
    print(tabulate(dogs.head()))

    merged = wild_runners_data.merge(dogs, left_on='name', right_on='nev')
    print(tabulate(merged.head()))
    return


pandas_processor()
#analyzis()
