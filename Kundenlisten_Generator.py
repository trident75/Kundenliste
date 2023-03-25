# Generierung von Zufalls-Kundenlisten
# Autor: Marko Landrath, trident75@live.de
'''
ToDo:
- JSON
- XML
- Webformular
- TKint?
'''
import random       #Zur generierung von Zufallswerten/Auswahl
import xlsxwriter   #Zum Schreiben von Excel-Tabellen
import sqlite3      #Zur Datenbank-Erstellung

#Definition des Kunden-"Arrays" und des ersten Datensatzes mit den Überschriften
kunden=[]
kunde=[('ID','Vorname','Nachname','Strasse','Ort','E-Mail','Gehalt')]

#Definition der Auswahl-Daten
vornamen=['Anna','Achim','Beathe','Bernd','Christine','Christian','Dorothee','Dieter','Erich','Elisabet','Friedrich','Franziska','Gudrun','Günther','Heidi','Heinrich','Ilse','Ingo','Jessica','Jürgen','Katrin','Klaus','Lena','Ludwig','Martina','Markus','Nina','Norbert','Odilie','Olaf','Petra','Peter','Quinn','Ramona','Richard','Susi','Stefan','Thea','Thomas','Ulrike','Ulrich','Vanessa','Veit','Walter','Waltraut','Xian','Yu','Zacharias']
nachnamen=['Abendbrot','Bunsen','Chow','Dübel','Elster','Feierabend','Grau','Habsburg','Ichsel','Jaworski','Klabauterbach','Laugenberg','Müller','Nawattnu','Obacht','Preuss','Quinoa','Reinlein','Schmidt','Tucholsky','Unwerth','Vechsel','Winterberg','Xu','Yuarez','Zerbern']
strassen1=['Haupt','Schul','Dorf','Kirchen','Tiergarten','Harburger','Mönckeberg','Kieler','Berliner','Industrie','Kraftwerks','Feld','Meisen','Rotkelchen','Bundes','Adler','Elb']
strassen2=['tangente','platz','chaussee','redder','weg','strasse','allee','laufbahn','berg','ufer']
orte=['Hamburg','Kiel','Berlin','Rothenburg','Bonn','Frankfurt','München','Stuttgart','Bremen','Hannover','Schwerin','Rostock','Jena','Lüneburg']
domains=['yahoo.de','web.de','googlemail.com','outlook.de','lycos.de']

#Definition der Augaben
def Bildschirmausgabe(kunden):
    print('Ausgabe auf den Bildschirm')
    print('')
    for i in range(len(kunden)):
        print("Datensatz-Nr.:",str(i+1))
        print(kunden[i][1], kunden[i][2])
        print(kunden[i][3])
        print(kunden[i][4])
        print(kunden[i][5])
        print(kunden[i][6])
        print()
    return

def CSV_Ausgabe(kunden):
    with open('kundenliste.csv','w', encoding='UTF-8') as datei:
        for innerlist in kunden:
            for item in innerlist:
                datei.write(str(item)+';')
            datei.write('\n')
    return
        
def XLSX_Ausgabe(kunden):
    workbook = xlsxwriter.Workbook('Kundenliste.xlsx')
    worksheet = workbook.add_worksheet('Kunden')
    bold=workbook.add_format({'bold':True})
    worksheet.write('A1','ID',bold)
    worksheet.write('B1','Vorname',bold)
    worksheet.write('C1','Nachname',bold)
    worksheet.write('D1','Strasse',bold)
    worksheet.write('E1','Ort',bold)
    worksheet.write('F1','eMail',bold)
    worksheet.write('G1','Gehalt',bold)
    row=0
    col=0
    for kunde in kunden:
        row+=1
        worksheet.write_number(row, col, kunde[0])
        worksheet.write_string(row, col+1, kunde[1])
        worksheet.write_string(row, col+2, kunde[2])
        worksheet.write_string(row, col+3, kunde[3])
        worksheet.write_string(row, col+4, kunde[4])
        worksheet.write_string(row, col+5, kunde[5])
        worksheet.write_number(row, col+6, kunde[6])
    worksheet.autofilter(0,0,row,6)
    worksheet.autofit()
    workbook.close()
    return

def html_ausgabe(kunden):
    html_table = "<table>\n"
    for innerlist in kunden:
        html_table += "<tr>"
        for item in innerlist:
            html_table += "<td>{}</td>".format(item)
        html_table += "</tr>\n"
    html_table += "</table>"
    with open('kundenliste.html','w', encoding='UTF-8') as datei:
        datei.write(html_table)
    return

def SQLite_ausgabe(kunden):
    db_connection = sqlite3.connect('Kundendaten.sqlite')
    cursor=db_connection.cursor()
    #cursor.execute('if exist drop table kundendaten')
    cursor.execute('create table if not exists kundendaten(ID integer,Vorname text,Nachname text,Strasse text,Ort text,eMail text,Gehalt integer)')
    cursor.executemany('insert into kundendaten values (?,?,?,?,?,?,?)',kunden)
    db_connection.commit()
    db_connection.close()
    return

def Testdruck(kunden):
    i=0
    for i in range(len(kunden)):
        print(kunden[i])
    return

#Abfrage: Wieviele Datensätze wohin?
anzahl=input('Wie viele Datensätze sollen generiert werden? ')
numbers=[1,2,3,4,5,8,9] #Liste der möglichen Ausgabeziele
ausgabe=0
while ausgabe not in numbers:
    ausgabe=int(input('Wie soll die Ausgabe erfolgen? \n1: Screen \n2: CSV-Datei \n3: XLSX-Datei \n4: HTML-Tabelle \n5: SQlite Datenbank \n8: Test-Druck Datensätze \n9: Abbrechen \n? '))
print()

#Schleife zur Datensatz-Generierung und anhängen des Datensatzes an Liste kunden
for i in range(int(anzahl)):
    idnummer=i+1
    vorname=random.choice(vornamen)
    nachname=random.choice(nachnamen)
    strasse=random.choice(strassen1)+random.choice(strassen2)+' '+str(random.randint(1,99))
    ort=str(random.randint(10000, 99999))+' '+random.choice(orte)
    email=vorname.lower()+'.'+nachname.lower()+'@'+random.choice(domains)
    gehalt=random.randint(10,250)*1000
    kunde=[idnummer, vorname, nachname, strasse, ort, email, gehalt]
    kunden.append(kunde)

#Fallabfrage, wo die Datensätze ausgegeben werden sollen und aufruf der entsprechenden Funktion
if ausgabe==1:
    Bildschirmausgabe(kunden)
elif ausgabe==2:
    CSV_Ausgabe(kunden)    
elif ausgabe==3:
    XLSX_Ausgabe(kunden)    
elif ausgabe==4:
    html_ausgabe(kunden)
elif ausgabe==5:
    SQLite_ausgabe(kunden)
elif ausgabe==8:
    Testdruck(kunden)
else:
    print('Abgebrochen.')

#Programmende
print('Programm beendet.')
