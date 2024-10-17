# Qualitätssicherung von Verkehrsnachfragemodellen
# Führt eine Bewertung des Umwegfaktors nach RIN durch
# Ermittelt die Werte einer SAQ-Matrix
# Berücksichtigt keine Filtermatrix
# Datum der letzten Änderung: 14.10.2024
# Universität Stuttgart
# Lehrstuhl Verkehrsplanung und Verkehrsleittechnik
# Autor Python: Kea Seelhorst

import win32com.client
from numpy import *
import numpy as np
import tkinter as tk
from tkinter import messagebox


#Funktion mit der eine Matrix durch vorgegebene Eigenschaften gefunden wird
def GetMatrixAttValue_by_Property(Prop,AttID):
    selected_mat = Visum.Net.Matrices.ItemsByRef('Matrix(' + Prop + ')').GetMultiAttValues(AttID)
    # z.B. selected_mat = Visum.Net.Matrices.ItemsByRef('Matrix([AV_ID_CSRS]=1001)').GetMultiAttValues("NO")
    num=selected_mat[0]
    return num[1]

# Kenngrößenmatrizen aus dem Modell lesen
i = Visum.Net.Matrices.GetMultipleAttributes(["MatrixType", "No", "Code"], True)
skimmat = [(int(value), code) for typ, value, code in i if typ == 'MATRIXTYPE_SKIM']


# Funktion zur Auswahl einer Matrix mit Scrollfeld
def matrix_auswahl(titel, frage):
    def on_select():
        selected_skimmat = var.get()
        if selected_skimmat:
            root.destroy()  # Fenster schließen, wenn eine Auswahl getroffen wurde
        else:
            messagebox.showwarning("Warnung", "Bitte eine Matrix auswählen.")

    # Erstelle das Hauptfenster
    root = tk.Tk()
    root.title(titel)
    
    # Fensterposition setzen 
    root.wm_geometry("+1000+600")  # Position des Fensters

    # Erstelle den Hauptframe
    main_frame = tk.Frame(root)
    main_frame.pack(fill=tk.BOTH, expand=1)

    # Erstelle Canvas für Scrollfeld mit weißem Hintergrund
    canvas = tk.Canvas(main_frame, bg="white")
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)

    # Scrollbar hinzufügen
    scrollbar = tk.Scrollbar(main_frame, orient=tk.VERTICAL, command=canvas.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    # Verknüpfe die Scrollbar mit dem Canvas
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

    # Frame im Canvas erstellen mit weißem Hintergrund
    scrollable_frame = tk.Frame(canvas, bg="white")
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

    # Variablen für die Matrixauswahl
    var = tk.StringVar()

    # Erstelle Label mit der Frage
    label = tk.Label(scrollable_frame, text=frage, font=("Arial", 14), bg="white")
    label.pack(pady=10)

    # Erstelle Radiobuttons für jede Matrix mit weißem Hintergrund
    for no, code in skimmat:
        rb = tk.Radiobutton(scrollable_frame, text=f"Matrix {no}: {code}", variable=var, value=f"{no}", font=("Arial", 12), bg="white")
        rb.pack(anchor=tk.W, padx=20, pady=5)

    # Bestätigungsbutton außerhalb des Scrollbereichs fest positionieren
    submit_button = tk.Button(root, text="Bestätigen", command=on_select, font=("Arial", 12))
    submit_button.pack(pady=20)

    # Starte die GUI und warte auf Benutzereingabe
    root.mainloop()

    # Rückgabe der Auswahl (Matrixnummer)
    return var.get()


# 1. Schritt: Auswahl der Luftlinienweite-Matrix
luftlinienweite_matrix = matrix_auswahl("Luftlinienweite-Matrix auswählen", "Bitte wählen Sie die Luftlinienweite-Matrix aus:")

# 2. Schritt: Auswahl der Fahrweite-Matrix
fahrweite_matrix = matrix_auswahl("Fahrweite-Matrix auswählen", "Bitte wählen Sie die Fahrweite-Matrix aus:")


# Matrizen aus VISUM laden
MatNo_DID = int(luftlinienweite_matrix)  # Nummer der ausgewählten Luftlinienweite-Matrix
MatNo_DIS = int(fahrweite_matrix)  # Nummer der ausgewählten Fahrweite-Matrix
MatNo_SAQ_DF = int(GetMatrixAttValue_by_Property("[NAME]=\"SAQ_DF\"","NO")) #Nummer der Ergebnismatrix SAQ Luftliniengeschwindigkeit

MatVal_DID = array(Visum.Net.Matrices.ItemByKey(MatNo_DID).GetValues(), float).tolist()
MatVal_DIS = array(Visum.Net.Matrices.ItemByKey(MatNo_DIS).GetValues(), float).tolist()
MatVal_SAQ_DF = array(Visum.Net.Matrices.ItemByKey(MatNo_SAQ_DF).GetValues(), int).tolist()
Bezirke = array(Visum.Net.Zones.GetMultiAttValues("Name"))

MatVal_DF = np.divide(MatVal_DIS,MatVal_DID)

#Parameter Berechnung SAQ-Kurven
a = [0.18, 0.27,0.37,0.48,0.60]
b = [0.950,0.925,0.900,0.870,0.840]


#Matrizen mit Grenzwerten für entsprechende SAQ berechnen
SAQ1_Grenz = np.divide(np.power(10,a[0]+b[0]*np.log10(MatVal_DID)),MatVal_DID)
SAQ2_Grenz = np.divide(np.power(10,a[1]+b[1]*np.log10(MatVal_DID)),MatVal_DID)
SAQ3_Grenz = np.divide(np.power(10,a[2]+b[2]*np.log10(MatVal_DID)),MatVal_DID)
SAQ4_Grenz = np.divide(np.power(10,a[3]+b[3]*np.log10(MatVal_DID)),MatVal_DID)
SAQ5_Grenz = np.divide(np.power(10,a[4]+b[4]*np.log10(MatVal_DID)),MatVal_DID)

#Zuordnung der Relationen zu einer SAQ-Stufe und Schreiben von Stufe
SAQ1_Zuordnung = np.where(MatVal_DF <= SAQ1_Grenz,1,9)
SAQ2_Zuordnung = np.where(MatVal_DF <= SAQ2_Grenz,2,9)
SAQ3_Zuordnung = np.where(MatVal_DF <= SAQ3_Grenz,3,9)
SAQ4_Zuordnung = np.where(MatVal_DF <= SAQ4_Grenz,4,9)
SAQ5_Zuordnung = np.where(MatVal_DF <= SAQ5_Grenz,5,9)
SAQ6_Zuordnung = np.where(MatVal_DF > SAQ5_Grenz,6,9)

SAQStufen = np.where(SAQ6_Zuordnung == 6,6,0) #SAQ 6 Zuordnen
SAQStufen = np.where(SAQ1_Zuordnung == 1,1,SAQStufen) #SAQ 1 Zuordnen
SAQStufen = np.where(SAQ2_Zuordnung < SAQ1_Zuordnung,2,SAQStufen) #SAQ 2 Zuordnen
SAQStufen = np.where(SAQ3_Zuordnung < SAQ2_Zuordnung,3,SAQStufen) #SAQ 3 Zuordnen
SAQStufen = np.where(SAQ4_Zuordnung < SAQ3_Zuordnung,4,SAQStufen) #SAQ 4 Zuordnen
SAQStufen = np.where(SAQ5_Zuordnung < SAQ4_Zuordnung,5,SAQStufen) #SAQ 5 Zuordnen
SAQStufen = np.where(MatVal_DID == 0,1,SAQStufen) #SAQ 1 setzen, wenn DID = 0
SAQStufen = np.where(MatVal_DIS == 0,1,SAQStufen) #SAQ 1 setzen, wenn DIS = 0
SAQStufen = np.where(MatVal_DF == 0,1,SAQStufen) #SAQ 1 setzen, wenn DF = 0

#Prozentuale Anteile SAQ pro Bezirk berechnen
AnzahlBezirke = np.shape(SAQStufen)[0]
SAQ1_Anteil = (np.count_nonzero(SAQStufen ==1 , 0)/AnzahlBezirke)*100
SAQ2_Anteil = (np.count_nonzero(SAQStufen ==2 , 0)/AnzahlBezirke)*100
SAQ3_Anteil = (np.count_nonzero(SAQStufen ==3 , 0)/AnzahlBezirke)*100
SAQ4_Anteil = (np.count_nonzero(SAQStufen ==4 , 0)/AnzahlBezirke)*100
SAQ5_Anteil = (np.count_nonzero(SAQStufen ==5 , 0)/AnzahlBezirke)*100
SAQ6_Anteil = (np.count_nonzero(SAQStufen ==6 , 0)/AnzahlBezirke)*100

#Umwandlung für das Einlesen in Visum
Bezirke = Bezirke[:, 0]
final1 = np.c_[Bezirke, SAQ1_Anteil].tolist()
final2 = np.c_[Bezirke, SAQ2_Anteil].tolist()
final3 = np.c_[Bezirke, SAQ3_Anteil].tolist()
final4 = np.c_[Bezirke, SAQ4_Anteil].tolist()
final5 = np.c_[Bezirke, SAQ5_Anteil].tolist()
final6 = np.c_[Bezirke, SAQ6_Anteil].tolist()

#Zurückschreiben der Werte nach VISUM
Visum.Net.Matrices.ItemByKey(MatNo_SAQ_DF).SetValues(SAQStufen)
#Visum.Net.Matrices.ItemByKey(7).SetValues(MatVal_DF)
Visum.Net.Zones.SetMultiAttValues("RIN_ANTEILSAQ1_DF", final1) 
Visum.Net.Zones.SetMultiAttValues("RIN_ANTEILSAQ2_DF", final2) 
Visum.Net.Zones.SetMultiAttValues("RIN_ANTEILSAQ3_DF", final3) 
Visum.Net.Zones.SetMultiAttValues("RIN_ANTEILSAQ4_DF", final4) 
Visum.Net.Zones.SetMultiAttValues("RIN_ANTEILSAQ5_DF", final5) 
Visum.Net.Zones.SetMultiAttValues("RIN_ANTEILSAQ6_DF", final6)
  
