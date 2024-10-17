# Qualitätssicherung von Verkehrsnachfragemodellen
# Führt eine Bewertung der IV-Luftliniengeschwindigkeit nach RIN durch mit 6 LOS-Stufen
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

    # Erstelle einen weiteren Frame, der in das Canvas eingebettet wird
    scrollable_frame = tk.Frame(canvas, bg="white")

    # Füge den Frame in das Canvas ein
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

    # Dynamisch die Höhe des Frames festlegen, um das Scrollen zu aktivieren
    scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

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

# 1. Schritt: Auswahl der Luftlinienmatrix
luftlinienmatrix = matrix_auswahl("Luftlinienmatrix auswählen", "Bitte wählen Sie die Luftlinienmatrix aus:")

# 2. Schritt: Auswahl der Reisezeitmatrix
reisezeitmatrix = matrix_auswahl("Reisezeitmatrix auswählen", "Bitte wählen Sie die Reisezeitmatrix aus:")

# Matrizen aus VISUM laden
MatNo_DID = int(luftlinienmatrix)  # Nummer der ausgewählten Luftlinienmatrix
MatNo_TT_Car = int(reisezeitmatrix)  # Nummer der ausgewählten Reisezeitmatrix
MatNo_SAQ_DS_Car = int(GetMatrixAttValue_by_Property("[NAME]=\"SAQ_VLuft\"","NO")) #Nummer der Ergebnismatrix SAQ Luftliniengeschwindigkeit

MatVal_DID = array(Visum.Net.Matrices.ItemByKey(MatNo_DID).GetValues(), float).tolist()
MatVal_TT_Car = array(Visum.Net.Matrices.ItemByKey(MatNo_TT_Car).GetValues(), float).tolist()
MatVal_LOS_DS_Car = array(Visum.Net.Matrices.ItemByKey(MatNo_SAQ_DS_Car).GetValues(), int).tolist()
Bezirke = array(Visum.Net.Zones.GetMultiAttValues("Name"))

MatVal_TT_Car = np.divide(MatVal_TT_Car, 60)
MatVal_LLV = np.divide(MatVal_DID,MatVal_TT_Car)

#Parameter Berechnung SAQ-Kurven
a = [0.18, 0.21,0.25,0.31,0.39]
b = [-0.676,-0.676,-0.676,-0.676,-0.676]
c = [0.0083,0.0089,0.0096,0.0104,0.0115]

#Matrizen mit Grenzwerten für entsprechende SAQ berechnen
SAQ1_Grenz = 1/ ((a[0]* np.power(MatVal_DID,b[0]))+c[0]) 
SAQ2_Grenz = 1/ ((a[1]* np.power(MatVal_DID,b[1]))+c[1])
SAQ3_Grenz = 1/ ((a[2]* np.power(MatVal_DID,b[2]))+c[2])
SAQ4_Grenz = 1/ ((a[3]* np.power(MatVal_DID,b[3]))+c[3])
SAQ5_Grenz = 1/ ((a[4]* np.power(MatVal_DID,b[4]))+c[4])

#Zuordnung der Relationen zu einer SAQ-Stufe und Schreiben von Stufe
SAQ1_Zuordnung = np.where(MatVal_LLV >= SAQ1_Grenz,1,9)
SAQ2_Zuordnung = np.where(MatVal_LLV >= SAQ2_Grenz,2,9)
SAQ3_Zuordnung = np.where(MatVal_LLV >= SAQ3_Grenz,3,9)
SAQ4_Zuordnung = np.where(MatVal_LLV >= SAQ4_Grenz,4,9)
SAQ5_Zuordnung = np.where(MatVal_LLV >= SAQ5_Grenz,5,9)
SAQ6_Zuordnung = np.where(MatVal_LLV < SAQ5_Grenz,6,9)

SAQStufen = np.where(SAQ6_Zuordnung == 6,6,0) #SAQ 6 Zuordnen
SAQStufen = np.where(SAQ1_Zuordnung == 1,1,SAQStufen) #SAQ 1 Zuordnen
SAQStufen = np.where(SAQ2_Zuordnung < SAQ1_Zuordnung,2,SAQStufen) #SAQ 2 Zuordnen
SAQStufen = np.where(SAQ3_Zuordnung < SAQ2_Zuordnung,3,SAQStufen) #SAQ 3 Zuordnen
SAQStufen = np.where(SAQ4_Zuordnung < SAQ3_Zuordnung,4,SAQStufen) #SAQ 4 Zuordnen
SAQStufen = np.where(SAQ5_Zuordnung < SAQ4_Zuordnung,5,SAQStufen) #SAQ 5 Zuordnen
SAQStufen = np.where(MatVal_DID == 0,1,SAQStufen) #SAQ 1 setzen, wenn DID = 0

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
Visum.Net.Matrices.ItemByKey(MatNo_SAQ_DS_Car).SetValues(SAQStufen)
Visum.Net.Zones.SetMultiAttValues("RIN_ANTEILSAQ1_VLUFT", final1) 
Visum.Net.Zones.SetMultiAttValues("RIN_ANTEILSAQ2_VLUFT", final2) 
Visum.Net.Zones.SetMultiAttValues("RIN_ANTEILSAQ3_VLUFT", final3) 
Visum.Net.Zones.SetMultiAttValues("RIN_ANTEILSAQ4_VLUFT", final4) 
Visum.Net.Zones.SetMultiAttValues("RIN_ANTEILSAQ5_VLUFT", final5) 
Visum.Net.Zones.SetMultiAttValues("RIN_ANTEILSAQ6_VLUFT", final6) 
  
  
