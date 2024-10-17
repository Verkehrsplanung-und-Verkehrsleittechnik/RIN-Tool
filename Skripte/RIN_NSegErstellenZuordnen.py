# Datum der letzten Änderung: 16.10.2024
# Universität Stuttgart
# Lehrstuhl Verkehrsplanung und Verkehrsleittechnik
# Autor Python: Kea Seelhorst


import tkinter as tk
from tkinter import messagebox

# Modi aus dem Modell lesen
modi = [modus[1] for modus in Visum.Net.Modes.GetMultiAttValues("Code")]

def on_select():
    selected_modus = var.get()
    if selected_modus:
        # Schließe das Fenster, wenn eine Auswahl getroffen wurde
        root.destroy()
    else:
        messagebox.showwarning("Warnung", "Bitte einen Modus auswählen.")

# Erstelle das Hauptfenster
root = tk.Tk()
root.title("Modus auswählen")

# Fensterposition setzen 
root.wm_geometry("+1000+600")  # Position des Fensters

# Dynamische Höhe basierend auf der Anzahl der Modi
modus_height = len(modi) * 40  # Höhe pro Radiobutton
window_width = 500  # Festgelegte Breite
root.geometry(f"{window_width}x{modus_height + 200}")

# Variablen für die Modusauswahl
var = tk.StringVar()


# Erstelle Radiobuttons für jeden Modus
for modus in modi:
    rb = tk.Radiobutton(root, text=modus, variable=var, value=modus, font=("Arial", 12))  # Schriftart und Größe
    rb.pack(anchor=tk.W, padx=20, pady=5)  # Padding hinzufügen

# Bestätigungsbutton
submit_button = tk.Button(root, text="Bestätigen", command=on_select, font=("Arial", 12))
submit_button.pack(pady=20)

# Starte die GUI und warte auf Benutzereingabe
root.mainloop()

# Speichere den ausgewählten Modus als Variable
moduswahl = var.get()

try:
    #Nachfragesegmente erstellen und Nachfragematrizen zurordnen
    Visum.Net.AddDemandSegment("VFS0", moduswahl)
    Visum.Net.DemandSegments.ItemByKey("VFS0").GetDemandDescription().SetAttValue("DemandTimeSeriesNo", "0")
    Visum.Net.DemandSegments.ItemByKey("VFS0").GetDemandDescription().SetAttValue("Matrix", "Matrix([CODE]=\"RIN_VFS 0_n=2\")")
    Visum.Net.AddDemandSegment("VFS1", moduswahl)
    Visum.Net.DemandSegments.ItemByKey("VFS1").GetDemandDescription().SetAttValue("DemandTimeSeriesNo", "0")
    Visum.Net.DemandSegments.ItemByKey("VFS1").GetDemandDescription().SetAttValue("Matrix", "Matrix([CODE]=\"RIN_VFS 1_n=2\")")
    Visum.Net.AddDemandSegment("VFS2", moduswahl)
    Visum.Net.DemandSegments.ItemByKey("VFS2").GetDemandDescription().SetAttValue("DemandTimeSeriesNo", "0")
    Visum.Net.DemandSegments.ItemByKey("VFS2").GetDemandDescription().SetAttValue("Matrix", "Matrix([CODE]=\"RIN_VFS 2_n=2\")")
    Visum.Net.AddDemandSegment("VFS3", moduswahl)
    Visum.Net.DemandSegments.ItemByKey("VFS3").GetDemandDescription().SetAttValue("DemandTimeSeriesNo", "0")
    Visum.Net.DemandSegments.ItemByKey("VFS3").GetDemandDescription().SetAttValue("Matrix", "Matrix([CODE]=\"RIN_VFS 3_n=2\")")
    Visum.Net.AddDemandSegment("VFS4", moduswahl)
    Visum.Net.DemandSegments.ItemByKey("VFS4").GetDemandDescription().SetAttValue("DemandTimeSeriesNo", "0")
    Visum.Net.DemandSegments.ItemByKey("VFS4").GetDemandDescription().SetAttValue("Matrix", "Matrix([CODE]=\"RIN_VFS 4_n=2\")")

    comment0 = "Umlegung VFS0"
    comment1 = "Umlegung VFS1"
    comment2 = "Umlegung VFS2"
    #in Umlegungsschritt Nachfragesegment zuordnen
    for procedure in Visum.Procedures.Operations.GetAll:
        if comment0 in procedure.AttValue("Comment"):
            procedure.PrTAssignmentParameters.SetAttValue("DSegSet", "VFS0")
        elif comment1 in procedure.AttValue("Comment"):
            procedure.PrTAssignmentParameters.SetAttValue("DSegSet", "VFS1")
        elif comment2 in procedure.AttValue("Comment"):
            procedure.PrTAssignmentParameters.SetAttValue("DSegSet", "VFS2")
except:
    #falls NSeg schon angelegt, Nachfragematrizen zuordnen
    Visum.Net.DemandSegments.ItemByKey("VFS4").GetDemandDescription().SetAttValue("DemandTimeSeriesNo", "0")
    Visum.Net.DemandSegments.ItemByKey("VFS4").GetDemandDescription().SetAttValue("Matrix", "Matrix([CODE]=\"RIN_VFS 4_n=2\")")
    Visum.Net.DemandSegments.ItemByKey("VFS3").GetDemandDescription().SetAttGValue("DemandTimeSeriesNo", "0")
    Visum.Net.DemandSegments.ItemByKey("VFS3").GetDemandDescription().SetAttValue("Matrix", "Matrix([CODE]=\"RIN_VFS 3_n=2\")")
    Visum.Net.DemandSegments.ItemByKey("VFS2").GetDemandDescription().SetAttValue("DemandTimeSeriesNo", "0")
    Visum.Net.DemandSegments.ItemByKey("VFS2").GetDemandDescription().SetAttValue("Matrix", "Matrix([CODE]=\"RIN_VFS 2_n=2\")")
    Visum.Net.DemandSegments.ItemByKey("VFS1").GetDemandDescription().SetAttValue("DemandTimeSeriesNo", "0")
    Visum.Net.DemandSegments.ItemByKey("VFS1").GetDemandDescription().SetAttValue("Matrix", "Matrix([CODE]=\"RIN_VFS 1_n=2\")")
    Visum.Net.DemandSegments.ItemByKey("VFS0").GetDemandDescription().SetAttValue("DemandTimeSeriesNo", "0")
    Visum.Net.DemandSegments.ItemByKey("VFS0").GetDemandDescription().SetAttValue("Matrix", "Matrix([CODE]=\"RIN_VFS 0_n=2\")")
	