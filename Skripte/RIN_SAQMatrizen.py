# Datum der letzten Änderung: 16.10.2024
# Universität Stuttgart
# Lehrstuhl Verkehrsplanung und Verkehrsleittechnik
# Autor Python: Kea Seelhorst


#Funktion mit der eine Matrix durch vorgegebene Eigenschaften gefunden wird
def GetMatrixAttValue_by_Property(Prop,AttID):
    selected_mat = Visum.Net.Matrices.ItemsByRef('Matrix(' + Prop + ')').GetMultiAttValues(AttID)
    # z.B. selected_mat = Visum.Net.Matrices.ItemsByRef('Matrix([AV_ID_CSRS]=1001)').GetMultiAttValues("NO")
    num=selected_mat[0]
    return num[1]

#Letzte Nummer der bestehenden Matrizen herausfinden
last_mat = Visum.Net.Matrices.GetMultiAttValues("NO")
last_mat = int(last_mat[-1][1])

try:
    #Matrix ist schon vorhanden
    GetMatrixAttValue_by_Property("[NAME]=\"SAQ_VLuft\"","NO")

except: 
    #Falls keine Matrix vorhanden, Erstelle Matrix für SAQ Luftliniengeschwindigkeit
    num_Mat_1 = last_mat + 1
    Visum.Net.AddMatrix(num_Mat_1,2,4)
    Visum.Net.Matrices.ItemByKey(num_Mat_1).SetAttValue("CODE", "SAQ_VLuft")
    Visum.Net.Matrices.ItemByKey(num_Mat_1).SetAttValue("NAME", "SAQ_VLuft")


try: 
    #Matrix ist schon vorhanden
    GetMatrixAttValue_by_Property("[NAME]=\"SAQ_DF\"","NO")

except:
    #Falls keine Matrix vorhanden, Erstelle Matrix für SAQ Umwegfaktor
    num_Mat_2 = last_mat + 2
    Visum.Net.AddMatrix(num_Mat_2,2,4)
    Visum.Net.Matrices.ItemByKey(num_Mat_2).SetAttValue("CODE", "SAQ_DF")
    Visum.Net.Matrices.ItemByKey(num_Mat_2).SetAttValue("NAME", "SAQ_DF")
    
