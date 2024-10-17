# Datum der letzten Änderung: 16.10.2024
# Universität Stuttgart
# Lehrstuhl Verkehrsplanung und Verkehrsleittechnik
# Autor Python: Kea Seelhorst

#POI-Kategorien erstellen, um Shapefiles von Zentralen Orten + Oberbezirke einzulesen

anzahl_poicategories = Visum.Net.POICategories.Count #Anzahl an POI kategorien zählen, um nächste freie Zahl zu ermitteln
poicategory_nr = Visum.Net.POICategories.GetMultiAttValues("No", OnlyActive=False)
nr_letzte_poicategory = poicategory_nr[anzahl_poicategories-1] #Nächste freie Zahl für POI Kategorie definieren

#POI Kategorie für Zentrale Orte erstellen
Visum.Net.AddPOICategory() #POI Kategorie hinzufügen
Visum.Net.POICategories.ItemByKey(int(nr_letzte_poicategory[1])+1).SetAttValue("Code", "RIN_ZentraleOrte") 
Visum.Net.POICategories.ItemByKey(int(nr_letzte_poicategory[1])+1).SetAttValue("Name", "RIN_ZentraleOrte")

#POI Kategorie für Mittelbereiche erstellen
Visum.Net.AddPOICategory()
Visum.Net.POICategories.ItemByKey(int(nr_letzte_poicategory[1])+2).SetAttValue("Code", "RIN_Mittelbereiche") 
Visum.Net.POICategories.ItemByKey(int(nr_letzte_poicategory[1])+2).SetAttValue("Name", "RIN_Mittelbereiche")

Visum.Net.POICategories.ItemByKey(int(nr_letzte_poicategory[1])+1).POIs.AddUserDefinedAttribute("RIN_Zentralität","RIN_Zentralität","RIN_Zentralität",1,0,Ignored=False,MinVal=None, MaxVal=None,DefVal=9,CrossSectionLogic=None, CslIgnoreClosed=False, Formula='', SubAttr='', CanBeEmpty=None,)

# shape-files einlesen
folder_path = Visum.UserPreferences.DocumentName.rsplit('\\', 1)[0]
path_zentraleorte = folder_path + "\\shapefiles\\" + "ZentraleOrte_SWPzone_zone_centroid.SHP"
path_gebiete = folder_path + "\\shapefiles\\" + "ZentraleOrte_mainzone.SHP"

# Shapefile der Zentrale Orte einlesen
importparameter1 = Visum.IO.CreateImportShapeFilePara()
importparameter1.ObjectType = 59 #POIs
importparameter1.SetAttValue("POIKey",int(nr_letzte_poicategory[1])+1)
importparameter1.SetAttributeAllocationsByIDs = True
importparameter1.CreateUserDefinedAttributes = True #Fehlende UDAs erstellen
Visum.IO.ImportShapefile(path_zentraleorte,importparameter1) #Zentrale Orte einlesen

# Shapefile der Mittelbereiche einlesen
importparameter2 = Visum.IO.CreateImportShapeFilePara()
importparameter2.ObjectType=9 #POIs
importparameter2.SetAttValue("POIKey",int(nr_letzte_poicategory[1])+2)
importparameter2.CreateUserDefinedAttributes = True #Fehlende UDAs erstellen
Visum.IO.ImportShapefile(path_gebiete,importparameter2) #Oberbezirke einlesen