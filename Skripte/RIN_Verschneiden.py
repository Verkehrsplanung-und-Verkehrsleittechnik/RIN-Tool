# Datum der letzten Änderung: 16.10.2024
# Universität Stuttgart
# Lehrstuhl Verkehrsplanung und Verkehrsleittechnik
# Autor Python: Kea Seelhorst

#Bestimmen der Nummer der POI-Kategorie der Zentralen Orte
anzahl_poicategories = Visum.Net.POICategories.Count
poicategory_nr = Visum.Net.POICategories.GetMultiAttValues("No", OnlyActive=False)
nr_poicategory = int(poicategory_nr[anzahl_poicategories-2][1])

poicat = "POIOFCAT_"+ str(nr_poicategory)

# Zentralität aller Zonen auf 9 setzen; 9 = keine Zentralität
Visum.Net.Zones.SetAllAttValues("RIN_Zentralität","9", False,False)
Visum.Filters.InitAll()

# Verschneiden-Schritt - Speichern der Nummer des zugehörigen POI der Zentralen Orte (zum filtern)
x = Visum.Net.CreateIntersectAttributePara()
x.SetAttValue("SOURCENETOBJECTTYPE", poicat)
x.SetAttValue("SOURCENETOBJECTINCLUDESUBCATEGORIES", "0")
x.SetAttValue("SOURCEONLYACTIVE", "0")
x.SetAttValue("SOURCEBUFFERSIZE", "0m")
x.SetAttValue("DESTNETOBJECTTYPE", "Zone")
x.SetAttValue("DESTNETOBJECTINCLUDESUBCATEGORIES", "0")
x.SetAttValue("DESTONLYACTIVE", "0")
x.SetAttValue("DESTBUFFERSIZE", "0m")
x.SetAttValue("RANKATTRNAME", "SPECIALENTRY_EMPTY")
x.SetAttValue("SMALLRANKIMPORTANT", "1")
x.SetAttValue("SOURCEATTRNAME", "NO")
x.SetAttValue("DESTATTRNAME", "RIN_NRZUGEHOERIGERPOI")
x.SetAttValue("NUMERICOPERATION", "INTERSECTION_MAXIMUMSHARE")
x.SetAttValue("STRINGOPERATION", "INTERSECTION_CONCATENATE")
x.SetAttValue("ROUND", "0")
x.SetAttValue("ADDVALUE", "0")
x.SetAttValue("WEIGHTBYINTERSECTIONAREASHARE", "0")
x.SetAttValue("CONCATMAXLEN", "255")
x.SetAttValue("CONCATSEPARATOR", ",")

Visum.Net.IntersectAttributes(x)

# Filtern nach Bezirken, denen ein POI der Zentalen Orte zugeordnet werden konnte
Visum.Filters.ZoneFilter().AddCondition(Op="OP_NONE", Complement=True, AttID="RIN_NrZugehoerigerPOI", CompareOperator="IsEmpty", Val="IsEmpty", Position=-1)

# Verschneiden-Schritt - Übertragen der Zentralität auf Bezirke
y = Visum.Net.CreateIntersectAttributePara()
y.SetAttValue("SOURCENETOBJECTTYPE", poicat)
y.SetAttValue("SOURCENETOBJECTINCLUDESUBCATEGORIES", "0")
y.SetAttValue("SOURCEONLYACTIVE", "0")
y.SetAttValue("SOURCEBUFFERSIZE", "0m")
y.SetAttValue("DESTNETOBJECTTYPE", "Zone")
y.SetAttValue("DESTNETOBJECTINCLUDESUBCATEGORIES", "0")
y.SetAttValue("DESTONLYACTIVE", "1")
y.SetAttValue("DESTBUFFERSIZE", "0m")
y.SetAttValue("RANKATTRNAME", "SPECIALENTRY_EMPTY")
y.SetAttValue("SMALLRANKIMPORTANT", "1")
y.SetAttValue("SOURCEATTRNAME", "RIN_ZENTRALITÄT")
y.SetAttValue("DESTATTRNAME", "RIN_ZENTRALITÄT")
y.SetAttValue("NUMERICOPERATION", "INTERSECTION_MAXIMUMSHARE")
y.SetAttValue("STRINGOPERATION", "INTERSECTION_CONCATENATE")
y.SetAttValue("ROUND", "0")
y.SetAttValue("ADDVALUE", "0")
y.SetAttValue("WEIGHTBYINTERSECTIONAREASHARE", "0")
y.SetAttValue("CONCATMAXLEN", "255")
y.SetAttValue("CONCATSEPARATOR", ",")

Visum.Net.IntersectAttributes(y)
