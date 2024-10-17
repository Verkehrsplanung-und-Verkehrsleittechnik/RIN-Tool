 # Datum der letzten Änderung: 16.10.2024
# Universität Stuttgart
# Lehrstuhl Verkehrsplanung und Verkehrsleittechnik
# Autor Python: Kea Seelhorst

#UDAs erstellen
try: 
    Visum.Net.Links.AddUserDefinedAttribute("RIN_WIDFAK","RIN_WidFak","RIN_WidFak",2,2,Ignored=False,MinVal=0.5, MaxVal=3,DefVal=1,CrossSectionLogic=None, CslIgnoreClosed=False, Formula='', SubAttr='', CanBeEmpty=None,)
    Visum.Net.Links.AddUserDefinedAttribute("RIN_VFS","RIN_VFS","RIN_VFS",1,0,Ignored=False,MinVal=None, MaxVal=None,DefVal=9,CrossSectionLogic=None, CslIgnoreClosed=False, Formula='', SubAttr='', CanBeEmpty=None,)
    Visum.Net.Links.AddUserDefinedAttribute("RIN_VFSKORR","RIN_VFSKorr","RIN_VFSKorr",1,0,Ignored=False,MinVal=None, MaxVal=None,DefVal=9,CrossSectionLogic=None, CslIgnoreClosed=False, Formula='', SubAttr='', CanBeEmpty=None,)
    Visum.Net.LinkTypes.AddUserDefinedAttribute("RIN_WIDFAK","RIN_WidFak","RIN_WidFak",2,2,Ignored=False,MinVal=None, MaxVal=None,DefVal=1,CrossSectionLogic=None, CslIgnoreClosed=False, Formula='', SubAttr='', CanBeEmpty=None,)
    Visum.Net.Zones.AddUserDefinedAttribute("RIN_Zentralität","RIN_Zentralität","RIN_Zentralität",1,0,Ignored=False,MinVal=None, MaxVal=None,DefVal=9,CrossSectionLogic=None, CslIgnoreClosed=False, Formula='', SubAttr='', CanBeEmpty=None,)
    Visum.Net.Zones.AddUserDefinedAttribute("RIN_NrZugehoerigerPOI","RIN_NrZugehoerigerPOI","RIN_NrZugehoerigerPOI",5,0,Ignored=False,MinVal=None, MaxVal=None,DefVal=None,CrossSectionLogic=None, CslIgnoreClosed=False, Formula='', SubAttr='', CanBeEmpty=None,)
except:
    pass

