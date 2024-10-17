# Datum der letzten Änderung: 16.10.2024
# Universität Stuttgart
# Lehrstuhl Verkehrsplanung und Verkehrsleittechnik
# Autor Python: Kea Seelhorst

modus = Visum.Net.DemandSegments.ItemByKey("VFS0").AttValue("Mode\TSysSet")

#Widerstandsfunktion für Pkw anpassen

imp_fcn_car = Visum.Procedures.Functions.ImpedanceFunctions(modus)
imp_fcn_car_link = imp_fcn_car.ImpedanceFunction("LINK")
imp_fcn_car_connector = imp_fcn_car.ImpedanceFunction("CONNECTOR")
imp_fcn_car_turn = imp_fcn_car.ImpedanceFunction("TURN")
imp_fcn_car_mainturn = imp_fcn_car.ImpedanceFunction("MAINTURN")

#erste Zeile Link
imp_fcn_car_link.ImpedanceFunctionTerm(1).SetAttValue("FIRSTATTRCOEFF", "1")
imp_fcn_car_link.ImpedanceFunctionTerm(1).SetAttValue("FIRSTATTRNAME", "T0_PRTSYS(P)")
imp_fcn_car_link.ImpedanceFunctionTerm(1).SetAttValue("OPERATION", "MULTIPLY")
imp_fcn_car_link.ImpedanceFunctionTerm(1).SetAttValue("SECONDATTRCOEFF", "1")
imp_fcn_car_link.ImpedanceFunctionTerm(1).SetAttValue("SECONDATTRNAME", "RIN_WIDFAK")

#zweite Zeile Link
imp_fcn_car_link.ImpedanceFunctionTerm(2).SetAttValue("FIRSTATTRCOEFF", "0")
imp_fcn_car_link.ImpedanceFunctionTerm(2).SetAttValue("FIRSTATTRNAME", "T0_PRTSYS(P)")
imp_fcn_car_link.ImpedanceFunctionTerm(2).SetAttValue("OPERATION", "MULTIPLY")
imp_fcn_car_link.ImpedanceFunctionTerm(2).SetAttValue("SECONDATTRCOEFF", "0")
imp_fcn_car_link.ImpedanceFunctionTerm(2).SetAttValue("SECONDATTRNAME", "RIN_VFS")

#dritte Zeile Link
imp_fcn_car_link.ImpedanceFunctionTerm(3).SetAttValue("FIRSTATTRCOEFF", "0")
imp_fcn_car_link.ImpedanceFunctionTerm(3).SetAttValue("FIRSTATTRNAME", "LENGTH")
imp_fcn_car_link.ImpedanceFunctionTerm(3).SetAttValue("OPERATION", "NONE")
imp_fcn_car_link.ImpedanceFunctionTerm(3).SetAttValue("SECONDATTRCOEFF", "0")
imp_fcn_car_link.ImpedanceFunctionTerm(3).SetAttValue("SECONDATTRNAME", "1.0")

#Connector
imp_fcn_car_connector.ImpedanceFunctionTerm(1).SetAttValue("FIRSTATTRCOEFF", "1")
imp_fcn_car_connector.ImpedanceFunctionTerm(1).SetAttValue("FIRSTATTRNAME", "T0_TSYS(P)")
imp_fcn_car_connector.ImpedanceFunctionTerm(1).SetAttValue("OPERATION", "MULTIPLY")
imp_fcn_car_connector.ImpedanceFunctionTerm(1).SetAttValue("SECONDATTRCOEFF", "1")
imp_fcn_car_connector.ImpedanceFunctionTerm(1).SetAttValue("SECONDATTRNAME", "1.0")

#turn
imp_fcn_car_turn.ImpedanceFunctionTerm(1).SetAttValue("FIRSTATTRCOEFF", "1")
imp_fcn_car_turn.ImpedanceFunctionTerm(1).SetAttValue("FIRSTATTRNAME", "T0_PRTSYS(P)")
imp_fcn_car_turn.ImpedanceFunctionTerm(1).SetAttValue("OPERATION", "MULTIPLY")
imp_fcn_car_turn.ImpedanceFunctionTerm(1).SetAttValue("SECONDATTRCOEFF", "1")
imp_fcn_car_turn.ImpedanceFunctionTerm(1).SetAttValue("SECONDATTRNAME", "1.0")

#mainturn
imp_fcn_car_mainturn.ImpedanceFunctionTerm(1).SetAttValue("FIRSTATTRCOEFF", "1")
imp_fcn_car_mainturn.ImpedanceFunctionTerm(1).SetAttValue("FIRSTATTRNAME", "T0_PRTSYS(P)")
imp_fcn_car_mainturn.ImpedanceFunctionTerm(1).SetAttValue("OPERATION", "MULTIPLY")
imp_fcn_car_mainturn.ImpedanceFunctionTerm(1).SetAttValue("SECONDATTRCOEFF", "1")
imp_fcn_car_mainturn.ImpedanceFunctionTerm(1).SetAttValue("SECONDATTRNAME", "1.0")
