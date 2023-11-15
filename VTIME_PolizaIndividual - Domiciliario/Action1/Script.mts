Option Explicit

Dim timeToWait
Dim fechaEfecto, ramo, producto, agencia, canalVenta, oficina, cliente, meses, frecuencia, moneda, tipoCanal, tipoTablaInter, intermediario, intermediarioDos, polizaMatriz
Dim transaccionPoliza, valorCapital, poliza, recibo, reciboDos, fechaDesde, fechaHasta, primaTotal
Dim tipoVia, nombreVia, departamento, provincia, distrito, anioConstruccion, numeroPisos, numeroSotano, tipoEstructura, usoEdificacion, longitud, latitud
Dim tipoIdentificacion, grupoIdenticacion, giroIdentificacion, detalleGiro, categoriaConstruccion, porcentajeAsegurado, empleados, area, vigilantes
Dim condicionComer, condicionOtros, condicionBanca, condicionOtrosYBanca, comercializador, otros, banca, conceptoComer, conceptoOtros, conceptoBanca, participacionComer, participacionOtros, participacionBanca, comisionCobranza, valorComision, costoInspeccion, tramiteInspeccion, reaseguroFacultativo
Dim tipoBienAsegurado, referenciaAsegurado, valorAsegurado, primaAsegurada, valorAseguradoDos, primaAseguradaDos, valorAseguradoTres

timeToWait = 10
fechaEfecto			 	= DataTable("in_fechaEfecto", dtLocalSheet)
transaccionPoliza	 	= DataTable("in_transaccionPoliza", dtLocalSheet)
agencia				 	= DataTable("in_agencia", dtLocalSheet)
canalVenta			 	= DataTable("in_canalVenta", dtLocalSheet)
ramo				 	= DataTable("in_ramo", dtLocalSheet)
producto			 	= DataTable("in_producto", dtLocalSheet)
oficina				 	= DataTable("in_oficina", dtLocalSheet)
reaseguroFacultativo 	= DataTable("in_reaseguroFacultativo", dtLocalSheet)
cliente 				 	= DataTable("in_cliente", dtLocalSheet)
meses 				 	= DataTable("in_meses", dtLocalSheet)
frecuencia 			 	= DataTable("in_frecuencia", dtLocalSheet)
moneda			 	= DataTable("in_moneda", dtLocalSheet)
tipoCanal 			 	= DataTable("in_tipoCanal", dtLocalSheet)
tipoTablaInter		 	= DataTable("in_tipoTablaInter", dtLocalSheet)
intermediario		 	= DataTable("in_intermediario", dtLocalSheet)
intermediarioDos	 	= DataTable("in_intermediarioDos", dtLocalSheet)
polizaMatriz			 	= DataTable("in_polizaMatriz", dtLocalSheet)
valorCapital			 	= DataTable("in_valorCapital", dtLocalSheet)
tipoVia				 	= DataTable("in_tipoVia", dtLocalSheet)
nombreVia			 	= DataTable("in_nombreVia", dtLocalSheet)
departamento		 	= DataTable("in_departamento", dtLocalSheet)
provincia		 	 	= DataTable("in_provincia", dtLocalSheet)
distrito				 	= DataTable("in_distrito", dtLocalSheet)
anioConstruccion		= DataTable("in_anioConstruccion", dtLocalSheet)
numeroPisos		 	= DataTable("in_numeroPisos", dtLocalSheet)
numeroSotano		 	= DataTable("in_numeroSotano", dtLocalSheet)
tipoEstructura		 	= DataTable("in_tipoEstructura", dtLocalSheet)
usoEdificacion		 	= DataTable("in_usoEdificacion", dtLocalSheet)
longitud			 	= DataTable("in_longitud", dtLocalSheet)
latitud				 	= DataTable("in_latitud", dtLocalSheet)
tipoIdentificacion	 	= DataTable("in_tipoIdentificacion", dtLocalSheet)
grupoIdenticacion	 	= DataTable("in_grupoIdenticacion", dtLocalSheet)
giroIdentificacion	 	= DataTable("in_giroIdentificacion", dtLocalSheet)
detalleGiro			 	= DataTable("in_detalleGiro", dtLocalSheet)
categoriaConstruccion	= DataTable("in_categoriaConstruccion", dtLocalSheet)
porcentajeAsegurado	= DataTable("in_porcentajeAsegurado", dtLocalSheet)
empleados				= DataTable("in_empleados", dtLocalSheet)
area					= DataTable("in_area", dtLocalSheet)
vigilantes				= DataTable("in_vigilantes", dtLocalSheet)

comisionCobranza	 	= DataTable("in_comisionCobranza", dtLocalSheet)
condicionComer		 	= DataTable("in_condicionComer", dtLocalSheet)
condicionOtros		 	= DataTable("in_condicionOtros", dtLocalSheet)
condicionBanca		 	= DataTable("in_condicionBanca", dtLocalSheet)
condicionOtrosYBanca 	= DataTable("in_condicionOtrosYBanca", dtLocalSheet)
comercializador		 	= DataTable("in_codigoComercializador", dtLocalSheet)
otros				 	= DataTable("in_codigoOtros", dtLocalSheet)
banca				 	= DataTable("in_codigoBanca", dtLocalSheet)
conceptoComer		 	= DataTable("in_conceptoComer", dtLocalSheet)
conceptoOtros		 	= DataTable("in_conceptoOtros", dtLocalSheet)
conceptoBanca 		 	= DataTable("in_conceptoBanca", dtLocalSheet)
participacionComer   	= DataTable("in_participacionComer", dtLocalSheet)
participacionOtros	 	= DataTable("in_participacionOtros", dtLocalSheet)
participacionBanca	 	= DataTable("in_participacionBanca", dtLocalSheet)
valorComision		 	= DataTable("in_valorComision", dtLocalSheet)
costoInspeccion		 	= DataTable("in_costoInspeccion", dtLocalSheet)
tramiteInspeccion	 	= DataTable("in_tramiteInspeccion", dtLocalSheet)

tipoBienAsegurado	 	= DataTable("in_tipoBienAsegurado", dtLocalSheet)
referenciaAsegurado	 	= DataTable("in_referenciaAsegurado", dtLocalSheet)
valorAsegurado	 		= DataTable("in_valorAsegurado", dtLocalSheet)
primaAsegurada 		= DataTable("in_primaAsegurada", dtLocalSheet)
valorAseguradoDos		= DataTable("in_valorAseguradoDos", dtLocalSheet)
primaAseguradaDos		= DataTable("in_primaAseguradaDos", dtLocalSheet)
valorAseguradoTres		= DataTable("in_valorAseguradoTres", dtLocalSheet)

Sub tratamientoPolizas()
	wait 4
	Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas").Frame("fraHeader").WebList("cbeTransactio").Select transaccionPoliza
	wait 3
	Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas").Frame("fraHeader").WebEdit("tcdEffecdate").Set fechaEfecto
	wait 3
	Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_4").Frame("fraHeader").WebList("cbeOffice").Select oficina
	wait 3
	Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas").Frame("fraHeader").WebList("cbeSellchannel").Select canalVenta
	wait 3
	Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas").Frame("fraHeader").WebList("cbeBranch").Select ramo
	wait 3
	Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas").Frame("fraHeader").WebEdit("valProduct").Set producto
	wait 3
	
	If reaseguroFacultativo = "Si" Then
	Browser("Tratamiento de pólizas_4").Page("Tratamiento de pólizas_2").Frame("fraHeader").WebCheckBox("chkCertifFac").Set "ON"
	End If
	wait 3
	Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_6").Frame("fraHeader").WebRadioGroup("optType").Select "1"
	wait 3
	Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_3").Frame("fraHeader").Image("Aceptar la información").Click
	wait 3
	Browser("Errores/advertencias encontrad").Page("Errores/advertencias encontrad").Image("Aceptar la información").Click
End Sub

Sub inicioAsegurables()
	wait 7
	Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_5").Frame("TreeSequence").Link("Asegurables").Click
	wait 4
	Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas_2").Frame("fraFolder").Link("Contratante").Click
	wait 5
	Browser("Browser_3").Page("Page").Frame("fraFolder").WebEdit("tctCode").Set cliente
	wait 3
	Browser("Browser_3").Page("Page").Frame("fraFolder").Image("Aceptar la información").Click
	wait 5
	If condicionComer = "Si"  Then
		wait 3
		Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_5").Frame("fraFolder_6").Link("Comercializador").Click
		wait 3
		Browser("Errores/advertencias encontrad_2").Page("Page_2").Frame("fraFolder_2").WebEdit("tctCode").Set comercializador
		wait 1
		Browser("Errores/advertencias encontrad_2").Page("Page_2").Frame("fraFolder").WebList("cbeConcept").Select conceptoComer
		wait 1
		Browser("Errores/advertencias encontrad_2").Page("Page_2").Frame("fraFolder").WebEdit("tcnPercent").Set participacionComer
		wait 2
		Browser("Browser_3").Page("Page").Frame("fraFolder").Image("Aceptar la información").Click
	End If
	
	If condicionOtros = "Si"  Then
		wait 3
		Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_5").Frame("fraFolder_7").Link("Otros Serv Comercializadores").Click
		wait 3
		Browser("Errores/advertencias encontrad_2").Page("Page_2").Frame("fraFolder_2").WebEdit("tctCode").Set otros
		wait 1
		Browser("Errores/advertencias encontrad_2").Page("Page_2").Frame("fraFolder").WebList("cbeConcept").Select conceptoOtros
		wait 1
		Browser("Errores/advertencias encontrad_2").Page("Page_2").Frame("fraFolder").WebEdit("tcnPercent").Set participacionOtros
		wait 2
		Browser("Browser_3").Page("Page").Frame("fraFolder").Image("Aceptar la información").Click
	End If
	
	If condicionBanca = "Si"  Then
		wait 3
		Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_5").Frame("fraFolder_7").Link("Banca-Seguros").Click
		wait 3
		Browser("Errores/advertencias encontrad_2").Page("Page_2").Frame("fraFolder_2").WebEdit("tctCode").Set banca
		wait 1
		Browser("Errores/advertencias encontrad_2").Page("Page_2").Frame("fraFolder").WebList("cbeConcept").Select conceptoBanca
		wait 1
		Browser("Errores/advertencias encontrad_2").Page("Page_2").Frame("fraFolder").WebEdit("tcnPercent").Set participacionBanca
		wait 2
		Browser("Browser_3").Page("Page").Frame("fraFolder").Image("Aceptar la información").Click
	End If
	
	If condicionOtrosYBanca = "Si" Then
		wait 3
		Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_5").Frame("fraFolder_7").Link("Otros Serv Comercializadores").Click
		wait 3
		Browser("Errores/advertencias encontrad_2").Page("Page_2").Frame("fraFolder_2").WebEdit("tctCode").Set otros
		wait 1
		Browser("Errores/advertencias encontrad_2").Page("Page_2").Frame("fraFolder").WebList("cbeConcept").Select conceptoOtros
		wait 1
		Browser("Errores/advertencias encontrad_2").Page("Page_2").Frame("fraFolder").WebEdit("tcnPercent").Set participacionOtros
		wait 2
		Browser("Browser_3").Page("Page").Frame("fraFolder").Image("Aceptar la información").Click
		wait 5
		Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_5").Frame("fraFolder_7").Link("Banca-Seguros").Click
		wait 3
		Browser("Errores/advertencias encontrad_2").Page("Page_2").Frame("fraFolder_2").WebEdit("tctCode").Set banca
		wait 1
		Browser("Errores/advertencias encontrad_2").Page("Page_2").Frame("fraFolder").WebList("cbeConcept").Select conceptoBanca
		wait 5
		Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas").Frame("fraHeader").Image("Aceptar la información").Click
		wait 3
		Browser("Errores/advertencias encontrad_2").Page("Errores/advertencias encontrad").Image("Aceptar la información").Click
	End If
	
	If condicionComer = "Si" or condicionOtros = "Si" or  condicionBanca = "Si" or condicionOtrosYBanca = "Si" Then
		wait 5
	       Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas").Frame("fraHeader").Image("Aceptar la información").Click
	Else 
		wait 5
		Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas").Frame("fraHeader").Image("Aceptar la información").Click
		wait 3
		Browser("Errores/advertencias encontrad_2").Page("Errores/advertencias encontrad").Image("Aceptar la información").Click
	End If 

End Sub

Sub facturacion()	
	wait 7
	Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas_2").Frame("TreeSequence").Link("Facturación").Click
	wait 3
	Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas_2").Frame("fraFolder_3").WebEdit("tcnDuration").Set meses
	wait 3
	Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_5").Frame("fraFolder_8").WebList("cbePayFreq").Select frecuencia
	wait 3
	Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas_2").Frame("fraHeader").Image("Aceptar la información").Click
	wait 3
	Browser("Errores/advertencias encontrad_2").Page("Errores/advertencias encontrad").Image("Aceptar la información").Click
End Sub

Sub seleccionarMoneda()
	wait 7
	Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas_2").Frame("TreeSequence").Link("Monedas").Click
	If moneda = "Dólares Americanos" Then
		Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas_2").Frame("fraFolder_4").WebCheckBox("Sel").Set "OFF"
		Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas_2").Frame("fraFolder_4").WebCheckBox("Sel_2").Set "ON"
	End If
	If moneda = "Soles" Then
		Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas_2").Frame("fraFolder_4").WebCheckBox("Sel_2").Set "OFF"
		Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas_2").Frame("fraFolder_4").WebCheckBox("Sel").Set "ON"
	End If
	wait 3
	Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas_2").Frame("fraHeader").Image("Aceptar la información").Click
End Sub 

Sub intermediarios()

	If tipoTablaInter = "Comisión fija" Then
		Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_5").Frame("fraFolder_9").WebList("cbeType").Select tipoTablaInter
		wait 3
		Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_5").Frame("fraFolder_10").WebEdit("tcnPercentCF").Set valorComision
		wait 3
		Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_5").Frame("fraFolder_10").WebElement("9715").Click
	End If
	
	If Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_5").Frame("fraFolder_18").WebElement("Sel").Exist Then
		Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_5").Frame("fraFolder_18").WebCheckBox("Sel").Set "ON"
		wait 1
		Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_5").Frame("fraFolder_18").Image("Eliminar").Click
		wait 3
		Browser("Tratamiento de pólizas_4").Page("Page").Frame("fraFolder_3").Image("Aceptar la información").Click
	End If
	
	If tipoCanal = "Directos" Then	
	wait 3
	Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas_2").Frame("fraFolder_5").Image("Agregar").Click
	wait 5
	Browser("Browser_3").Page("Page").Frame("fraFolder_2").WebEdit("valIntermed").Set intermediario
	wait 3
	Browser("Browser_3").Page("Page").Frame("fraFolder_2").WebElement("Intermediarios").Click
	wait 3
	Browser("Browser_3").Page("Page").Frame("fraFolder").Image("Aceptar la información").Click
	wait 12
	Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas_2").Frame("fraHeader").Image("Aceptar la información").Click
	End If
	
	If tipoCanal = "Digital" Then	
	wait 3
	Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas_2").Frame("fraFolder_5").Image("Agregar").Click
	wait 5
	Browser("Browser_3").Page("Page").Frame("fraFolder_2").WebEdit("valIntermed").Set intermediario
	wait 3
	Browser("Browser_3").Page("Page").Frame("fraFolder_2").WebElement("Intermediarios").Click
	wait 3
	Browser("Browser_3").Page("Page").Frame("fraFolder").Image("Aceptar la información").Click
	wait 8
	Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas_2").Frame("fraFolder_5").Image("Agregar").Click
	wait 5
	Browser("Browser_3").Page("Page").Frame("fraFolder_2").WebEdit("valIntermed").Set intermediarioDos
	wait 3
	Browser("Browser_3").Page("Page").Frame("fraFolder_2").WebElement("Intermediarios").Click
	wait 3
	Browser("Browser_3").Page("Page").Frame("fraFolder").Image("Aceptar la información").Click
	wait 12
	Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas_2").Frame("fraHeader").Image("Aceptar la información").Click
	End If
	
	If tipoCanal = "Broker" Then
	wait 3
	Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas_2").Frame("fraFolder_5").Image("Agregar").Click
	wait 5
	Browser("Browser_3").Page("Page").Frame("fraFolder_2").WebEdit("valIntermed").Set intermediario
	wait 3
	Browser("Browser_3").Page("Page").Frame("fraFolder_2").WebElement("Intermediarios").Click
	wait 3
	Browser("Browser_3").Page("Page").Frame("fraFolder").Image("Aceptar la información").Click
	wait 12
	Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas_2").Frame("fraHeader").Image("Aceptar la información").Click
	End If
	
End Sub

Sub moduloPoliza()
	wait 5
	Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas_2").Frame("fraFolder_5").Image("Agregar").Click
	wait 3
	Browser("Browser_3").Page("Page").Frame("fraFolder_5").WebEdit("valModulec").Set polizaMatriz
	wait 1
	Browser("Browser_3").Page("Page").Frame("fraFolder_5").WebElement("Módulos de la póliza").Click
	wait 3
	Browser("Browser_3").Page("Page").Frame("fraFolder").Image("Aceptar la información").Click
End Sub

Sub clausulaPolizaMatriz()
	wait 4
	'Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_5").Frame("TreeSequence").Link("Clau ind/cer").Click
	'wait 3
	'Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_5").Frame("fraFolder_2").WebCheckBox("Sel").Set "ON"
	If Window("Trabajo: Microsoft​ Edge_2").Dialog("Mensaje de página web").WinButton("Aceptar").Exist Then
		Window("Trabajo: Microsoft​ Edge_2").Dialog("Mensaje de página web").WinButton("Aceptar").Click
	End If
	'Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas_2").Frame("fraHeader").Image("Aceptar la información").Click
End Sub
	
Sub irPolizaMatriz()
	wait 7
	Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas_2").Frame("fraFolder_7").Image("Agregar").Click
	wait 5
	Browser("Browser_3").Page("Page").Frame("fraFolder_3").WebEdit("valModulec").Set polizaMatriz
	wait 3
	Browser("Browser_3").Page("Page").Frame("fraFolder").Image("Aceptar la información").Click
	wait 5
	Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas_2").Frame("fraHeader").Image("Aceptar la información").Click
End Sub

Sub capitalesAsegurados()
	wait 7
	Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_5").Frame("fraFolder_5").Link("Edificio").Click
	wait 4
	Browser("Browser").Page("Page").Frame("fraFolder").WebEdit("tcnSumins_real").Set valorCapital
	wait 3
	Browser("Browser_3").Page("Page").Frame("fraFolder").Image("Aceptar la información").Click
	wait 5
	Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas_2").Frame("fraHeader").Image("Aceptar la información").Click
End Sub

Sub bienesAsegurados()
	wait 7
	Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas_2").Frame("fraFolder_7").Image("Agregar").Click
	wait 3
	Browser("Browser_3").Page("Page").Frame("fraFolder_6").WebList("cbeTabGoods").Select tipoBienAsegurado
	wait 2
	Browser("Browser_3").Page("Page").Frame("fraFolder_6").WebEdit("tctDescript").Set referenciaAsegurado
	wait 2
	Browser("Browser_3").Page("Page").Frame("fraFolder_6").WebList("cbeCurrency").Select moneda
	wait 2
	Browser("Browser_3").Page("Page").Frame("fraFolder_6").WebEdit("tcnCapital").Set valorAsegurado
	wait 2
	Browser("Browser_3").Page("Page").Frame("fraFolder_6").WebEdit("tcnPremium").Set primaAsegurada
	wait 2
	Browser("Browser_3").Page("Page").Frame("fraFolder").Image("Aceptar la información").Click
	wait 3
	Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas_2").Frame("fraHeader").Image("Aceptar la información").Click
End Sub

Sub ubicacionRiesgo()
	wait 7
	If Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_5").Frame("fraFolder_13").Link("Riesgo").Exist Then	
		Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_5").Frame("fraFolder_13").Link("Riesgo").Click
		wait 5
		Browser("Browser").Page("Page").Frame("fraFolder").WebList("valWayType").Select tipoVia
		wait 2
		Browser("Browser").Page("Page").Frame("fraFolder").WebEdit("tcnBuildYear").Set anioConstruccion
		wait 1
		Browser("Browser").Page("Page").Frame("fraFolder").WebEdit("tcnFloorNumber").Set numeroPisos
		wait 1
		Browser("Browser").Page("Page").Frame("fraFolder").WebEdit("tcnBaseNumber").Set numeroSotano
		wait 2
		Browser("Browser").Page("Page").Frame("fraFolder").WebEdit("valBuildMaterial").Set tipoEstructura
		wait 2
		Browser("Browser").Page("Page").Frame("fraFolder").WebEdit("valBuildUse").Set usoEdificacion
		wait 1
		Browser("Browser").Page("Page").Frame("fraFolder").WebEdit("tctLongitude").Set longitud
		Browser("Browser").Page("Page").Frame("fraFolder").WebEdit("tctLatitude").Set latitud
		wait 4
		Browser("Browser_3").Page("Page").Frame("fraFolder").Image("Aceptar la información").Click
		wait 4
		Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas_2").Frame("fraHeader").Image("Aceptar la información").Click
	Else
		Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas_2").Frame("fraFolder_7").Image("Agregar").Click
		wait 5
		Browser("Browser").Page("Page").Frame("fraFolder").WebList("valWayType").Select tipoVia
		wait 2
		Browser("Browser").Page("Page").Frame("fraFolder").WebEdit("tctAddress").Set nombreVia
		wait 2
		Browser("Browser").Page("Page").Frame("fraFolder").WebEdit("valGeographiczone1").Set departamento
		wait 4
		Browser("Browser").Page("Page").Frame("fraFolder").WebEdit("valGeographiczone2").Set provincia
		wait 4
		Browser("Browser").Page("Page").Frame("fraFolder").WebEdit("valGeographiczone3").Set distrito
		wait 4
		Browser("Browser").Page("Page").Frame("fraFolder").WebEdit("tcnBuildYear").Set anioConstruccion
		wait 1
		Browser("Browser").Page("Page").Frame("fraFolder").WebEdit("tcnFloorNumber").Set numeroPisos
		wait 1
		Browser("Browser").Page("Page").Frame("fraFolder").WebEdit("tcnBaseNumber").Set numeroSotano
		wait 1
		Browser("Browser_2").Page("Page").Frame("fraFolder").WebEdit("valBuildMaterial").Set tipoEstructura
		wait 2
		Browser("Browser_2").Page("Page").Frame("fraFolder").WebEdit("valBuildUse").Set usoEdificacion
		wait 1
		Browser("Browser").Page("Page").Frame("fraFolder").WebEdit("tctLongitude").Set longitud
		Browser("Browser").Page("Page").Frame("fraFolder").WebEdit("tctLatitude").Set latitud
		wait 3
		Browser("Browser_3").Page("Page").Frame("fraFolder").Image("Aceptar la información").Click
		wait 5
		Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas_2").Frame("fraHeader").Image("Aceptar la información").Click
	End If
End Sub

Sub datosParticularesDomiciliario()
	wait 7
	Browser("Tratamiento de pólizas").Page("Page").Frame("fraFolder").WebList("cbeBusinessTy").Select tipoIdentificacion
	wait 3
	Browser("Tratamiento de pólizas").Page("Page").Frame("fraFolder").WebEdit("valCommerGrp").Set grupoIdenticacion
	wait 4	
	Browser("Tratamiento de pólizas").Page("Page").Frame("fraFolder").WebEdit("valCodKind").Set giroIdentificacion
	wait 4
	Browser("Tratamiento de pólizas").Page("Page").Frame("fraFolder").WebEdit("tctDescBussi").Set detalleGiro
	Browser("Tratamiento de pólizas").Page("Page").Frame("fraFolder").WebEdit("valConstCat").Set categoriaConstruccion
	wait 3
	'Browser("Tratamiento de pólizas").Page("Page").Frame("fraFolder").WebList("cbeSpCombType").Select tipoRiesgo
	'Browser("Tratamiento de pólizas").Page("Page").Frame("fraFolder").WebList("cbeBuildType").Select tipoConstruccionIdentificacion
	'Browser("Tratamiento de pólizas").Page("Page").Frame("fraFolder").WebList("cbeSeismicZone").Select zonaSismica
	'Browser("Tratamiento de pólizas").Page("Page").Frame("fraFolder").WebList("cbeSideCloseType").Select tipoCerramiento
	'Browser("Tratamiento de pólizas").Page("Page").Frame("fraFolder").WebList("cbeRoofType").Select tipoTrecho
	
	Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_5").Frame("fraFolder_21").WebEdit("tcnInsured").Set porcentajeAsegurado
	Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_5").Frame("fraFolder_21").WebEdit("tcnEmployees").Set empleados
	Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_5").Frame("fraFolder_21").WebEdit("tcnArea").Set area
	Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_5").Frame("fraFolder_21").WebEdit("tcnVigilance").Set vigilantes
	wait 3
	
	If reaseguroFacultativo = "Si"  Then
		Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_5").Frame("fraFolder_3").Image("Registro de Costos de").Click
		wait 3
		Browser("Errores/advertencias encontrad_2").Page("Page_3").Image("Agregar").Click
		wait 3
		Browser("Browser").Page("Page").Frame("fraFolder").WebTable("Fecha efecto").WebList("name:=cbeCurrency").Select moneda
		wait 1
		Browser("Browser").Page("Page_2").Frame("fraFolder_3").WebEdit("tcnCosto").Set costoInspeccion
		wait 1
		Browser("Browser").Page("Page_2").Frame("fraFolder_3").WebEdit("tcttramite").Set tramiteInspeccion
		wait 2
		Browser("Browser").Page("Page_2").Frame("fraFolder_3").Image("Aceptar la información").Click
		wait 3
		Dim mySendKeys
		set mySendKeys = CreateObject("WScript.shell")
		mySendKeys.SendKeys("^") + "{W}"
		wait 2
		Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas").Frame("fraHeader").Image("Aceptar la información").Click
	Else
	Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas").Frame("fraHeader").WebTable("WebTable").Image("Aceptar la información").Click
	End If
End Sub

Sub coberturaPoliza()
	wait 8
	'Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_5").Frame("fraFolder_4").Link("Incendio y/o Rayo").Click
	'wait 4
	'Browser("Browser_3").Page("Page").Frame("fraFolder_7").WebEdit("tcnCapital").Set valorAseguradoDos
	'wait 2
	'Browser("Browser_3").Page("Page").Frame("fraFolder_7").WebEdit("tcnPremium").Set primaAseguradaDos
	'wait 3
	'Browser("Browser_3").Page("Page").Frame("fraFolder").Image("Aceptar la información").Click
	'wait 3
	'Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_5").Frame("fraFolder_19").Link("Explosión").Click
	'wait 4
	'Browser("Browser_3").Page("Page").Frame("fraFolder_7").WebEdit("tcnCapital").Set valorAseguradoTres
	'wait 3
	'Browser("Browser_3").Page("Page").Frame("fraFolder").Image("Aceptar la información").Click
	'wait 4
	Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas").Frame("fraHeader").WebTable("WebTable").Image("Aceptar la información").Click
End Sub

Sub impPolizaIndividual()
	wait 7
	'If reaseguroFacultativo = "Si"  Then
	'	Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_5").Frame("TreeSequence").Link("Rec/Desc/Imp").Click
	'	wait 3
	'	Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_5").Frame("fraFolder_14").Link("COMISION POR COBRANZA").Click
	'	wait 3
	'	Browser("Errores/advertencias encontrad_2").Page("Page_2").Frame("fraFolder_8").WebEdit("tcnAmount").Set comisionCobranza
	'	wait 3
	'	Browser("Browser_3").Page("Page").Frame("fraFolder").Image("Aceptar la información").Click
	'	wait 5
	'	Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas").Frame("fraHeader").Image("Aceptar la información").Click
	'	wait 3
	'	Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_5").Frame("TreeSequence").Link("Reasegur").Click
	'If moneda = "Dólares Americanos" Then
	'	Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_5").Frame("fraFolder_16").Link("US$").Click
	'Else 
	'	Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_5").Frame("fraFolder_16").Link("S/").Click
	'End If
	'	Browser("Errores/advertencias encontrad_2").Page("Page_2").Frame("fraFolder_6").WebList("cbeChange").Select "Facultativo"
	'	wait 2
	'	Browser("Browser_3").Page("Page").Frame("fraFolder").Image("Aceptar la información").Click
	'	wait 6
	'	Dim valorFacultativo
	'	valorFacultativo = Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_5").Frame("fraFolder_20").WebTable("Módulo").GetCellData(2,13)		
	'	wait 4
	'	Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_5").Frame("fraFolder_17").Image("Agregar").Click
	'	wait 3
	'	Browser("Errores/advertencias encontrad_2").Page("Page_2").Frame("fraFolder_9").WebEdit("cbeFacultty").Set "4"
	'	wait 3
	'	Browser("Errores/advertencias encontrad_2").Page("Page_2").Frame("fraFolder_9").WebEdit("cbeCompany").Set "12"
	'	wait 3
	'	Browser("Errores/advertencias encontrad_2").Page("Page_2").Frame("fraFolder_9").WebEdit("tcnParticip").Set valorFacultativo
	'	wait 3
	'	Browser("Errores/advertencias encontrad_2").Page("Page_2").Frame("fraFolder_9").WebEdit("tcnComission").Set "10"
	'	wait 3
	'	Browser("Errores/advertencias encontrad_2").Page("Page_2").Frame("fraFolder_9").WebList("valCorredor").Select "AIG EUROPE LIMITED"
	'	wait 3
	'	Browser("Browser_3").Page("Page").Frame("fraFolder").Image("Aceptar la información").Click
	'	wait 6
	'	Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas").Frame("fraHeader").WebTable("WebTable").Image("Aceptar la información").Click
	'Else
	'Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas").Frame("fraHeader").WebTable("WebTable").Image("Aceptar la información").Click
	'wait 4
	Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas").Frame("fraHeader").WebTable("WebTable").Image("Aceptar la información").Click
	'End If
End Sub

Sub informacionRecibo()
	'wait 3		
	'Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas_5").Frame("TreeSequence").Link("Recibo(s)").Click
	wait 3
	poliza = Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas_2").Frame("fraHeader").WebTable("Ramo/Producto").GetCellData(1,4)
	DataTable("out_poliza", dtLocalSheet) =  poliza
	
	recibo = Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas").Frame("fraFolder_4").WebTable("Recibo(s)").GetCellData(1,2)
	DataTable("out_reciboUno", dtLocalSheet) =  recibo
		
	fechaDesde = Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas").Frame("fraFolder_4").WebTable("Recibo(s)").GetCellData(4,5)
	DataTable("out_fechaDesde", dtLocalSheet) =  fechaDesde
	
	fechaHasta = Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas").Frame("fraFolder_4").WebTable("Recibo(s)").GetCellData(5,6)
	DataTable("out_fechaHasta", dtLocalSheet) =  fechaHasta
	
	primaTotal = Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas").Frame("fraFolder_4").WebTable("Prima Total").GetCellData(1,2)
	DataTable("out_primaTotal", dtLocalSheet) =  primaTotal
	
	wait 3
	Browser("Tratamiento de pólizas").Page("Tratamiento de pólizas").Frame("fraHeader").WebTable("WebTable").Image("Aceptar la información").Click	
	wait 4
	Browser("Tratamiento de pólizas_2").Page("Tratamiento de pólizas").Frame("fraHeader").Image("Finaliza la ejecución").Click
	wait 3
	Browser("Browser_3").Page("Page").Frame("fraFolder").Image("Aceptar la información").Click
End Sub

Call tratamientoPolizas()
Call inicioAsegurables()
Call facturacion()
Call seleccionarMoneda()
Call irPolizaMatriz()
Call capitalesAsegurados()
Call ubicacionRiesgo()
Call datosParticularesDomiciliario()
Call Intermediarios()
'Call bienesAsegurados()
Call coberturaPoliza()
Call impPolizaIndividual()
Call clausulaPolizaMatriz()
Call informacionRecibo()
validarRegistroPolizaBD(poliza)
