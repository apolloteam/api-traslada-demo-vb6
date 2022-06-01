VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTrasladaApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************************************
' API TRASLADA
' Fecha: 30/05/2022
'****************************************************************************************************************************************


'-----------------------------------------------------------------------------------------------------------------------------------------
' Genera una reserva, modificación o cancelación de acuerdo a los datos especificados.
'
' sUrlApi:              URL del API de Traslada.
' sApiKey:              ApiKey de acceso y autenticación.
' sJsonRequest: (ref)   Se asignan los datos del diccionario generado convertido a json (para control de lo que realmente se manda al api).
' sJsonResponse: (ref)  Se asigna el json de respuesta del api.
' oResponse: (ref)      Se asigna el json de respuesta convertido a objeto de VB6 para poder leer los datos.
' sSource:              Dato requerido.
' sIntegratorServiceId: Código identificador de reserva de Intercargo.
' sScheduleDate:        Fecha del traslado (formato: yyyy-MM-ddTHH:mm).
' sAction:              Determina el tipo de acción a realizar (reserva, modificación o cancelación). Datos válidos Booking|Update|Cancelation.
' sContactPhone:        Teléfono de contacto de la reserva (el de operaciones de Intercargo).
' sVehicleCategoryCode: Tipo de vehículo ("STD").
' sCostCenter1:         Centro de costo 1 (ID: Código identificador de reserva de Intercargo).
' sCostCenter2:         Centro de costo 2 (Cabecera).
' sCustomProperties1:   Propiedad personalizada 1 (Sentido: "INGRESO", "EGRESO").
' sCustomProperties2:   Propiedad personalizada 2 (Hora de ingreso. Formato: HH:mm).
' sAddress_0:           Dirección de origen.
' sLocality_0:          Localidad de la dirección de origen (dejar en blanco si está incluida en sAdress_0).
' sPassengerName_0:     Nombre del pasajero o pasajeros separados por ";" del origen.
' sAddress_1:           Dirección de destino 1.
' sLocality_1:          Localidad de la dirección de destino 1 (dejar en blanco si está incluida en sAdress_0).
' sPassengerName_1:     Nombre del pasajero o pasajeros separados por ";" del destino 1.

'-----------------------------------------------------------------------------------------------------------------------------------------
Public Function PostBookingServicePlane(sUrlApi As String, sApiKey As String, _
                                        ByRef sJsonRequest As String, ByRef sJsonResponse As String, ByRef oResponse As Object, _
                                        sSource As String, sIntegratorServiceId As String, sScheduleDate As String, _
                                        sAction As String, sContactPhone As String, sVehicleCategoryCode As String, _
                                        sCostCenter1 As String, sCostCenter2 As String, _
                                        sCustomProperties1 As String, sCustomProperties2 As String, _
                                        sAddress_0 As String, sLocality_0 As String, sPassengerName_0 As String, _
                                        sAddress_1 As String, sLocality_1 As String, sPassengerName_1 As String, _
                                        sAddress_2 As String, sLocality_2 As String, sPassengerName_2 As String, _
                                        sAddress_3 As String, sLocality_3 As String, sPassengerName_3 As String) As Boolean

On Error GoTo ErrorHandler
    
    Dim bOkApi As Boolean
    
    ' Limpia variables de retorno.
    sJsonResponse = ""
    
    ' Queries.
    Dim oQueries As New Dictionary
    Set oQueries = GetQueryString() _
                .Add("apikey", sApiKey) _
                .ToDictionary()
    
    ' Body.
    Dim dBody As New Dictionary
    Dim cRoutePoints As New Collection
    
    ' Para pegarle al API hay que crear un diccionario con cada propiedad,
    ' en caso en haya una lista de determinadad entidad, se usa una coleccion de diccionario (caso direcciones)
    ' luego este diccionario se convertira en un json.
            
    ' Datos generales.
    dBody.Add "source", sSource
    dBody.Add "serviceId", ""
    dBody.Add "integratorServiceId", sIntegratorServiceId
    dBody.Add "scheduleDate", sScheduleDate
    dBody.Add "action", sAction
    dBody.Add "contactEmail", ""
    dBody.Add "contactPhone", sContactPhone
    dBody.Add "timeRequired", 0
    dBody.Add "passengerCount", 0
    dBody.Add "luggageCount", 0
    dBody.Add "vehicleCategoryCode", sVehicleCategoryCode
    dBody.Add "costCenter1", sCostCenter1
    dBody.Add "costCenter2", sCostCenter2
    
    ' Origen.
    Dim dRoutePoints_0 As New Dictionary
    dRoutePoints_0.Add "name", ""
    dRoutePoints_0.Add "address", sAddress_0
    dRoutePoints_0.Add "address2", ""
    dRoutePoints_0.Add "locality", sLocality_0
    dRoutePoints_0.Add "passengerName", sPassengerName_0
    dRoutePoints_0.Add "passengerEmail", ""
    dRoutePoints_0.Add "passengerPhone", ""
    dRoutePoints_0.Add "passengerIdentificationNumber", ""
    dRoutePoints_0.Add "latitud", ""
    dRoutePoints_0.Add "longitude", ""
    dRoutePoints_0.Add "flightNumber", Null
    dRoutePoints_0.Add "flightAirline", Null
    cRoutePoints.Add dRoutePoints_0
    
    ' Destino 1.
    Dim dRoutePoints_1 As New Dictionary
    dRoutePoints_1.Add "name", ""
    dRoutePoints_1.Add "address", sAddress_1
    dRoutePoints_1.Add "address2", ""
    dRoutePoints_1.Add "locality", sLocality_1
    dRoutePoints_1.Add "passengerName", sPassengerName_1
    dRoutePoints_1.Add "passengerEmail", ""
    dRoutePoints_1.Add "passengerPhone", ""
    dRoutePoints_1.Add "passengerIdentificationNumber", ""
    dRoutePoints_1.Add "latitud", ""
    dRoutePoints_1.Add "longitude", ""
    dRoutePoints_1.Add "flightNumber", Null
    dRoutePoints_1.Add "flightAirline", Null
    cRoutePoints.Add dRoutePoints_1
    
    ' Destino 2.
    Dim dRoutePoints_2 As New Dictionary
    dRoutePoints_2.Add "name", ""
    dRoutePoints_2.Add "address", sAddress_2
    dRoutePoints_2.Add "address2", ""
    dRoutePoints_2.Add "locality", sLocality_2
    dRoutePoints_2.Add "passengerName", sPassengerName_2
    dRoutePoints_2.Add "passengerEmail", ""
    dRoutePoints_2.Add "passengerPhone", ""
    dRoutePoints_2.Add "passengerIdentificationNumber", ""
    dRoutePoints_2.Add "latitud", ""
    dRoutePoints_2.Add "longitude", ""
    dRoutePoints_2.Add "flightNumber", Null
    dRoutePoints_2.Add "flightAirline", Null
    cRoutePoints.Add dRoutePoints_2
    
    ' Destino 3.
    Dim dRoutePoints_3 As New Dictionary
    dRoutePoints_3.Add "name", ""
    dRoutePoints_3.Add "address", sAddress_3
    dRoutePoints_3.Add "address2", ""
    dRoutePoints_3.Add "locality", sLocality_3
    dRoutePoints_3.Add "passengerName", sPassengerName_3
    dRoutePoints_3.Add "passengerEmail", ""
    dRoutePoints_3.Add "passengerPhone", ""
    dRoutePoints_3.Add "passengerIdentificationNumber", ""
    dRoutePoints_3.Add "latitud", ""
    dRoutePoints_3.Add "longitude", ""
    dRoutePoints_3.Add "flightNumber", Null
    dRoutePoints_3.Add "flightAirline", Null
    cRoutePoints.Add dRoutePoints_3
            
    dBody.Add "routePoints", cRoutePoints
    
    ' Propiedades personalizadas.
    Dim dCustomProperties As New Dictionary
    dCustomProperties.Add "COND.INTERCARGO.SENTIDO", sCustomProperties1
    dCustomProperties.Add "COND.INTERCARGO.HORA.INGRESO", sCustomProperties2
    dBody.Add "customProperties", dCustomProperties
        
    sJsonRequest = basJSON.ToString(dBody)
    
    ' Headers.
    Dim oHeaders As Dictionary
    Set oHeaders = GetHeaders() _
                   .ContentType("application/x-www-form-urlencoded") _
                   .ToDictionary()

    ' Realiza la petición al servidor.
    Set oResponse = GetRestHelper(sUrlApi) _
                    .PostVerb("booking/services/plane", dBody, oQueries, , oHeaders)

 
    If Not oResponse Is Nothing Then
  
        Select Case GetResponseType(oResponse)
    
            Case StatusMessage
                ' Error del sistema con formato conocido.
                sJsonResponse = basJSON.ToString(oResponse)
                
                bOkApi = False
                                            
            Case UnHandled
            
                ' Error formato desconocido.
                sJsonResponse = basJSON.ToString(oResponse)
                    
                If IsUnHandledStatus200(oResponse) Then
                ' Lo tomo como un OK.
                    bOkApi = True
                Else
                    bOkApi = False
                End If
                            
            Case Expected
                ' Respuesta esperada.
                bOkApi = True
            
        End Select
    End If
    
    PostBookingServicePlane = bOkApi
    
Exit Function

ErrorHandler:

    Err.Raise Err.Number, "clsTrasladaApi:PostBookingServicePlane(" & CStr(Erl) & ")->" & Err.Source, Trim(Err.Description), Err.HelpFile, Err.HelpContext

End Function


