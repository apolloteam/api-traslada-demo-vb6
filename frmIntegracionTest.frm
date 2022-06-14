VERSION 5.00
Begin VB.Form frmIntegracionTest 
   Caption         =   "Test de Integración con API Traslada"
   ClientHeight    =   6555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   13575
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame 
      Height          =   4935
      Index           =   4
      Left            =   480
      TabIndex        =   40
      Top             =   480
      Width           =   5655
      Begin VB.TextBox txtCustomProperties2 
         Height          =   285
         Left            =   2160
         TabIndex        =   7
         Top             =   2760
         Width           =   1215
      End
      Begin VB.ComboBox cboCustomProperties1 
         Height          =   315
         ItemData        =   "frmIntegracionTest.frx":0000
         Left            =   2160
         List            =   "frmIntegracionTest.frx":000A
         TabIndex        =   6
         Text            =   "INGRESO"
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox txtCostCenter2 
         Height          =   285
         Left            =   2160
         TabIndex        =   9
         Top             =   3480
         Width           =   2895
      End
      Begin VB.TextBox txtCostCenter1 
         Height          =   285
         Left            =   2160
         TabIndex        =   8
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox txtContactPhone 
         Height          =   285
         Left            =   2160
         TabIndex        =   10
         Text            =   "54 11 56676942"
         Top             =   3840
         Width           =   2895
      End
      Begin VB.ComboBox cboVehicleCategoryCode 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmIntegracionTest.frx":001F
         Left            =   2160
         List            =   "frmIntegracionTest.frx":0026
         TabIndex        =   5
         Text            =   "STD"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txttxtScheduleTime 
         Height          =   285
         Left            =   2160
         TabIndex        =   4
         Text            =   "00:00"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtScheduleDate 
         Height          =   285
         Left            =   2160
         TabIndex        =   3
         Text            =   "01/01/2022"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.ComboBox cboAction 
         Height          =   315
         ItemData        =   "frmIntegracionTest.frx":002F
         Left            =   2160
         List            =   "frmIntegracionTest.frx":003C
         TabIndex        =   1
         Text            =   "Booking"
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox txtIntegratorServiceId 
         Height          =   285
         Left            =   2160
         TabIndex        =   2
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox txtSource 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   0
         Text            =   "IntegradorPool"
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hora de ingreso"
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sentido"
         Height          =   255
         Left            =   -600
         TabIndex        =   47
         Top             =   2400
         Width           =   2655
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Centro costo 2 (Cabecera)"
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Centro costo 1 (ID)"
         Height          =   375
         Left            =   120
         TabIndex        =   45
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label lblContactPhone 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono de contacto"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label lblVehicleCategoryCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Categoria de Vehículo"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Hora viaje"
         Height          =   255
         Left            =   1320
         TabIndex        =   24
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblScheduleDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha viaje"
         Height          =   255
         Left            =   1080
         TabIndex        =   23
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblAction 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Operación"
         Height          =   255
         Left            =   -600
         TabIndex        =   21
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label lblIntegratorServiceId 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. reserva Intercargo (ID)"
         Height          =   375
         Left            =   -120
         TabIndex        =   22
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblSource 
         Alignment       =   1  'Right Justify
         Caption         =   "Source"
         Height          =   255
         Left            =   960
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame 
      Height          =   1095
      Index           =   3
      Left            =   6360
      TabIndex        =   39
      Top             =   4320
      Width           =   6735
      Begin VB.TextBox txtAddress 
         Height          =   285
         Index           =   3
         Left            =   2040
         TabIndex        =   17
         Top             =   240
         Width           =   4455
      End
      Begin VB.TextBox txtPassengerName 
         Height          =   285
         Index           =   3
         Left            =   2040
         TabIndex        =   18
         Top             =   600
         Width           =   4455
      End
      Begin VB.Label lblAddress 
         Alignment       =   1  'Right Justify
         Caption         =   "Dirección/Lugar destino 3"
         Height          =   255
         Index           =   3
         Left            =   60
         TabIndex        =   33
         Top             =   300
         Width           =   1875
      End
      Begin VB.Label lblPassengerName 
         Alignment       =   1  'Right Justify
         Caption         =   "Pasajero/s destino 3 *"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   34
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Frame Frame 
      Height          =   1095
      Index           =   2
      Left            =   6360
      TabIndex        =   38
      Top             =   3120
      Width           =   6735
      Begin VB.TextBox txtAddress 
         Height          =   285
         Index           =   2
         Left            =   2040
         TabIndex        =   15
         Top             =   240
         Width           =   4455
      End
      Begin VB.TextBox txtPassengerName 
         Height          =   285
         Index           =   2
         Left            =   2040
         TabIndex        =   16
         Top             =   600
         Width           =   4455
      End
      Begin VB.Label lblAddress 
         Alignment       =   1  'Right Justify
         Caption         =   "Dirección/Lugar destino 2"
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   31
         Top             =   240
         Width           =   1875
      End
      Begin VB.Label lblPassengerName 
         Alignment       =   1  'Right Justify
         Caption         =   "Pasajero/s destino 2 *"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   32
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame Frame 
      Height          =   1095
      Index           =   1
      Left            =   6360
      TabIndex        =   37
      Top             =   1920
      Width           =   6735
      Begin VB.TextBox txtAddress 
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   13
         Top             =   240
         Width           =   4455
      End
      Begin VB.TextBox txtPassengerName 
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   14
         Top             =   600
         Width           =   4455
      End
      Begin VB.Label lblAddress 
         Alignment       =   1  'Right Justify
         Caption         =   "Dirección/Lugar destino 1"
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   29
         Top             =   240
         Width           =   1875
      End
      Begin VB.Label lblPassengerName 
         Alignment       =   1  'Right Justify
         Caption         =   "Pasajero/s destino 1 *"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   30
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Frame Frame 
      Height          =   1095
      Index           =   0
      Left            =   6360
      TabIndex        =   36
      Top             =   480
      Width           =   6735
      Begin VB.TextBox txtAddress 
         Height          =   285
         Index           =   0
         Left            =   2040
         TabIndex        =   11
         Top             =   240
         Width           =   4455
      End
      Begin VB.TextBox txtPassengerName 
         Height          =   285
         Index           =   0
         Left            =   2040
         TabIndex        =   12
         Top             =   600
         Width           =   4455
      End
      Begin VB.Label lblAddress 
         Alignment       =   1  'Right Justify
         Caption         =   "Dirección/Lugar origen"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblPassengerName 
         Alignment       =   1  'Right Justify
         Caption         =   "Pasajero/s origen *"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdIntegrar 
      Caption         =   "Integrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10920
      TabIndex        =   19
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Label Label 
      Caption         =   "* Para ingresar varios pasajeros debe separarlos con punto y coma."
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   49
      Top             =   5520
      Width           =   4935
   End
   Begin VB.Label lblRespuestaError 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   600
      TabIndex        =   44
      Top             =   5880
      Width           =   5775
   End
   Begin VB.Label lblRespuestaOk 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   6600
      TabIndex        =   43
      Top             =   5880
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "Destino/s"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   42
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Origen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   41
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label 
      Caption         =   "Datos Generales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   35
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmIntegracionTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************************************************************
' INTEGRACION TEST (API TRASLADA)
' Fecha: 30/05/2022
'****************************************************************************************************************************************

Private Sub Form_Load()
    txtScheduleDate.Text = Format(Now, "DD/MM/YYYY")
End Sub

Private Sub IntegrationTest()
    Dim TrasladaApi As New clsTrasladaApi
    Dim bOK As Boolean
    Dim sJsonRequest As String
    Dim sJsonResponse As String
    Dim oResponse As Object
    Dim sCode As String
    Dim sErrorCode As String
    Dim sServiceId As String
    Dim sSource As String, sIntegratorServiceId As String, sScheduleDate As String, sAction As String
    Dim sContactPhone As String, sVehicleCategoryCode As String
    Dim sDescripError As String
    Dim sCostCenter1 As String, sCostCenter2 As String
    Dim sCustomProperties1 As String, sCustomProperties2 As String
    Dim sAddress_0 As String, sPassengerName_0 As String
    Dim sAddress_1 As String, sPassengerName_1 As String
    Dim sAddress_2 As String, sPassengerName_2 As String
    Dim sAddress_3 As String, sPassengerName_3 As String
    Dim sMensaje As String
    
    ' Datos Generales.
    sSource = txtSource.Text
    sIntegratorServiceId = txtIntegratorServiceId.Text
    sScheduleDate = Format(txtScheduleDate.Text, "yyyy-MM-dd") & "T" & Format(txttxtScheduleTime.Text, "HH:mm") ' Formato: yyyy-MM-ddTHH:mm
    sAction = cboAction.Text
    sContactPhone = txtContactPhone.Text
    sVehicleCategoryCode = cboVehicleCategoryCode.Text
    sCostCenter1 = txtCostCenter1.Text
    sCostCenter2 = txtCostCenter2.Text
    sCustomProperties1 = cboCustomProperties1.Text
    sCustomProperties2 = txtCustomProperties2.Text
    
    ' Origen.
    sAddress_0 = txtAddress(0).Text
    sPassengerName_0 = txtPassengerName(0).Text
    
    ' Destino 1.
    sAddress_1 = txtAddress(1).Text
    sPassengerName_1 = txtPassengerName(1).Text
    
    ' Destino 2.
    sAddress_2 = txtAddress(2).Text
    sPassengerName_2 = txtPassengerName(2).Text
    
    ' Destino 3.
    sAddress_3 = txtAddress(3).Text
    sPassengerName_3 = txtPassengerName(3).Text
    
    ' Request API.
    bOK = TrasladaApi.PostBookingServicePlane(g_sBaseApiURL, g_sApiKey, sJsonRequest, sJsonResponse, oResponse, _
                                                sSource, sIntegratorServiceId, sScheduleDate, sAction, sContactPhone, sVehicleCategoryCode, _
                                                sCostCenter1, sCostCenter2, sCustomProperties1, sCustomProperties2, _
                                                sAddress_0, sPassengerName_0, _
                                                sAddress_1, sPassengerName_1, _
                                                sAddress_2, sPassengerName_2, _
                                                sAddress_3, sPassengerName_3)
    ' Evalua respuesta.
    If bOK Then
        ' OK.
        sServiceId = oResponse.Item("value")
        sMensaje = "Respuesta de API OK. ServiceId: " & sServiceId
        
        lblRespuestaOk.Caption = sMensaje
    Else
        ' Con error.
        If Not oResponse Is Nothing Then
            sCode = oResponse.Item("code")
            sErrorCode = oResponse.Item("errorCode")
            sDescripError = oResponse.Item("descrip")
            sMensaje = "ERROR API. Code: " & sCode & " " & sDescripError & vbCrLf & "ErrorCode: " & sErrorCode
        Else
            sMensaje = "Error desconocido"
        End If
        
        lblRespuestaError.Caption = sMensaje
    End If
    
End Sub

Private Sub cmdIntegrar_Click()
    IntegrationTest
End Sub



