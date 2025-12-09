VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Dome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "RRCI Driver"
'
'   =============
'   Dome.cls
'   =============
'
' Written:  10-Oct-2020   Chuck Faranda ccd@ccdastro.net
' Provide AS IS
' Edits:
'
' When         Who     What
' -----------  ---     --------------------------------------------------
' 10-Oct-2020  cf     Initial edit
' -----------------------------------------------------------------------------

' This module operates as a multiplexing ASCOM Roof Controller

Option Explicit
Private Util As Object

Public Enum scopestate
scopesafe = 0
scopenotsafe = 1
End Enum

Public Enum roofState
    roofopen = 0
    roofclosed = 1
    roofopening = 2
    roofclosing = 3
    roofError = 4
    roofunknown = 5
End Enum

Implements IDome

Private m_Util As DriverHelper.Util

Private ThisObjectConnected As Boolean

Private Sub Class_Initialize()

    Set m_Util = New DriverHelper.Util
    
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
    If ThisObjectConnected Then ThisObjectConnected = False
        
        If RoofConnectedCount = 0 Then
           
        If m_Serial.Connected Then m_Serial.Connected = False
           Unload frmSetUp
           Unload frmRoof
           Set m_Util = Nothing
           If Not g_bRunExecutable Then Set g_ControllerProfile = Nothing
           Set m_Roof = Nothing
           Set m_Serial = Nothing
        End If
  If (RoofConnectedCount < 1 Or (DoSetup And RoofConnectedCount = 1)) Then 'Or (RoofConnectedCount = 1 And RoofIsConnected = False) Then
      
      Unload frmRoof
         
     Set m_Roof = Nothing
      '
      End
  End If
End Sub

Public Sub CloseShutter()

Dim buf As String
Dim temp As String

On Error GoTo ErrorMsg
MoveTimer = Now
m_Util.WaitForMilliseconds (300)

If g_ControllerProfile.GetValue(ID, "Auto Close", "", "") = "1" Then
   frmRoof.Text3.Text = "Auto Close Disabled"
   Beep
   Exit Sub
End If
 
 m_Util.WaitForMilliseconds (1000)
 
   If Not g_scopestatus = scopesafe And frmSetUp.chkSafe = Checked Then
      AbortSlew
      Exit Sub
   End If
    
If 1 = g_ControllerProfile.GetValue(ID, "Scope Safe", "", "") Then
    If m_Safe Then
    If g_roofstatus = roofclosed Then
      If g_bRunExecutable Then
         frmRoof.Text3.Text = "Roof is already closed"
      
      End If
        Exit Sub
    End If
    
    g_roofstatus = roofclosing
    RoofIsMoving = True
    counta = 0
    
       Me.Roof1
    
    
Else
   frmRoof.Text3.Text = "Can't close scope unsafe"
End If

ElseIf 0 = g_ControllerProfile.GetValue(ID, "Scope Safe", "", "") Then
  If g_roofstatus = roofclosed Then
      If g_bRunExecutable Then
         frmRoof.Text3.Text = "Roof is already closed"
      
      End If
        Exit Sub
    End If
    
    g_roofstatus = roofclosing
    RoofIsMoving = True
    counta = 0
    
       Me.Roof1
   
    
End If

Exit Sub

ErrorMsg:
g_roofstatus = roofError
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Property Get Connected() As Boolean
    
    On Error GoTo ErrorMsg
    
    Connected = RoofIsConnected
    Exit Property

ErrorMsg:
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Property

Public Property Let Connected(ByVal RHS As Boolean)


Dim buf As String, buf2 As String
    
    Dim tmo As Long
   On Error GoTo NO_SCOPE

If RHS Then ' CONNECTING
If RoofConnectedCount >= 1 Or (g_bRunExecutable And Me.Connected) Then GoTo Final
     '
        ' (1) Set up the communications link
        '
        
        On Error GoTo NO_SCOPE                          ' -- error trap --
      
        
        buf = g_ControllerProfile.GetValue(ID, "Com Port", "", "")
       
        If buf = "" Then                                ' No Com port
            GoTo NO_SCOPE
           
        End If
         
         m_Serial.PortName = buf ' Set port
         m_Serial.Speed = ps9600 'RRCI uses 9600 ps115200
         m_Serial.Connected = True                       ' Grab the serial port
         m_Serial.ReceiveTimeout = 5                  ' 2 second timeout
         'm_Serial.DTREnable = True
         'm_Serial.RTSEnable = False
         m_Serial.Transmit ("Leo#")
         
        ' (2) Determine that there is a roof controller there, initialize
        buf2 = m_Serial.ReceiveTerminated("#")
        
        
        On Error GoTo 0                                 ' -- end trap --
    If InStr(buf2, "RRCI") Then
    
      ThisObjectConnected = True
      RoofIsConnected = True
      
      If g_ControllerProfile.GetValue(ID, "Show Main", "", "") = 1 Then
         frmRoof.Show
      Else
         Load frmRoof
      End If
      
      frmRoof.chkLink.Caption = "UnLink"
      frmRoof.chkLink.BackColor = vbGreen
      
     
        If Not frmSetUp.Visible Then RoofConnectedCount = RoofConnectedCount + 1
        frmRoof.lblConnection.Caption = "Clients Connected " & RoofConnectedCount
      
      m_Util.WaitForMilliseconds (1000)
      
      Me.CommandBlind g_ControllerProfile.GetValue(ID, "Altitude", "", "0") & "Alt" & "#"
      m_Serial.ReceiveTimeout = 5                  ' 5 second timeout
      frmRoof.Timer1.Enabled = True
   Else
      frmRoof.LED2.FillColor = vbBlack
      m_Serial.Connected = False                      ' Release COM port
      ThisObjectConnected = False
      RoofIsConnected = False
      frmRoof.lblConnection.Caption = "Clients Connected " & RoofConnectedCount
      frmRoof.Shape1.FillColor = vbBlack
      frmRoof.Label1.Caption = "Scope Safe ?"
     frmRoof.Timer1.Enabled = False
      GoTo NO_SCOPE
      Exit Property
    
   End If
    
Else ' DISCONNECTING
frmRoof.Timer1.Enabled = False
        Me.CommandBlind "stop"
        RoofConnectedCount = RoofConnectedCount - 1
      
        frmRoof.lblConnection.Caption = "Clients Connected " & RoofConnectedCount
        
        
        On Error Resume Next                            ' Best efforts...
      
        If RoofConnectedCount <= 0 Then
           m_Serial.ClearBuffers                           ' Clear serial buffers
           m_Serial.Connected = False                      ' Release COM port
           RoofIsConnected = False
           ThisObjectConnected = False
           If g_bRunExecutable Then
           frmRoof.chkLink.Caption = "Link"
           frmRoof.chkLink.BackColor = &H8000000F
           frmRoof.chkLink.Value = 0
           frmRoof.Timer1.Enabled = False
           frmRoof.Shape1.FillColor = vbBlack
           frmRoof.Label1.Caption = "Scope Safe ?"
           frmRoof.Text3.Text = "Roof Status ?"
        
        End If
        
        
    End If
    
End If

 Exit Property

Final:
   If Not frmSetUp.Visible Then RoofConnectedCount = RoofConnectedCount + 1
   RoofIsConnected = True
   frmRoof.lblConnection.Caption = "Clients Connected " & RoofConnectedCount
   

    Exit Property

NO_SCOPE:
  frmRoof.Timer1.Enabled = False
  frmRoof.Text3.Text = "Roof Status ?"
  frmRoof.chkLink.Value = 0
  ThisObjectConnected = False

    m_Serial.ClearBuffers
                             ' Release the port
    RoofIsConnected = False
    If Err.Number = 8005 Or Err.Number = 8012 Then
     MsgBox Err.Number & Err.Source & Err.Description & Err.HelpFile & Err.HelpContext, vbSystemModal
    Else
     MsgBox MSG_NO_SCOPE, vbSystemModal
    End If


End Property

Public Property Get Altitude() As Double
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property Altitude" & MSG_NOT_IMPLEMENTED
End Property
Public Property Get AtHome() As Boolean
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property AtHome" & MSG_NOT_IMPLEMENTED
End Property

Public Property Get AtPark() As Boolean
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property AtPark" & MSG_NOT_IMPLEMENTED
End Property

Public Property Get Azimuth() As Double
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property Azimuth" & MSG_NOT_IMPLEMENTED
End Property

Public Property Get CanFindHome() As Boolean

   CanFindHome = False
                
End Property

Public Property Get CanPark() As Boolean
    
               CanPark = False
End Property

Public Property Get CanSetAltitude() As Boolean
    
                CanSetAltitude = False
End Property

Public Property Get CanSetAzimuth() As Boolean
    
                CanSetAzimuth = False
End Property

Public Property Get CanSetPark() As Boolean
    
                CanSetPark = False
End Property
Public Property Get CanSetShutter() As Boolean
   
                CanSetShutter = True
                
End Property

Public Property Get CanSyncAzimuth() As Boolean
   
                CanSyncAzimuth = False
End Property

Public Property Get CanSlave() As Boolean

    CanSlave = False
    
      
End Property

Public Property Get Description() As String
    
    On Error Resume Next
    
    Description = "Ascom Roll-Off Roof Observatory Controller"
    
End Property

Public Property Get DriverInfo() As String
    
    On Error Resume Next
    
    DriverInfo = App.FileDescription & " " & _
                 App.Major & "." & App.Minor & "." & App.Revision
    
    If App.CompanyName <> "" Then _
        DriverInfo = DriverInfo & vbCrLf & App.CompanyName
    If App.LegalCopyright <> "" Then _
        DriverInfo = DriverInfo & vbCrLf & App.LegalCopyright
    If App.Comments <> "" Then _
        DriverInfo = DriverInfo & vbCrLf & App.Comments
    
End Property
Public Sub FindHome()
        Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "FindHome" & MSG_NOT_IMPLEMENTED
    End Sub

Public Property Get InterfaceVersion() As Integer
    On Error Resume Next
    InterfaceVersion = App.Major
    Exit Property


End Property

Public Property Get Name() As String
    
    On Error Resume Next
    
    Name = "RRCI Roof Controller"
    

End Property
Public Sub SOn()
Me.CommandBlind "on"
End Sub
Public Sub SOff()
Me.CommandBlind "off"
End Sub

Public Sub OpenShutter()

Dim buf As String

On Error GoTo ErrorMsg
 
 MoveTimer = Now
 m_Util.WaitForMilliseconds (300)
 g_roofstatus = roofopening
 frmRoof.Text3.Text = ""
 
If g_ControllerProfile.GetValue(ID, "Scope Safe", "", "") = 1 Then ' added for scope sensor present


If g_ControllerProfile.GetValue(ID, "Auto Close", "", "") = "1" Then
   frmRoof.Text3.Text = "Auto Close Disabled"
   Beep
   Exit Sub
End If

'Me.CommandBlind "on"
'm_Util.WaitForMilliseconds (1000)

   If Not g_scopestatus = scopesafe And frmSetUp.chkSafe = Checked Then
      AbortSlew
      Exit Sub
   End If

  If m_Safe Then
    If g_roofstatus = roofopen Then
    
      If g_bRunExecutable Then
         MsgBox "Roof is already open"
      
      End If
         Exit Sub
   End If
 
     g_roofstatus = roofopening

  
    RoofIsMoving = True
    counta = 0
  
           
       Me.Roof2
   
 Else
   frmRoof.Text3.Text = "Can't Open, Raining or Scope unsafe"

 End If
 
ElseIf g_ControllerProfile.GetValue(ID, "Scope Safe", "", "") = 0 Then
    
    If g_roofstatus = roofopen Then
    
      If g_bRunExecutable Then
         MsgBox "Roof is already open"
      
      End If
         Exit Sub
   End If
 
     g_roofstatus = roofopening

  
  RoofIsMoving = True
  counta = 0
  
       m_Util.WaitForMilliseconds (100)
    
       Me.Roof2
    


End If
 



Exit Sub

ErrorMsg:
g_roofstatus = roofError
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub Park()
        Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Park" & MSG_NOT_IMPLEMENTED
    End Sub
Public Property Get Slaved() As Boolean
Slaved = False
End Property
Public Property Let Slaved(ByVal Value As Boolean)
Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property Slaved" & MSG_NOT_IMPLEMENTED
End Property
Public Sub SlewToAltitude(ByVal Altitude As Double)
        ' TODO Replace this with your implementation
       Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "SlewToAltitude " & MSG_NOT_IMPLEMENTED
    End Sub

    Public Sub SlewToAzimuth(ByVal Azimuth As Double)
        Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "SlewToAltitude" & MSG_NOT_IMPLEMENTED
    End Sub
    
Public Sub SyncToAzimuth(ByVal Azimuth As Double)
Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "SyncToAzimuth" & MSG_NOT_IMPLEMENTED
End Sub



Public Sub AbortSlew()
Dim opt1, opt2, opt2A, opt3 As Boolean

On Error Resume Next
opt1 = CBool(g_ControllerProfile.GetValue(ID, "1 Button", "", ""))
opt2 = CBool(g_ControllerProfile.GetValue(ID, "2 Button", "", ""))
opt3 = CBool(g_ControllerProfile.GetValue(ID, "3 Button", "", ""))

    
    On Error GoTo ErrorMsg
   counta = 61
 
  
 
     frmRoof.Text3.Text = "Roof Stopped"
  

AbortMove = True
RoofIsMoving = False

   Me.CommandBlind "x"

   Me.CommandBlind "y"
MoveTimer = Now
    Exit Sub

ErrorMsg:
g_roofstatus = roofError
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub


Public Property Get ShutterStatus() As roofState
    
On Error GoTo ErrorMsg
    m_Util.WaitForMilliseconds (200)
    ShutterStatus = g_roofstatus
    
    Exit Property
'Debug.Print ShutterStatus
ErrorMsg:
g_roofstatus = roofError

    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Property

Public Property Get Slewing() As Boolean
    If Not RoofIsConnected Then Err.Raise SCODE_NOT_CONNECTED, ERR_SOURCE, MSG_NOT_CONNECTED
    On Error GoTo ErrorMsg
    
    Slewing = RoofIsMoving
    Exit Property

ErrorMsg:
g_roofstatus = roofError
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Property
Public Sub Roof1()
Dim opt1, opt2, opt2A, opt3 As Boolean
Dim buf As String
Dim counta As Integer

If Not RoofIsConnected Then Err.Raise SCODE_NOT_CONNECTED, ERR_SOURCE, MSG_NOT_CONNECTED
On Error Resume Next

opt1 = CBool(g_ControllerProfile.GetValue(ID, "1 Button", "", ""))
opt2 = CBool(g_ControllerProfile.GetValue(ID, "2 Button", "", ""))
opt3 = CBool(g_ControllerProfile.GetValue(ID, "3 Button", "", ""))


If m_Serial.Connected Then

   Me.CommandBlind "close"
   
   g_roofstatus = roofclosing
   m_Util.WaitForMilliseconds (500)
   
   If opt1 And frmSetUp.chkHold.Value = Unchecked Then Me.CommandBlind "y"
   'Me.CommandBlind "off"
   'm_Util.WaitForMilliseconds (500)
   counta = 0
   Do Until g_roofstatus = roofclosed
   DoEvents
   m_Util.WaitForMilliseconds (500)
   counta = counta + 1
   If counta >= CInt(frmSetUp.txtTO.Text) * 2 Then Exit Do
   RoofIsMoving = True
   If AbortMove Then
      AbortMove = False
      DoEvents
      Exit Do
   End If
   Loop
   counta = 0
   If frmSetUp.chkHold.Value = Checked Then m_Roof.CommandBlind "y"
    RoofIsMoving = False
Else
g_roofstatus = roofError
      If g_bRunExecutable Then
         MsgBox "No Roof Controller connected"
      
      Else
         Err.Raise SCODE_NOT_CONNECTED, ERR_SOURCE, MSG_NOT_CONNECTED
      End If
End If

End Sub
Public Sub Roof2()
Dim counta As Integer
If Not RoofIsConnected Then Err.Raise SCODE_NOT_CONNECTED, ERR_SOURCE, MSG_NOT_CONNECTED
If m_Serial.Connected Then
  MoveTimer = Now
   
   Me.CommandBlind "open"
   g_roofstatus = roofopening
   m_Util.WaitForMilliseconds (500)
   
   If frmSetUp.chkHold.Value = Unchecked Then
      Me.CommandBlind "x"
   End If
  
   m_Util.WaitForMilliseconds (500)
   Me.CommandBlind "off"
m_Util.WaitForMilliseconds (500)
   Do Until g_roofstatus = roofopen
   DoEvents
   RoofIsMoving = True
   DoEvents
   m_Util.WaitForMilliseconds (500)
   counta = counta + 1
   If counta >= CInt(frmSetUp.txtTO.Text) * 2 Then Exit Do
   If AbortMove Then
      AbortMove = False
      Exit Do
      
   End If
   Loop
   counta = 0
   
   If frmSetUp.chkHold.Value = Checked Then
      m_Roof.CommandBlind "x"
   End If
   'If g_roofstatus = roofopen Then m_Roof.CommandBlind "x"
   RoofIsMoving = False
Else
   g_roofstatus = roofError
      If g_bRunExecutable Then
         MsgBox "No Roof Controller connected"
      
      Else
         Err.Raise SCODE_NOT_CONNECTED, ERR_SOURCE, MSG_NOT_CONNECTED
      End If
      
End If
End Sub


Public Sub SetupDialog()
DoSetup = True

If RoofIsConnected Then
   frmSetUp.Show 1
   DoSetup = False
Else
   frmSetUp.Show 1
   frmRoof.Hide
End If
    
    
End Sub

Public Sub CommandBlind(ByVal Command As String)
On Error Resume Next
    m_Serial.ClearBuffers                           ' Clear remaining junk in buffer
    m_Serial.Transmit Command & "#"
    
    
End Sub

Public Function CommandBool(ByVal Command As String) As Boolean
    Dim buf As String
    On Error Resume Next
    m_Serial.ClearBuffers                           ' Clear remaining junk in buffer
    m_Serial.Transmit Command
    buf = m_Serial.ReceiveCounted(1)                ' Just a 1 or 0
    If buf = "1" Then
        CommandBool = True
    ElseIf buf = "0" Then
        CommandBool = False
    Else
        CommandBool = True
    End If

End Function

Public Function CommandString(ByVal Command As String) As String
    Dim buf As String
   
  On Error GoTo Err
    m_Serial.ClearBuffers                           ' Clear remaining junk in buffer
    m_Serial.Transmit Command
    
    buf = m_Serial.ReceiveTerminated("#")
    Debug.Print buf
    If buf <> "" Then                   ' Overflow protection
        CommandString = Left$(buf, Len(buf) - 1)   ' Strip '#'
    Else
        CommandString = ""
    End If
  
    m_Serial.ClearBuffers
    
Exit Function
Err:

  Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext


    
End Function


Public Sub SetPark()
Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "SetPark" & MSG_NOT_IMPLEMENTED
End Sub

' =======================
' Implementation of IDome
' =======================

Private Sub IDome_AbortSlew()
    AbortSlew
End Sub

Private Property Get IDome_Altitude() As Double
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property Altitude" & MSG_NOT_IMPLEMENTED
End Property

Private Property Get IDome_AtHome() As Boolean
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property AtHome" & MSG_NOT_IMPLEMENTED
End Property

Private Property Get IDome_AtPark() As Boolean
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property AtPark" & MSG_NOT_IMPLEMENTED
End Property

Private Property Get IDome_Azimuth() As Double
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property Azimuth" & MSG_NOT_IMPLEMENTED
End Property

Private Property Get IDome_CanFindHome() As Boolean
    IDome_CanFindHome = False
End Property

Private Property Get IDome_CanPark() As Boolean
    IDome_CanPark = False
End Property

Private Property Get IDome_CanSetAltitude() As Boolean
    IDome_CanSetAltitude = False
End Property

Private Property Get IDome_CanSetAzimuth() As Boolean
    IDome_CanSetAzimuth = False
End Property

Private Property Get IDome_CanSetPark() As Boolean
    IDome_CanSetPark = False
End Property

Private Property Get IDome_CanSetShutter() As Boolean
    IDome_CanSetShutter = True
End Property

Private Property Get IDome_CanSlave() As Boolean
    IDome_CanSlave = False
End Property

Private Property Get IDome_CanSyncAzimuth() As Boolean
    IDome_CanSyncAzimuth = False
End Property

Private Sub IDome_CloseShutter()
    CloseShutter
End Sub

Private Sub IDome_CommandBlind(ByVal Command As String)
    CommandBlind (Command)
End Sub

Private Function IDome_CommandBool(ByVal Command As String) As Boolean
    IDome_CommandBool = CommandBool(Command)
End Function

Private Function IDome_CommandString(ByVal Command As String) As String
    IDome_CommandString = CommandString(Command)
End Function

Private Property Let IDome_Connected(ByVal RHS As Boolean)
    Connected = RHS
End Property

Private Property Get IDome_Connected() As Boolean
    IDome_Connected = Connected
End Property

Private Property Get IDome_Description() As String
    IDome_Description = Description
End Property

Private Property Get IDome_DriverInfo() As String
    IDome_DriverInfo = DriverInfo
End Property

Private Sub IDome_FindHome()
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property FindHome" & MSG_NOT_IMPLEMENTED
End Sub

Private Property Get IDome_InterfaceVersion() As Integer
    IDome_InterfaceVersion = InterfaceVersion
End Property

Private Property Get IDome_Name() As String
    IDome_Name = "I_" & Name
End Property

Private Sub IDome_OpenShutter()
    OpenShutter
End Sub

Private Sub IDome_Park()
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property Park" & MSG_NOT_IMPLEMENTED
End Sub

Private Sub IDome_SetPark()
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property SetPark" & MSG_NOT_IMPLEMENTED
End Sub

Private Sub IDome_SetupDialog()
    SetupDialog
End Sub

Private Property Get IDome_ShutterStatus() As AscomInterfacesLib.ShutterState
    IDome_ShutterStatus = ShutterStatus
End Property

Private Property Let IDome_Slaved(ByVal RHS As Boolean)
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property Slaved" & MSG_NOT_IMPLEMENTED
End Property

Private Property Get IDome_Slaved() As Boolean
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property Slaved" & MSG_NOT_IMPLEMENTED
End Property

Private Property Get IDome_Slewing() As Boolean
    IDome_Slewing = Slewing
End Property

Private Sub IDome_SlewToAltitude(ByVal Altitude As Double)
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property SlewToAltitude" & MSG_NOT_IMPLEMENTED
End Sub

Private Sub IDome_SlewToAzimuth(ByVal Azimuth As Double)
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property SlewToAzimuth" & MSG_NOT_IMPLEMENTED
End Sub

Private Sub IDome_SyncToAzimuth(ByVal Azimuth As Double)
    Err.Raise SCODE_NOT_IMPLEMENTED, ERR_SOURCE, _
                "Property SyncToAltitude" & MSG_NOT_IMPLEMENTED
End Sub

