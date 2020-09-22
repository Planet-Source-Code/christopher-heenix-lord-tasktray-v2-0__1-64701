VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' Christopher Lord
' heenix@project-angel.net
' 17th March 2006
'
' You are free to reuse this code, in segments or whole, in
' any projects that you may want. As long as the name and
' date of the original creator is shown in one way or another.
'
' * Fixed refresh bug (thanks to Tom Pydeski)
' * Fixed bug where changing icon stops balloons!
' * Added error reporting code
' * Improved demo interface

' The data type for the icon in
' side the task bar, very simple
Private Type NOTIFYICONDATAW
    icoSize As Long
    icoHwnd As Long
    icoId As Long
    icoFlags As Long
    icoCallbackMessage As Long
    icoSource As Long
    icoTooltip(0 To 255) As Byte
    icoState As Long
    icoStateMask As Long
    szInfo(0 To 511) As Byte
    uTimeOutOrVersion As Long
    szInfoTitle(0 To 127) As Byte
    dwInfoFlags As Long
End Type

' The structure that contains all
' the possible types of balloons
Public Enum EBalloonIconTypes
    NIIF_NONE = 0
    NIIF_INFO = 1
    NIIF_WARNING = 2
    NIIF_ERROR = 3
    NIIF_NOSOUND = &H10
End Enum

' Events that this class can raise here
Event DoubleClick(Button As Integer)
Event MouseDown(Button As Integer)
Event MouseUp(Button As Integer)
Event BalloonTimeout()
Event BalloonClicked()
Event BalloonShow()
Event BalloonHide()

' These are modular level variables that allow
' us to determine certain aspects of the icon
' and share control of the forms events
Private WithEvents m_Form As Form
Attribute m_Form.VB_VarHelpID = -1
Private m_AutoRefresh As Boolean
Private m_IconLoaded As Boolean
Private m_IconData As NOTIFYICONDATAW

' Constants related to events that we can
' extract from the system tray are here
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206

' The events we can extract from the balloons
Private Const WM_USER = &H400
Private Const NIN_BALLOONSHOW = (WM_USER + 2)
Private Const NIN_BALLOONHIDE = (WM_USER + 3)
Private Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Private Const NIN_BALLOONUSERCLICK = (WM_USER + 5)

' Constants releated to adding and removing
' the icon from the task tray and response level
Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Private Const NIF_INFO = &H10

' These inform windows what action
' we are about to perform with the icon
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETVERSION = &H4

' The windows API call that allows us to
' manage the icon in the task tray
Private Declare Function Shell_NotifyIconW Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATAW) As Long

Public Function Initialize(Form As Form) As Boolean
    '
    ' Initialize the icon handler and any variables
    ' that may be required by the api call
    '
    On Error GoTo ErrorHandler
    Set m_Form = Form
    
    With m_IconData
        ' Setup the flags and other settings of the icon
        ' like we normally would using the forms settings
        .icoHwnd = m_Form.hWnd
        .icoId = m_Form.Icon
        .icoCallbackMessage = WM_MOUSEMOVE
        .icoSource = m_Form.Icon.Handle
        
        ' Setup new variables to suit the balloon message
        .uTimeOutOrVersion = 3
        .icoSize = 936
    End With
    
    ' Ensure we are on a pixel
    ' scalemode otherwise messages
    ' will not get interpreted properly
    m_Form.ScaleMode = 3
    
    ' Completed sucessfully
    Initialize = True
    Exit Function

ErrorHandler:
    Initialize = False
End Function

Public Function ShowBalloonTip(Message As String, Optional Title As String, Optional Style As EBalloonIconTypes, Optional Timeout = 10000) As Boolean
    '
    ' You must know the following in order to
    ' use this feature properly:
    '
    '      If the timeout is bigger than the systems
    '      maximum then it will be brought down.
    '      (Typically, the maximum is 30 seconds)
    '
    '      If the timeout is less than the systems
    '      minimum then it will be raised upwards.
    '      (Typically, the minimum is 10 seconds)
    '
    On Error GoTo ErrorHandler
    
    If m_IconLoaded Then
        ' Convert the title and message into an array
        stringToArray Message, m_IconData.szInfo, 512
        stringToArray Title, m_IconData.szInfoTitle, 128
       
        ' Store the timeout value here and the icon
        m_IconData.uTimeOutOrVersion = Timeout
        m_IconData.dwInfoFlags = Style
        m_IconData.icoFlags = NIF_INFO
        
        ' Update the icon with the new information
        Shell_NotifyIconW NIM_MODIFY, m_IconData
        
        ' Completed it correctly
        ShowBalloonTip = True
        Exit Function
    End If

ErrorHandler:
    ' Failed showing the balloon
    ShowBalloonTip = False
End Function

Public Function AddIcon() As Boolean
    '
    ' Adds the icon to the task tray so
    ' the form can handle the events
    '
    On Error GoTo ErrorHandler
    
    If m_IconLoaded = False Then
        ' Change the flags to the ones we should use when adding an icon
        m_IconData.icoFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        
        ' Dont bother adding an icon if its already been done
        Call Shell_NotifyIconW(NIM_ADD, m_IconData)
        Call Shell_NotifyIconW(NIM_SETVERSION, m_IconData)
        
        ' Only allowed one icon per form
        m_IconLoaded = True
        
        ' Completed with no errors
        AddIcon = True
        Exit Function
    End If
    
ErrorHandler:
    ' An error somewhere
    AddIcon = False
End Function

Public Function Refresh() As Boolean
    '
    ' Refresh the icon in the task
    ' tray if it exists at all
    '
    On Error GoTo ErrorHandler
    
    If m_IconLoaded Then
        ' Thanks to Tom Pydeski for fixing this bug!
        m_IconData.icoFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        
        ' Only bother to refresh if it actually exists
        Call Shell_NotifyIconW(NIM_MODIFY, m_IconData)
        
        ' No errors and were done
        Refresh = True
        Exit Function
    End If
    
ErrorHandler:
    ' Error, or its not loaded
    Refresh = False
End Function

Property Get Visible() As Boolean
    ' Returns if the icon is loaded
    Visible = m_IconLoaded
End Property


Public Function RemoveIcon()
    '
    ' Removes the icon from the task tray
    ' but only if we already have one to delete
    '
    On Error GoTo ErrorHandler
    
    If m_IconLoaded Then
        ' Remove the icon from the task tray
        Call Shell_NotifyIconW(NIM_DELETE, m_IconData)
        m_IconLoaded = False
        
        ' No errors and removed
        RemoveIcon = True
        Exit Function
    End If
    
ErrorHandler:
    ' Error or its not loaded
    RemoveIcon = False
End Function

Private Sub Class_Terminate()
    ' If the icon is still in the tray, remove it
    If m_IconLoaded = True Then RemoveIcon
End Sub

Property Get AutoRefresh() As Boolean
    ' Return the value of m_AutoRefresh
    AutoRefresh = m_AutoRefresh
End Property

Property Let AutoRefresh(Value As Boolean)
    ' Save the auto update value
    m_AutoRefresh = Value
End Property

Property Get Tooltip() As String
    ' Simply return the tool tip
    Tooltip = m_IconData.icoTooltip
End Property

Property Let Tooltip(Message As String)
    ' Ensure the delimiter of null is kept here
    stringToArray Message, m_IconData.icoTooltip, 256
    
    ' m_IconData.icoTooltip = Message & Chr(0)
    If m_AutoRefresh Then Call Refresh
End Property

Property Get Icon() As Long
    ' Simply return the icon
    Icon = m_IconData.icoSource
End Property

Property Let Icon(Icon As Long)
    ' Store the data into our data
    m_IconData.icoSource = Icon
    If m_AutoRefresh Then Call Refresh
End Property

Private Sub m_Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '
    ' We will handle the MouseMove event
    ' for the form so we can trap events
    '
    Select Case X
        ' Events related to double clicking the icon
        Case WM_LBUTTONDBLCLK: RaiseEvent DoubleClick(0)
        Case WM_RBUTTONDBLCLK: RaiseEvent DoubleClick(1)
        
        ' Events related to button down on the icon
        Case WM_LBUTTONDOWN: RaiseEvent MouseDown(0)
        Case WM_RBUTTONDOWN: RaiseEvent MouseDown(1)
        
        ' Events related to button up on the icon
        Case WM_LBUTTONUP: RaiseEvent MouseUp(0)
        Case WM_LBUTTONUP: RaiseEvent MouseUp(1)

        ' Events related to the balloon are here xD
        Case NIN_BALLOONSHOW: RaiseEvent BalloonShow
        Case NIN_BALLOONHIDE: RaiseEvent BalloonHide
        Case NIN_BALLOONTIMEOUT: RaiseEvent BalloonTimeout
        Case NIN_BALLOONUSERCLICK: RaiseEvent BalloonClicked
    End Select
End Sub

Private Sub stringToArray(ByVal sString As String, bArray() As Byte, ByVal lMaxSize As Long)
    '
    ' Converts a string into a byte array then
    ' transfers it to the main array and obeying
    ' any limits that have been set
    '
    Dim Bytes() As Byte
    Dim Pointer As Long
    Dim PointerEmpty As Long
    
    If Len(sString) > 0 Then
        ' Get the string into an array
        ' of bytes so we can use it
        Bytes = sString

        For Pointer = 0 To UBound(Bytes)
            ' Store it into the next array and exit
            ' when we have reached the limit
            bArray(Pointer) = Bytes(Pointer)
            If (Pointer = (lMaxSize - 2)) Then Exit For
        Next Pointer

        For PointerEmpty = Pointer To lMaxSize - 1
            ' Fill the rest of the array with an
            ' empty character (in this case 0)
            bArray(PointerEmpty) = 0
        Next PointerEmpty
    End If
End Sub
