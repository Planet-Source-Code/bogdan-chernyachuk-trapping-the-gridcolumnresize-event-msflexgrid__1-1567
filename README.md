<div align="center">

## Trapping the GridColumnResize event \(MSFlexGrid\)


</div>

### Description

The MSFlexGrid doesn't provide a ColumnResize event. However it's needed in some apps, where this grid is used. The code below shows how emulate the ColumnResize event by analizing the sequences of other windows messages.
 
### More Info
 
The knowlege of API is required.

You must have MSflxgrd.ocx installed.

Create a Form frmMain and place a MSFlexGrid control with name MSFlexGrid1 and a Module. The code is given below.

Don't ever finish the application with the Stop button of VB Environment. This will cause VBE to terminate.

Application was tested under VB 5.0 (SP3)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Bogdan Chernyachuk](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bogdan-chernyachuk.md)
**Level**          |Unknown
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bogdan-chernyachuk-trapping-the-gridcolumnresize-event-msflexgrid__1-1567/archive/master.zip)

### API Declarations

```
' Function to retrieve the address of the current Message-Handling routine
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
' Function to define the address of the Message-Handling routine
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
' Function to execute a function residing at a specific memory address
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Windows messages constants
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_ERASEBKGND = &H14
```


### Source Code

```
/***************************   frmMain   ****************************/
VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmMain
  BorderStyle   =  3 'Fixed Dialog
  Caption     =  "Resize the Grid !!!"
  ClientHeight  =  4110
  ClientLeft   =  4650
  ClientTop    =  3750
  ClientWidth   =  6735
  LinkTopic    =  "Form1"
  MaxButton    =  0  'False
  MinButton    =  0  'False
  ScaleHeight   =  4110
  ScaleWidth   =  6735
  ShowInTaskbar  =  0  'False
  Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1
   Height     =  3015
   Left      =  120
   TabIndex    =  0
   Top       =  960
   Width      =  6495
   _ExtentX    =  11456
   _ExtentY    =  5318
   _Version    =  65541
   Rows      =  4
   Cols      =  4
   AllowUserResizing=  1
  End
  Begin VB.Label Label2
   Caption     =  "Try to resize the columns of MSFlexGrid. All the columns will be resized proportionally."
   BeginProperty Font
     Name      =  "MS Sans Serif"
     Size      =  9.75
     Charset     =  204
     Weight     =  400
     Underline    =  0  'False
     Italic     =  0  'False
     Strikethrough  =  0  'False
   EndProperty
   ForeColor    =  &H8000000D&
   Height     =  615
   Left      =  1320
   TabIndex    =  1
   Top       =  120
   Width      =  3975
  End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' This constant is used to refer to the Message Handling function in a given window
Private Const GWL_WNDPROC = (-4)
Private Sub Form_Load()
  'Save the address of the existing Message Handler
  g_lngDefaultHandler = GetWindowLong(Me.MSFlexGrid1.hwnd, GWL_WNDPROC)
  'Define new message handler routine
  Call SetWindowLong(Me.MSFlexGrid1.hwnd, GWL_WNDPROC, AddressOf GridMessage)
End Sub
Private Sub Form_Unload(Cancel As Integer)
  'Return the old handler back
  Call SetWindowLong(Me.MSFlexGrid1.hwnd, GWL_WNDPROC, g_lngDefaultHandler)
End Sub
Public Sub ResizeGridProportional()
Dim SumWidth  As Long
Dim i As Integer
With MSFlexGrid1
  For i = 1 To .Cols
    SumWidth = SumWidth + .ColWidth(i - 1)
  Next i
  For i = 1 To .Cols
    .ColWidth(i - 1) = SumWidth / .Cols
  Next i
End With
End Sub
/* ******************** MODULE ***********************************/
Attribute VB_Name = "mHandlers"
'
Option Explicit
Public g_lngDefaultHandler As Long ' Original handler of the grid events
Private m_bLMousePressed As Boolean 'true if the left button is pressed
Private m_bLMouseClicked As Boolean 'true just after the click (i.e. just after the left button is released)
'API declarations ============================================================
' Function to retrieve the address of the current Message-Handling routine
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
' Function to define the address of the Message-Handling routine
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
' Function to execute a function residing at a specific memory address
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Windows messages constants
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_ERASEBKGND = &H14
'==============================================================================
'this is our event handler
Public Function GridMessage(ByVal hwnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
  If m_bLMousePressed And Msg = WM_LBUTTONUP Then
  'button have been just released
    m_bLMousePressed = False
    m_bLMouseClicked = True
  End If
  If Not (m_bLMousePressed) And Msg = WM_LBUTTONDOWN Then
  'button have been just pressed
    m_bLMousePressed = True
    m_bLMouseClicked = False
  End If
  If m_bLMouseClicked And (Msg = WM_ERASEBKGND) Then
  'Only when resize happens this event may occur after releasing the button !
  'When user is making a simple click on grid,
  'the WM_ERASEBKGND event occurs before WM_LBUTTONUP,
  'and therefore will not be handled there
    frmMain.ResizeGridProportional
    m_bLMouseClicked = False
  End If
  'call the default message handler
  GridMessage = CallWindowProc(g_lngDefaultHandler, hwnd, Msg, wp, lp)
End Function
```

