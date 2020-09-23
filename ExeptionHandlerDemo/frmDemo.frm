VERSION 5.00
Begin VB.Form frmDemo 
   BackColor       =   &H80000018&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exception Handler Demo"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExeption3 
      Caption         =   "Show exeption trace"
      Height          =   435
      Left            =   4080
      TabIndex        =   5
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmdCrash 
      Caption         =   "&Crash"
      Height          =   435
      Left            =   4080
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin VB.ComboBox cmbException 
      Height          =   315
      ItemData        =   "frmDemo.frx":0000
      Left            =   120
      List            =   "frmDemo.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1200
      Width           =   3840
   End
   Begin VB.CommandButton cmdExeption2 
      Caption         =   "Show exeption"
      Height          =   435
      Left            =   4080
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton cmdExeption 
      Caption         =   "Show custom warning"
      Height          =   435
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "The exception handler demo will show you how you can implement the exception handler dll into your own project. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Purpose:
' This demo will show you how to use the exeption handler dll
' Created by Edwin Vermeer
' Website http://evict.nl
'
'Description:
' Adding code for handeling Errors/Exceptions/Warnings should be:
' Easy = Just add a dll, add 2 lines to install and uninstall, Add 4 lines of code to each sub
' Generic = Handle all errors in the same way.
' Tracable = You do not need an error stack. Just let the error float up in the procedure tree
'   while adding information in each step.
' Position = Use line numbers. The only disadvantage is that it will add 10% to your .exe file size.
' Robust = Even handle GPF exceptions. (They are getting rare aren't they)
'
'Credits:
' Most of the Exception hanler is from Thushan Fernando.
' see http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=41471&lngWId=1
'
'Modifications:
' The code was wraped into a dll with a global class for easy access.
' An exceptiontype was added to the call so that you can specify if you want the continue button and what
' type of text will be shown.
Option Explicit


Private Sub Form_Load()
On Error GoTo ErrorHandler
      
      ' This exception handler is even capable to capture GPF exceptions.
1     InstallExceptionHandler Me, App, True

     ' Set up the combo to demonstrate exceptions.
2    cmbException.AddItem GetExceptionName(enumExceptionType_AccessViolation)
3    cmbException.ItemData(0) = enumExceptionType_AccessViolation
4    cmbException.AddItem GetExceptionName(enumExceptionType_IllegalInstruction)
5    cmbException.ItemData(1) = enumExceptionType_IllegalInstruction
6    cmbException.AddItem GetExceptionName(enumExceptionType_PriviledgedInstruction)
7    cmbException.ItemData(2) = enumExceptionType_PriviledgedInstruction
8    cmbException.AddItem GetExceptionName(enumExceptionType_ArrayBoundsExceeded)
9    cmbException.ItemData(3) = enumExceptionType_ArrayBoundsExceeded
10   cmbException.AddItem GetExceptionName(enumExceptionType_Breakpoint)
11   cmbException.ItemData(4) = enumExceptionType_Breakpoint
12   cmbException.AddItem GetExceptionName(enumExceptionType_ControlCExit)
13   cmbException.ItemData(5) = enumExceptionType_ControlCExit
14   cmbException.AddItem GetExceptionName(enumExceptionType_DataTypeMisalignment)
15   cmbException.ItemData(6) = enumExceptionType_DataTypeMisalignment
16   cmbException.AddItem GetExceptionName(enumExceptionType_NoncontinuableException)
17   cmbException.ItemData(7) = enumExceptionType_NoncontinuableException
18   cmbException.AddItem GetExceptionName(enumExceptionType_SingleStep)
19   cmbException.ItemData(8) = enumExceptionType_SingleStep
20   cmbException.ListIndex = 0

21   Exit Sub
ErrorHandler:
22   HandleTheException "frmDemo :: Error in cmdExeption2_Click() on line " & Erl() & " triggered by " & Err.Source & "   (Error " & Err.Number & ")" & vbCrLf & Err.Description, "Error in cmdExeption2_Click()", enumExceptionHandling_Exception
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrorHandler
     
     ' You have to make sure that you uninstall this exception handler.
23   UninstallExceptionHandler

24   Exit Sub
ErrorHandler:
25   HandleTheException "frmDemo :: Error in cmdExeption2_Click() on line " & Erl() & " triggered by " & Err.Source & "   (Error " & Err.Number & ")" & vbCrLf & Err.Description, "Error in cmdExeption2_Click()", enumExceptionHandling_Exception
End Sub


'Purpose:
' Sample code that can be use to show a your own warnings
' The last field in the HandleTheException call will specify if that the user can only click continue
Private Sub cmdExeption_Click()
On Error GoTo WarningHandler

26   Err.Raise 999, "My own code", "You clicked on a button and I have a problem with that! Never do that again!"

27   Exit Sub
WarningHandler:
28   HandleTheException "frmDemo :: Custom warning in cmdExeption_Click() on line " & Erl() & " triggered by '" & Err.Source & "'   (Error " & Err.Number & ")" & vbCrLf & Err.Description, "Warning in cmdListen_Click()", enumExceptionHandling_Warning
End Sub


'Purpose:
' Sample code for standard error handling
Private Sub cmdExeption2_Click()
On Error GoTo ErrorHandler

29   MsgBox 1 / 0

30   Exit Sub
ErrorHandler:
31   HandleTheException "frmDemo :: Error in cmdExeption2_Click() on line " & Erl() & " triggered by " & Err.Source & "   (Error " & Err.Number & ")" & vbCrLf & Err.Description, "Error in cmdExeption2_Click()", enumExceptionHandling_Exception
End Sub


'Purpose:
' Code that will show you what happens in case of a GPF exception
Private Sub cmdCrash_Click()
On Error GoTo ErrorHandler
    
32   Call RaiseAnException(cmbException.ItemData(cmbException.ListIndex))

33   Exit Sub
ErrorHandler:
34   HandleTheException "frmDemo :: Error in cmdCrash_Click() on line " & Erl() & " triggered by " & Err.Source & "   (Error " & Err.Number & ")" & vbCrLf & Err.Description, "Error in cmdCrash_Click()", enumExceptionHandling_Exception
End Sub


'Purpose:
' Code that will show you how you can make your error trace
Private Sub cmdExeption3_Click()
On Error GoTo ErrorHandler

35   DoSomeWork

36   Exit Sub
ErrorHandler:
37   HandleTheException "frmDemo :: Error in cmdExeption2_Click() on line " & Erl() & " triggered by " & Err.Source & "   (Error " & Err.Number & ")" & vbCrLf & Err.Description, "Error in cmdExeption2_Click()", enumExceptionHandling_Exception
End Sub


'Purpose:
' This is a module that is used in the error trace.
' All modules that are called from your code should pass on errors with the Err.Raise command.
' Only modules in the top (usually user interface events) will then show the error with the HandleException command.
Private Sub DoSomeWork()
On Error GoTo ErrorHandler
    
38   DoSpecificTask

39   Exit Sub
ErrorHandler:
40   Err.Raise Err.Number, "frmDemo.DoSomeWork", "frmDemo :: Error in DoSomeWork() on line " & Erl() & " triggered by " & Err.Source & "   (Error " & Err.Number & ")" & vbCrLf & Err.Description, "Error in DoSomeWork()"
End Sub


'Purpose:
' This is a module that is used in the error trace.
' All modules that are called from your code should pass on errors with the Err.Raise command.
' Only modules in the top (usually user interface events) will then show the error with the HandleException command.
Private Sub DoSpecificTask()
On Error GoTo ErrorHandler

41   MsgBox 1 / 0 ' Just generate an error for this demo.

42   Exit Sub
ErrorHandler:
43   Err.Raise Err.Number, "frmDemo.DoSpecificTask", "frmDemo :: Error in DoSpecificTask() on line " & Erl() & " triggered by " & Err.Source & "   (Error " & Err.Number & ")" & vbCrLf & Err.Description, "Error in DoSpecificTask()"
End Sub


