VERSION 5.00
Begin VB.Form frmCalc 
   Caption         =   "Calculator-rupish"
   ClientHeight    =   2970
   ClientLeft      =   3870
   ClientTop       =   3615
   ClientWidth     =   3735
   Icon            =   "frmCalc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   3735
   Begin VB.CommandButton cmdArrayNum 
      Caption         =   "9"
      Height          =   375
      Index           =   9
      Left            =   1920
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton cmdArrayNum 
      Caption         =   "8"
      Height          =   375
      Index           =   8
      Left            =   1320
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton cmdArrayNum 
      Caption         =   "7"
      Height          =   375
      Index           =   7
      Left            =   720
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton cmdArrayNum 
      Caption         =   "6"
      Height          =   375
      Index           =   6
      Left            =   1920
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton cmdArrayNum 
      Caption         =   "5"
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton cmdArrayNum 
      Caption         =   "4"
      Height          =   375
      Index           =   4
      Left            =   720
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton cmdArrayNum 
      Caption         =   "3"
      Height          =   375
      Index           =   3
      Left            =   1920
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton cmdArrayNum 
      Caption         =   "2"
      Height          =   375
      Index           =   2
      Left            =   1320
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1920
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdArrayNum 
      Caption         =   "1"
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton cmdArrayNum 
      Caption         =   "0"
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton cmdSign 
      Caption         =   "+/-"
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton cmdDec 
      Caption         =   "."
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton cmdEql 
      Caption         =   "="
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton cmdPlus 
      Caption         =   "+"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton cmdInv 
      Caption         =   "1/x"
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton cmdMinus 
      Caption         =   "-"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton cmdPer 
      Caption         =   "%"
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton cmdMul 
      Caption         =   "*"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton cmdDiv 
      Caption         =   "/"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton cmdSqrt 
      Caption         =   "sqrt"
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton cmdMPlus 
      Caption         =   "M+"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton cmdMS 
      Caption         =   "MS"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton cmdMR 
      Caption         =   "MR"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton cmdMC 
      Caption         =   "MC"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "C"
      Height          =   375
      Left            =   2700
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   910
   End
   Begin VB.CommandButton cmdCE 
      Caption         =   "CE"
      Height          =   375
      Left            =   1710
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   900
   End
   Begin VB.CommandButton cmdBkspc 
      Caption         =   "BkSpace"
      Height          =   375
      Left            =   720
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   900
   End
   Begin VB.TextBox txtDummy 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Left            =   200
      TabIndex        =   1
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblOut 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      Caption         =   "0."
      Height          =   195
      Left            =   240
      TabIndex        =   29
      Top             =   120
      Width           =   3250
   End
   Begin VB.Label lblDummy 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   0
      Width           =   3495
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu submnuEditCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu submnuEditPaste 
         Caption         =   "Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu submnuViewStandard 
         Caption         =   "Standard"
         Checked         =   -1  'True
      End
      Begin VB.Menu submnuViewScientific 
         Caption         =   "Scientific"
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu submnuViewGroup 
         Caption         =   "Digit Grouping"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help "
      Begin VB.Menu submnuHelpTopics 
         Caption         =   "Help Topics"
      End
      Begin VB.Menu dash2 
         Caption         =   "-"
      End
      Begin VB.Menu submnuHelpCalc 
         Caption         =   "About Calculator"
      End
   End
End
Attribute VB_Name = "frmCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 ' MY CALCULATOR'S CODE GOES HERE
 
  Option Explicit

  Dim boolValid1 As Boolean
 ' Flag boolValid1 is to handle 0. on the screen
  
  Dim boolValid2 As Boolean
 ' Flag boolValid2 is used to handle decimal _
   point on the screen
 
  Dim bytTempNum As Byte
 ' Byte variable bytTempNum is used to store _
   0 or 1 only
  
  Dim bytNumLength As Byte
 ' Byte variable bytNumLength is to count no of _
   Characters on the screen
  
  Dim strCurrentOper As String
 ' String variable strCurrentOper is used to _
   store current operator
 
  Dim dblNum1 As Double
 ' Double variable dblNum1 is used to store _
   the previous number
   
 
Private Sub cmdArrayNum_Click(Index As Integer)
  
 ' This part clears the screen when C or = is pressed
   If Not boolValid1 And Not boolValid2 Then
      lblOut.Caption = ""
      boolValid1 = True
   End If
 
 ' This part counts the length of chars in the screen
    bytNumLength = Len(lblOut.Caption)
 
 ' This part sets the limit for max no of chars _
   to be entered
   If bytNumLength > 32 Then
      Beep
      Exit Sub
   End If
  
 ' This part prints the pressed nos on the screen _
   by checking the Index of the Command Button as _
   Control Array
 ' For Non Zero numbers
   If Index <> 0 Then
      If bytNumLength <> 0 Then
         lblOut.Caption = Left(lblOut.Caption, _
         bytNumLength - bytTempNum) + cmdArrayNum _
         (Index).Caption
      Else
         lblOut.Caption = lblOut.Caption + _
         cmdArrayNum(Index).Caption
      End If
   Else
 
 ' For zero
      If bytNumLength <> 0 Then
         lblOut.Caption = Left(lblOut.Caption, _
         bytNumLength - bytTempNum) + "0"
      Else
         lblOut.Caption = "0"
         boolValid1 = False
      End If
   End If
  
 ' This part prints a Decimal at the end of the number
   If Not boolValid2 Then
      lblOut.Caption = lblOut.Caption + "."
   End If

End Sub

Private Sub cmdBkspc_Click()

 ' This subroutine has the code for back scpace
 
 Dim bytstrlength As Byte
 ' Byte variable bytStrLength is used for length of the _
  string
  
 Dim bytTempNum1 As Byte
 ' Byte variable bytTempNum1 is used to store 2 or 3 only

 ' Finds the length of the string
   bytstrlength = Len(lblOut.Caption)
  
 ' Checks for the decimal at the right most position _
   if found then deletes the last two numbers and then printing _
   the decimal, hence deleting the second last number
  
   If Right(lblOut.Caption, 1) = "." Then
      lblOut.Caption = Left(lblOut.Caption, bytstrlength - 2)
      lblOut.Caption = lblOut.Caption + "."
      boolValid2 = False
      bytTempNum = 1
   Else
 ' If not found then deletes the last number

      lblOut.Caption = Left(lblOut.Caption, bytstrlength - 1)
   End If
  
 ' Checks for the "-" at the left most position and _
   sets bytTempNum1 accordingly
   If Left(lblOut.Caption, 1) = "-" Then
      bytTempNum1 = 3
   Else
      bytTempNum1 = 2
   End If
  
 ' This part prints "0." at the end
   If bytstrlength = bytTempNum1 Then
      lblOut.Caption = "0."
      bytTempNum = 1
      boolValid1 = False
      boolValid2 = False
   End If

End Sub

Private Sub cmdC_Click()

 ' This subroutine runs when C is pressed

   lblOut.Caption = "0."
   bytTempNum = 1
   boolValid1 = False
   boolValid2 = False

End Sub

Private Sub cmdCE_Click()

 ' This subroutine runs when CE is pressed

   lblOut.Caption = "0."
   bytTempNum = 1
   boolValid1 = False
   boolValid2 = False

End Sub

Private Sub cmdDec_Click()
  
 ' This subroutine sets the Flag boolValid2 to True _
   when a decimal Character is found at right of the number
   If Not boolValid1 Then
      lblOut.Caption = "0."
   End If
      
   If Right(lblOut, 1) = "." Then
      boolValid2 = True
      bytTempNum = 0
   End If

End Sub

Private Sub cmdDiv_Click()

 ' Division code goes here
   strCurrentOper = "/"
   If Right(lblOut.Caption, 1) = "." Then
      dblNum1 = Val(Left(lblOut.Caption, _
                Len(lblOut.Caption) - 1))
   Else
      dblNum1 = Val(lblOut.Caption)
   End If
   boolValid1 = False
   boolValid2 = False
   bytTempNum = 1

End Sub

Private Sub cmdEql_Click()

 ' This part is for + operator
   If strCurrentOper = "+" Then
      If Right(lblOut.Caption, 1) = "." Then
         lblOut.Caption = dblNum1 + Val(Left _
         (lblOut.Caption, Len(lblOut.Caption) - 1))
      Else
         lblOut.Caption = dblNum1 + Val(lblOut.Caption)
      End If
   End If
   
 ' This part is for - operator
   If strCurrentOper = "-" Then
      If Right(lblOut.Caption, 1) = "." Then
         lblOut.Caption = dblNum1 - Val(Left _
         (lblOut.Caption, Len(lblOut.Caption) - 1))
      Else
         lblOut.Caption = dblNum1 - Val(lblOut.Caption)
      End If
   End If
 
 ' This part is * operator
   If strCurrentOper = "*" Then
      If Right(lblOut.Caption, 1) = "." Then
         lblOut.Caption = dblNum1 * Val(Left _
         (lblOut.Caption, Len(lblOut.Caption) - 1))
      Else
         lblOut.Caption = dblNum1 * Val(lblOut.Caption)
      End If
   End If
   
 ' This part is for / operator
   If strCurrentOper = "/" Then
      
 ' This part checks for divide by zero error
      If Val(lblOut.Caption) = 0 Then
         lblOut.Caption = "Divide by Zero Error"
         Exit Sub
         boolValid1 = False
         boolValid2 = False
         bytTempNum = 1
      End If
 ' This part divide the valid numbers
      If Right(lblOut.Caption, 1) = "." Then
         lblOut.Caption = dblNum1 / Val(Left _
         (lblOut.Caption, Len(lblOut.Caption) - 1))
      Else
         lblOut.Caption = dblNum1 / Val(lblOut.Caption)
      End If
   End If
  
 ' This part is for inserting the decimal point
   If InStr(lblOut.Caption, ".") Then
      lblOut.Caption = lblOut.Caption
   Else
     lblOut.Caption = lblOut.Caption + "."
   End If
 ' This part is for clearing the screen
   boolValid1 = False
   boolValid2 = False
   bytTempNum = 1

End Sub

Private Sub cmdInv_Click()

  ' Inverse code goes here
  If Val(lblOut.Caption) <> 0 Then
     lblOut.Caption = 1 / Val(lblOut.Caption)
  Else
     lblOut.Caption = "Divide by Zero Error"
  End If
 
 ' This part adds "." to the text if it is "1"
  If Val(lblOut.Caption) = 1 Then
     lblOut.Caption = lblOut.Caption + "."
  End If
  
  boolValid1 = False
  boolValid2 = False
  bytTempNum = 1

End Sub

Private Sub cmdMC_Click()

 ' Memory clear code goes here

End Sub

Private Sub cmdMinus_Click()
   
 ' Minus code goes here
   strCurrentOper = "-"
   If Right(lblOut.Caption, 1) = "." Then
      dblNum1 = Val(Left(lblOut.Caption, _
                Len(lblOut.Caption) - 1))
   Else
      dblNum1 = Val(lblOut.Caption)
   End If
   boolValid1 = False
   boolValid2 = False
   bytTempNum = 1


End Sub

Private Sub cmdMPlus_Click()

 ' Memory plus code goes here

End Sub

Private Sub cmdMR_Click()

 ' Memory recall code goes here

End Sub

Private Sub cmdMS_Click()

 ' MS code goes here

End Sub

Private Sub cmdMul_Click()

 ' Multiply code goes here
   strCurrentOper = "*"
   If Right(lblOut.Caption, 1) = "." Then
      dblNum1 = Val(Left(lblOut.Caption, _
                Len(lblOut.Caption) - 1))
   Else
      dblNum1 = Val(lblOut.Caption)
   End If
   boolValid1 = False
   boolValid2 = False
   bytTempNum = 1

End Sub

Private Sub cmdPer_Click()

 ' Percentage code goes here

End Sub

Private Sub cmdPlus_Click()
   
 ' Plus code goes here
   strCurrentOper = "+"
   If Right(lblOut.Caption, 1) = "." Then
      dblNum1 = Val(Left(lblOut.Caption, _
                Len(lblOut.Caption) - 1))
   Else
      dblNum1 = Val(lblOut.Caption)
   End If
   boolValid1 = False
   boolValid2 = False
   bytTempNum = 1
   
End Sub

Private Sub cmdSign_Click()
  
 ' This subroutine adds a minus sign at the left most _
   position and deletes it if minus sign is already present
  
   If lblOut.Caption = "0." Then
      Exit Sub
   End If
  
   If Left(lblOut.Caption, 1) <> "-" Then
      lblOut.Caption = "-" + lblOut.Caption
   Else
      lblOut.Caption = Right(lblOut.Caption, _
      Len(lblOut.Caption) - 1)
   End If
  
End Sub

Private Sub cmdSqrt_Click()

 ' Square root code goes here
 

End Sub

Private Sub Form_Load()

 ' This subroutine sets the variable bytTempNum to 1
   bytTempNum = 1
   submnuEditPaste.Enabled = True

End Sub

Private Sub submnuEditCopy_Click()

 ' This subroutine copies the data in the screen _
   stores the text in the clipboard
   If Right(lblOut.Caption, 1) <> "." Then
      Clipboard.SetText (lblOut.Caption)
   Else
      Clipboard.SetText (Left(lblOut.Caption, Len(lblOut.Caption) - 1))
   End If
   
End Sub

Private Sub submnuEditPaste_Click()
 
 ' Paste code goes here
   If IsNumeric(Clipboard.GetText) Then
      lblOut.Caption = Trim(Clipboard.GetText)
   End If
   
   If Not InStr(lblOut.Caption, ".") Then
 '    Dummy statement
   Else
      lblOut.Caption = lblOut.Caption + "."
   End If
   
End Sub

Private Sub submnuViewScientific_Click()
   
   submnuViewScientific.Checked = True
   submnuViewStandard.Checked = False
 ' Scientific calculator's code goes here
 ' Load Form frmScientific
 ' Hide frmCalc

End Sub

Private Sub submnuViewStandard_Click()
 ' Standard calculator
  
   submnuViewScientific.Checked = False
   submnuViewStandard.Checked = True
 ' Loads Form frmCalc
 ' Hide Form frmScientific

End Sub
