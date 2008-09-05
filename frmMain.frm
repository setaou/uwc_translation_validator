VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Language file validator"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9045
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   9045
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox log 
      Height          =   4815
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   2040
      Width           =   8775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   285
      Left            =   8640
      TabIndex        =   6
      Top             =   360
      Width           =   315
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   285
      Left            =   8640
      TabIndex        =   5
      Top             =   960
      Width           =   315
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check"
      Height          =   375
      Left            =   7680
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   8415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   8415
   End
   Begin VB.Label Label2 
      Caption         =   "File to check :"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Reference :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Dim i As Integer
    Dim ref As New cTranslationCompare
    Dim trans As New cTranslationCompare
    
    If Not FileExists(Text1.Text) Then
        MsgBox "Please choose an existing reference file"
        Exit Sub
    End If
    If Not FileExists(Text1.Text) Then
        MsgBox "Please choose an existing file to check"
        Exit Sub
    End If
    
    Debug.Print ref.LoadTranslations(Text1.Text)
    Debug.Print trans.LoadTranslations(Text2.Text)
    
    log.Text = ""
    
    AddToLog "> UWC Language file checker V0.1"
    AddToLog ""
    AddToLog "Reference file : " + Text1.Text
    AddToLog "   - Language : " + ref.GetTransMetadata("language")
    AddToLog "   - Translator : " + ref.GetTransMetadata("translator")
    AddToLog "   - Version : " + ref.GetTransMetadata("version")
    AddToLog "   - Date : " + ref.GetTransMetadata("date")
    AddToLog ""
    AddToLog "File to check : " + Text2.Text
    AddToLog "   - Language : " + trans.GetTransMetadata("language")
    AddToLog "   - Translator : " + trans.GetTransMetadata("translator")
    AddToLog "   - Version : " + trans.GetTransMetadata("version")
    AddToLog "   - Date : " + trans.GetTransMetadata("date")
    AddToLog ""
    AddToLog "Checking ..."
    AddToLog ""
    
    If Len(trans.GetTransMetadata("language")) = 0 Then AddToLog "Error : ""language"" metadata piece is not set"
    If Len(trans.GetTransMetadata("version")) = 0 Then AddToLog "Error : ""version"" metadata piece is not set"
    If Len(trans.GetTransMetadata("translator")) = 0 Then AddToLog "Warning : ""translator"" metadata piece is not set"
    If Len(trans.GetTransMetadata("date")) = 0 Then AddToLog "Warning : ""date"" metadata piece is not set"
    
    If trans.GetTransMetadata("version") <> ref.GetTransMetadata("version") Then AddToLog "Error : The language file does not apply to the same version of UWC as the reference file !"
    
    For i = 1 To NB_TRANSLATIONS
        'Défini en trop
        If (Len(ref.GetTranslation(i)) = 0) And (Len(trans.GetTranslation(i)) > 0) Then
            AddToLog "Warning : Text #" + CStr(i) + " (" + trans.GetTranslation(i) + ") should not be defined !"
        'Doublon
        ElseIf trans.GetTranslationNb(i) > 1 Then
            AddToLog "Error : Text #" + CStr(i) + " is defined " + CStr(trans.GetTranslationNb(i)) + " times! Only the last value will be used (" + trans.GetTranslation(i) + "). The value in the reference file is """ + ref.GetTranslation(i) + """"
        'Non défini
        ElseIf (Len(ref.GetTranslation(i)) > 0) And (Len(trans.GetTranslation(i)) = 0) Then
                AddToLog "Error : Text #" + CStr(i) + " is not defined ! Its value in the reference file is """ + ref.GetTranslation(i) + """"
        'Variables non respectées
        ElseIf (CountInStr("%", ref.GetTranslation(i)) <> CountInStr("%", trans.GetTranslation(i))) Then
                AddToLog "Error : Text #" + CStr(i) + " does not have the right number of variables (%something). There should be " + CStr(CountInStr("%", ref.GetTranslation(i))) + " but there is " + CStr(CountInStr("%", trans.GetTranslation(i)))
        End If
    Next i
    
    AddToLog ""
    AddToLog "Check completed ! If you had no error and no warning then the language file checked is correct !"
End Sub

Private Sub Command2_Click()
    Dim ret As String
    Dim cd As New cCommonDialog

    cd.VBGetOpenFileName ret, , , , , True, "Language files (*.lng)|*.lng", , , "Choose the reference file", , Me.hWnd
    
    If Len(ret) > 0 Then Text1.Text = ret
End Sub


Private Sub Command3_Click()
    Dim ret As String
    Dim cd As New cCommonDialog

    cd.VBGetOpenFileName ret, , , , , True, "Language files (*.lng)|*.lng", , , "Choose the file to check", , Me.hWnd
    
    If Len(ret) > 0 Then Text2.Text = ret
End Sub


Private Sub Form_Load()
    Text1.Text = App.path + "\english.lng"
End Sub


'Ajout d'un texte au log
Private Sub AddToLog(Texte As String)
    log.Text = Right(log.Text, 30000) + Texte + vbCrLf
    log.SelStart = Len(log.Text)
End Sub

