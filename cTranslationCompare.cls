VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTranslationCompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type tMetaData
    ID As String
    Value As String
End Type

Private Translations() As String
Private TranslationsNb() As Integer
Private TransMetadata() As tMetaData

Public TransLoaded As Boolean
Public Sub ClearTranslation()
    'Efface les anciennes traductions et métadonnées
    ReDim Translations(1 To NB_TRANSLATIONS)
    ReDim TranslationsNb(1 To NB_TRANSLATIONS)
    ReDim TransMetadata(0 To 0)
    TransLoaded = False
End Sub

Public Function LoadTranslations(Optional ByVal File As String = "français.lng") As Boolean
    Dim f As Long, Ligne As String
    Dim ID As String, Value As String

    File = File
    
    'Efface les anciennes traductions et métadonnées
    ClearTranslation
    
    'Si le fichier existe
    If FileExists(File) Then
    'Efface les anciennes traductions et métadonnées
    ClearTranslation
    
        'Ouverture
        f = FreeFile
        Open File For Input As #f
        
        'Pour chaque ligne du fichier
        Do
            Line Input #f, Ligne
            
            'Si la ligne n'est pas vide, n'est pas un commentaire et est valide
            If ((Len(Ligne) > 0) And (Left(Ligne, 2) <> "//") And (InStr(1, Ligne, "=") > 0)) Then
                
                'Sépare l'id et la valeur
                ID = Left(Ligne, InStr(1, Ligne, "=") - 2)
                Value = Right(Ligne, Len(Ligne) - InStr(1, Ligne, "=") - 1)
                
                'Si l'entrée est numérique
                If IsNumeric(ID) Then
                    'Ajoute la traduction
                    Translations(Val(ID)) = Value
                    'Compte les répétisions
                    TranslationsNb(Val(ID)) = TranslationsNb(Val(ID)) + 1
                'Si l'entrée est textuelle
                Else
                    'Ajoute aux métadonnées
                    ReDim Preserve TransMetadata(0 To UBound(TransMetadata) + 1)
                    With TransMetadata(UBound(TransMetadata))
                        .ID = ID
                        .Value = Value
                    End With
                End If

            End If
        Loop Until EOF(f)
        
        Close #f
        
        'Chargement OK
        ReDim Preserve TransMetadata(0 To UBound(TransMetadata) + 1)
        With TransMetadata(UBound(TransMetadata))
            .ID = "file"
            .Value = GetFileName(File)
        End With
        TransLoaded = True
        
        LoadTranslations = True
        
    Else
    
        LoadTranslations = False

    End If
    
End Function

Public Function GetFileName(ByVal path As String) As String
    GetFileName = Mid(path, InStrRev(path, "\") + 1, Len(path) - InStrRev(path, "\"))
End Function

Public Function GetTransMetadata(ID As String) As String
    Dim i As Integer
    
    For i = 1 To UBound(TransMetadata)
        If TransMetadata(i).ID = ID Then GetTransMetadata = TransMetadata(i).Value
    Next i
End Function




Public Function GetTranslation(Number As Integer) As String
    GetTranslation = Translations(Number)
End Function
Public Function GetTranslationNb(Number As Integer) As Integer
    GetTranslationNb = TranslationsNb(Number)
End Function

