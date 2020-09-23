VERSION 5.00
Begin VB.UserControl UserControl1 
   ClientHeight    =   720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1440
   ClipBehavior    =   0  'None
   ScaleHeight     =   720
   ScaleWidth      =   1440
   Begin VB.CommandButton ThButton 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_SYSCOMMAND = &H112
'Déclarations d'événements:
Event Click() 'MappingInfo=ThButton,ThButton,-1,Click
Attribute Click.VB_Description = "Se produit lorsque l'utilisateur appuie sur un bouton de la souris puis le relâche au-dessus d'un objet."



Private Sub ThButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
UserControl_MouseDown Button, Shift, X, Y
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Shift = 3 Then
        ReleaseCapture
        SendMessage UserControl.hwnd, WM_SYSCOMMAND, 61458, 0
        UserControl_Resize
    Else
        SendMessage UserControl.hwnd, WM_SYSCOMMAND, 0, 0
    End If

End Sub

Private Sub UserControl_Resize()
ThButton.Move 0, 0, ScaleWidth, ScaleHeight
End Sub
'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=ThButton,ThButton,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Renvoie ou définit le texte affiché dans la barre de titre d'un objet ou sous l'icône d'un objet."
    Caption = ThButton.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    ThButton.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

Private Sub ThButton_Click()
    RaiseEvent Click
End Sub

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=ThButton,ThButton,-1,Style
Public Property Get Style() As Integer
Attribute Style.VB_Description = "Renvoie ou définit l'aspect du contrôle, qu'il soit standard (style Windows), ou graphique (avec image personnalisée)."
    Style = ThButton.Style
End Property

'Charger les valeurs des propriétés à partir du stockage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    ThButton.Caption = PropBag.ReadProperty("Caption", "Command1")
End Sub

'Écrire les valeurs des propriétés dans le stockage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Caption", ThButton.Caption, "Command1")
End Sub

