VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   9225
   ClientLeft      =   2475
   ClientTop       =   1635
   ClientWidth     =   11025
   Icon            =   "FormG1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   650
   ScaleMode       =   0  'User
   ScaleWidth      =   835.922
   Begin VB.PictureBox Picture1 
      Height          =   8775
      Left            =   120
      ScaleHeight     =   600
      ScaleMode       =   0  'User
      ScaleWidth      =   800
      TabIndex        =   0
      Top             =   0
      Width           =   10785
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   8880
      Width           =   8895
   End
   Begin VB.Menu mnuZoomMenu 
      Caption         =   "Zoom"
      Visible         =   0   'False
      Begin VB.Menu mnuCancel 
         Caption         =   "&Cancel"
      End
      Begin VB.Menu mnuInzoomen 
         Caption         =   "Inzoomen"
         Begin VB.Menu mnuZinx1_2 
            Caption         =   "Zoom In x 1.2"
         End
         Begin VB.Menu mnuInx2 
            Caption         =   "Zoom In x&2"
         End
         Begin VB.Menu mnuInx4 
            Caption         =   "Zoom In x&4"
         End
         Begin VB.Menu mnuInx10 
            Caption         =   "Zoom In x1&0"
         End
      End
      Begin VB.Menu mnuGelijk 
         Caption         =   "Niet Zoomen"
      End
      Begin VB.Menu mnuUitzoomen 
         Caption         =   "Uitzoomen"
         Begin VB.Menu mnuZuitx1_2 
            Caption         =   "Zoom uit x 1.2"
         End
         Begin VB.Menu mnuUitx2 
            Caption         =   "Zoom Uit x2"
         End
         Begin VB.Menu mnuUitx4 
            Caption         =   "Zoom Uit x4"
         End
         Begin VB.Menu mnuUitx10 
            Caption         =   "Zoom Uit x10"
         End
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuKleuren 
         Caption         =   "Kleuren"
         Begin VB.Menu mnuK32 
            Caption         =   "Aantal &kleuren = 32"
         End
         Begin VB.Menu mnuK128 
            Caption         =   "Aantal k&leuren = 128"
         End
         Begin VB.Menu mnuK512 
            Caption         =   "Aantal kl&euren = 512"
         End
         Begin VB.Menu mnuK1024 
            Caption         =   "Aantal kle&uren=1024"
         End
      End
      Begin VB.Menu mnuRestart 
         Caption         =   "Restart"
      End
      Begin VB.Menu mnuChangecolors 
         Caption         =   "Ander kleurenpalet"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Dim X As Double, Y As Double
Dim x1 As Double, y1 As Double 'Hulpvariabelen voor de berekeningen
Dim cx As Double, cy As Double
Dim BeeldHoogte As Integer
Dim BeeldBreedte As Integer
Dim BeeldMiddenX As Double, BeeldMiddenY As Double  'x- en y-coörd van het beeldmidden
Dim k As Double 'Vergrotingsfactor
Dim Begintijd As Date ' Om de rekentijd te berekenen
Dim Eindtijd As Date
Dim UpRood As Boolean
Dim UpGroen As Boolean
Dim UpBlauw As Boolean

Dim Rood As Integer
Dim Groen As Integer
Dim Blauw As Integer

Dim ROODINCREMENT As Integer
Dim GROENINCREMENT As Integer
Dim BLAUWINCREMENT As Integer

'Factor waarmee het beeld wordt vergroot bij een linkermuisklik
Dim ZoomFactor As Double

' Palet van 512 kleuren
Dim Palet(1024) As ColorConstants

'Max rekendiepte (en dus ook nodige kleuren)
Dim MaxKleuren As Integer

'Telt aantal iteraties bij het berekenen van de kleur
Dim Teller As Integer
'Inladen (openen) van het formulier
Private Sub Form_Load()
Dim i As Integer

'Zet initiële vergroting op 1
k = 0.5
'Bepaal breedte en hoogte van het grafische veld
BeeldHoogte = 600
BeeldBreedte = 800
ZetBeeldgrootte
ROODINCREMENT = 19
GROENINCREMENT = 5
BLAUWINCREMENT = 11

'Begin met 1024 kleuren
MaxKleuren = 1024
GenereerPalet

' Definieer het beeldmidden
BeeldMiddenX = 0.5
BeeldMiddenY = 0

' Zet Zoomfactor initieel op 2
ZoomFactor = 0.5

'Toon het formulier
Form1.Show

'Genereer een nieuwe mandelbrot-fractal
Genereer

'Zet info in de statusbar
lblInfo.Caption = "x=" + Format(cx, "0.00000000000") + " y=" + Format(cy, "0.00000000000") + " Vergr.:" + Format(1 / ZoomFactor, "0.00") + "  Aantal kl.:" + Str(MaxKleuren)
End Sub
Private Sub ZetBeeldgrootte()
'Picture1.Width = BeeldBreedte
'Picture1.Height = BeeldHoogte
End Sub
'Genereer een kleurenpalet
Private Sub GenereerPalet()
Dim i As Integer

UpRood = True
UpGroen = True
UpBlauw = True
Rood = 0
Groen = 0
Blauw = 0

'genereer een kleurenpalet van 512 kleuren
For i = 1 To 1024
'rood
If UpRood Then
    Rood = Rood + ROODINCREMENT
Else
    Rood = Rood - ROODINCREMENT
End If
If Rood > 255 Then
    UpRood = False
    Rood = Rood - ROODINCREMENT
ElseIf Rood < 0 Then
    UpRood = True
    Rood = Rood + ROODINCREMENT
End If

' groen
If UpGroen Then
    Groen = Groen + GROENINCREMENT
Else
    Groen = Groen - GROENINCREMENT
End If
If Groen > 255 Then
    UpGroen = False
    Groen = Groen - GROENINCREMENT
ElseIf Groen < 0 Then
    UpGroen = True
    Groen = Groen + GROENINCREMENT
End If

' blauw
If UpBlauw Then
    Blauw = Blauw + BLAUWINCREMENT
Else
    Blauw = Blauw - BLAUWINCREMENT
End If
If Blauw > 255 Then
    UpBlauw = False
    Blauw = Blauw - BLAUWINCREMENT
ElseIf Blauw < 0 Then
    UpBlauw = True
    Blauw = Blauw + BLAUWINCREMENT
End If

Rem Palet(i) = RGB(4 * (i Mod 64), 16 * (i Mod 16), i Mod 256)
Palet(i) = RGB(Rood, Groen, Blauw)
Next i
End Sub

'Genereer een nieuw beeld
Private Sub Genereer()
Dim i, j As Integer
Form1.Caption = "Even geduld"
Begintijd = Time
'Doorloop alle pixels
For i = 0 To BeeldBreedte
For j = 0 To BeeldHoogte
  Teller = 0
  X = 0: Y = 0: x1 = 0: y1 = 0
  
  'cx en cy: coordinaten van het te berekenen punt
  cx = k * ((i - BeeldBreedte / 2) / 100) - BeeldMiddenX
  cy = k * ((j - BeeldHoogte / 2) / 100) - BeeldMiddenY

    'Loop tot afstand tot middelpunt > 2
    While (X * X + Y * Y < 9) And (Teller < MaxKleuren)
        x1 = X: y1 = Y
        X = x1 * x1 - y1 * y1 + cx
        Y = 2 * x1 * y1 + cy
        Teller = Teller + 1
    Wend

'Zet de kleuren volgens het aantal iteraties
If Teller = MaxKleuren Then
    Picture1.PSet (i, j), RGB(0, 0, 0) 'Bij max wordt het zwart
Else
    Picture1.PSet (i, j), Palet(Teller) 'Anders een kleur uit het kleurenpalet
End If
Next j, i
Eindtijd = Time
'Schrijf info in de Caption van het formulier
Form1.Caption = "X=" + Format(BeeldMiddenX, "0.00##########") + " Y=" + Format(BeeldMiddenY, "0.00###########") + "  Vergr.= " + Format(1 / k, "#######0") + "   Rekentijd:" + Format((Eindtijd - Begintijd) * 24 * 60 * 60, "##0") + " sec        (Klik rechtermuisknop voor HELP)"
End Sub



Private Sub mnuHelp_Click()
Dim s As String
Dim cl As String
cl = Chr(10) + Chr(13)
s = "Klik met de LINKERmuisknop om een nieuw beeld te genereren."
s = s + cl + "De muisaanwijzer wordt het centrum van het nieuwe beeld."
s = s + cl + cl + "Door te klikken met de RECHTERmuisknop bekomt u een popupmenu."
s = s + cl + "Hiermee kunt u instellen of u:"
s = s + cl + "        - Wil inzoomen of uitzoomen"
s = s + cl + "        - Het aantal kleuren wil instellen (minder kleuren rekent sneller)"
s = s + cl + "        - Refresh: het beeld opnieuw wil genereren."
s = s + cl + "        - Helemaal opnieuw wil beginnen"
s = s + cl + "        - Een ander kleurenpalet wil kiezen"
Form2.Caption = "HELP"
Form2.Label1.Caption = s
Form2.Show
End Sub

Private Sub mnuChangecolors_Click()
ROODINCREMENT = Rnd() * 29
GROENINCREMENT = Rnd() * 29
BLAUWINCREMENT = Rnd() * 29
GenereerPalet
Genereer
End Sub

Private Sub mnuGelijk_Click()
ZoomFactor = 1
lblInfo.Caption = "x=" + Format(cx, "0.00000000000") + " y=" + Format(cy, "0.00000000000") + " Vergr.:" + Format(1 / ZoomFactor, "0.00") + "  Aantal kl.:" + Str(MaxKleuren)
End Sub



Private Sub mnuZinx1_2_Click()
ZoomFactor = 1 / 1.2
lblInfo.Caption = "x=" + Format(cx, "0.00000000000") + " y=" + Format(cy, "0.00000000000") + " Vergr.:" + Format(1 / ZoomFactor, "0.00") + "  Aantal kl.:" + Str(MaxKleuren)
mnuZinx1_2.Checked = True
mnuInx2.Checked = False
mnuInx4.Checked = False
mnuInx10.Checked = False
End Sub
Private Sub mnuInx10_Click()
ZoomFactor = 0.1
lblInfo.Caption = "x=" + Format(cx, "0.00000000000") + " y=" + Format(cy, "0.00000000000") + " Vergr.:" + Format(1 / ZoomFactor, "0.00") + "  Aantal kl.:" + Str(MaxKleuren)
mnuZinx1_2.Checked = False
mnuInx2.Checked = False
mnuInx4.Checked = False
mnuInx10.Checked = True
End Sub

Private Sub mnuInx2_Click()
ZoomFactor = 0.5
lblInfo.Caption = "x=" + Format(cx, "0.00000000000") + " y=" + Format(cy, "0.00000000000") + " Vergr.:" + Format(1 / ZoomFactor, "0.00") + "  Aantal kl.:" + Str(MaxKleuren)
mnuZinx1_2.Checked = False
mnuInx2.Checked = True
mnuInx4.Checked = False
mnuInx10.Checked = False
End Sub

Private Sub mnuInx4_Click()
ZoomFactor = 0.25
lblInfo.Caption = "x=" + Format(cx, "0.00000000000") + " y=" + Format(cy, "0.00000000000") + " Vergr.:" + Format(1 / ZoomFactor, "0.00") + "  Aantal kl.:" + Str(MaxKleuren)
mnuZinx1_2.Checked = False
mnuInx2.Checked = False
mnuInx4.Checked = True
mnuInx10.Checked = False
End Sub



Private Sub mnuK32_Click()
MaxKleuren = 32
lblInfo.Caption = "x=" + Format(cx, "0.00000000000") + " y=" + Format(cy, "0.00000000000") + " Vergr.:" + Format(1 / ZoomFactor, "0.00") + "  Aantal kl.:" + Str(MaxKleuren)
mnuK32.Checked = True
mnuK128.Checked = False
mnuK512.Checked = False
mnuK1024.Checked = False
End Sub
Private Sub mnuK128_Click()
MaxKleuren = 128
lblInfo.Caption = "x=" + Format(cx, "0.00000000000") + " y=" + Format(cy, "0.00000000000") + " Vergr.:" + Format(1 / ZoomFactor, "0.00") + "  Aantal kl.:" + Str(MaxKleuren)
mnuK32.Checked = False
mnuK128.Checked = True
mnuK512.Checked = False
mnuK1024.Checked = False
End Sub
Private Sub mnuK512_Click()
MaxKleuren = 512
lblInfo.Caption = "x=" + Format(cx, "0.00000000000") + " y=" + Format(cy, "0.00000000000") + " Vergr.:" + Format(1 / ZoomFactor, "0.00") + "  Aantal kl.:" + Str(MaxKleuren)
mnuK32.Checked = False
mnuK128.Checked = False
mnuK512.Checked = True
mnuK1024.Checked = False
End Sub
Private Sub mnuK1024_Click()
MaxKleuren = 1024
lblInfo.Caption = "x=" + Format(cx, "0.00000000000") + " y=" + Format(cy, "0.00000000000") + " Vergr.:" + Format(1 / ZoomFactor, "0.00") + "  Aantal kl.:" + Str(MaxKleuren)
mnuK32.Checked = False
mnuK128.Checked = False
mnuK512.Checked = False
mnuK1024.Checked = True
End Sub
Private Sub mnuRefresh_Click()
Genereer
End Sub

Private Sub mnuRestart_Click()
Dim ans As Integer
ans = MsgBox("Helemaal opnieuw beginnen?", vbQuestion + vbYesNo, "Opgepast!")
If ans = vbYes Then
    BeeldMiddenX = 0.5
    BeeldMiddenY = 0
    ZoomFactor = 0.5
    MaxKleuren = 32
    k = 1
    Genereer
    lblInfo.Caption = Format(cx, "0.0##########") + " " + Format(cy, "0.0##########") + " x " + Format(1 / ZoomFactor, "0.0#")
    End If
End Sub

Private Sub mnuUitx10_Click()
ZoomFactor = 10
lblInfo.Caption = "x=" + Format(cx, "0.00000000000") + " y=" + Format(cy, "0.00000000000") + " Vergr.:" + Format(1 / ZoomFactor, "0.00") + "  Aantal kl.:" + Str(MaxKleuren)
End Sub

Private Sub mnuUitx2_Click()
ZoomFactor = 2
lblInfo.Caption = "x=" + Format(cx, "0.00000000000") + " y=" + Format(cy, "0.00000000000") + " Vergr.:" + Format(1 / ZoomFactor, "0.00") + "  Aantal kl.:" + Str(MaxKleuren)
End Sub

Private Sub mnuUitx4_Click()
ZoomFactor = 4
lblInfo.Caption = "x=" + Format(cx, "0.00000000000") + " y=" + Format(cy, "0.00000000000") + " Vergr.:" + Format(1 / ZoomFactor, "0.00") + "  Aantal kl.:" + Str(MaxKleuren)
End Sub


Private Sub mnuZuitx1_2_Click()
ZoomFactor = 1.2
lblInfo.Caption = "x=" + Format(cx, "0.00000000000") + " y=" + Format(cy, "0.00000000000") + " Vergr.:" + Format(1 / ZoomFactor, "0.00") + "  Aantal kl.:" + Str(MaxKleuren)
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Genereer een nieuw beeld (linker muisknop)
If Button = 1 Then
    BeeldMiddenX = -k * ((X - BeeldBreedte / 2) / 100) + BeeldMiddenX
    BeeldMiddenY = -k * ((Y - BeeldHoogte / 2) / 100) + BeeldMiddenY
    k = k * ZoomFactor
    Genereer
'of geef popup-menu (rechter muisknop)
Else
    PopupMenu Form1.mnuZoomMenu
End If
End Sub

'Toon lopende coordinaten in de statusbar bij het bewegen van de muis
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cx = -k * ((X - BeeldBreedte / 2) / 100) + BeeldMiddenX
cy = -k * ((Y - BeeldHoogte / 2) / 100) + BeeldMiddenY
lblInfo.Caption = "x=" + Format(cx, "0.00000000000") + " y=" + Format(cy, "0.00000000000") + " Vergr.:" + Format(1 / ZoomFactor, "0.00") + "  Aantal kl.:" + Str(MaxKleuren)
End Sub
