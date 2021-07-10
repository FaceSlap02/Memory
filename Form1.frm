VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Memory Game"
   ClientHeight    =   13275
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   15600
   LinkTopic       =   "Form1"
   ScaleHeight     =   13275
   ScaleWidth      =   15600
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ListBox ordered_list 
      Height          =   2985
      Left            =   240
      TabIndex        =   14
      Top             =   1560
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Timer Timer4 
      Interval        =   100
      Left            =   480
      Top             =   9000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "NEXT ROUND"
      Height          =   735
      Left            =   22560
      TabIndex        =   13
      Top             =   -120
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   480
      Top             =   8520
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      DragIcon        =   "Form1.frx":0000
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2760
      TabIndex        =   1
      Top             =   8400
      Width           =   14895
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "STATISTICS"
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Points (click)"
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   975
         Left            =   12360
         TabIndex        =   11
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   12240
         TabIndex        =   10
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "RE-PLAYS LEFT (CLICK)"
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   855
         Left            =   8760
         TabIndex        =   9
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9000
         TabIndex        =   8
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Round"
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   4800
         TabIndex        =   7
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5280
         TabIndex        =   6
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "moves"
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   2640
         TabIndex        =   5
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   4
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "level (click)"
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   975
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   480
      Top             =   8040
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   480
      Top             =   7560
   End
   Begin VB.PictureBox box 
      Height          =   1095
      Index           =   1000
      Left            =   3240
      ScaleHeight     =   1035
      ScaleWidth      =   2715
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   2055
      Left            =   2760
      TabIndex        =   15
      Top             =   8520
      Width           =   14895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   6720
      TabIndex        =   12
      Top             =   4560
      Width           =   10815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" _
 (ByVal lpszName As String, ByVal hModule As Long, _
 ByVal dwFlags As Long) As Long

Dim attributes_interval As Integer
Dim hide_attributes As Boolean
Dim loaded As Boolean

Dim WithEvents slika As image
Attribute slika.VB_VarHelpID = -1
Dim starting_round_moves As Integer



Private Sub box_Click(index As Integer)
    If (igra = False And loaded = True) Then ' sprijecimo spamovanje klika dok timer traje sekundu
        Timer1.Enabled = False
        Timer1.Interval = 1000
        
        box(index).Visible = False
        
        If (firstClick = True) Then
            Dim position As Boolean
            
            If (index Mod 2 = 0) Then
                position = index = lastClick + 1
            Else
                position = index = lastClick - 1
            End If
        
                    
            moves = moves + 1
                     
            If (position = False) Then
                timer_index_1 = lastClick
                timer_index_2 = index
                
                Timer1.Enabled = True
                
                If (failed_once = True) Then
                    failed_once = False
                    
                    If (points >= 0.8) Then
                        points = points - 0.8
                    Else
                        points = 0
                    End If
                Else
                    failed_once = True
                End If
    
                igra = True
            Else
                points = points + 1.5
                failed_once = False
                
                guesses = guesses + 1
                
                PlaySound App.Path & "\level_up.wav", 0, 1
                
                If (guesses >= (image_count / 2)) Then
                    Dim message As String
                    
                    message = MsgBox("Runda je zavrsena (" & rounds_played & "/3), da li zelis dalje?", vbYesNo, "Game")
                    
                    If (message = "6") Then
                        re_plays = 2
                        hideAttributes
                        loadGame
                        showAttributes
                    End If
                        
                    MsgBox "Nova igra je ucitana", vbOKOnly, "Game"
                End If
            
            End If
            
            firstClick = False
        Else
            firstClick = True
        End If
        
        lastClick = index
    End If
End Sub

Private Function hideImage(index As Integer)
    On Error Resume Next:
        Dim image As image
            
        Set image = Me.Controls("slika" & index)
            
        If (Not image Is Nothing) Then
            image.Visible = False
        End If
End Function

Private Function showImage(index As Integer)
    On Error Resume Next:
        Dim image As image
            
        Set image = Me.Controls("slika" & index)
            
        If (Not image Is Nothing) Then
            image.Visible = True
        End If
End Function

Private Sub Command1_Click()
level = level + 1
If (level > 3) Then
    level = 1
End If
loadGame

End Sub

Private Sub Command2_Click()

  hideAttributes
  loadGame
  showAttributes

End Sub

Private Sub Form_Load()

Timer1.Enabled = False
Timer2.Enabled = False
Timer4.Enabled = False

image_count = 12

Form1.Picture = LoadPicture(App.Path & "\background.jpg")

Frame1.BackColor = RGB(82, 187, 196)
Label2.BackColor = RGB(82, 187, 196)
Frame1.Visible = False
Label2.Visible = False

Label2.height = Frame1.height + 80
Label2.width = Frame1.width + 75

start_countdown = 5

Timer2.Interval = 500
Timer4.Interval = 100

loadGame

Label3.ForeColor = RGB(16, 82, 110)

Timer2.Enabled = True

Form1.height = Screen.height
Form1.width = Screen.width

End Sub

Private Function hideAttributes()

    attributes_interval = 0
    hide_attributes = True
    
    Timer4.Enabled = True
                                 
End Function

Private Function showAttributes()

    attributes_interval = 0
    hide_attributes = False
    
    Timer4.Enabled = True
    
End Function
Private Function removeImage(index As Integer)
On Error Resume Next:
    Me.Controls.Remove ("slika" & index)
End Function

Private Function loadGame()

If (level = 3 And rounds_played = 3) Then
    MsgBox "Igra je zavrsena" & vbCrLf & " " & vbCrLf & "Total moves made: " & moves & vbCrLf & "Total points earned: " & points & vbCrLf & " " & vbCrLf & "Thank you for playing!", vbOKOnly, "Game"
    End
End If

Timer1.Enabled = False

ordered_list.Clear
guesses = 0

If (level = 0) Then
    rounds_played = 1
    level = 1
End If

If (previous_level <> level) Then
    re_plays = 2
    rounds_played = 1
End If

If (previous_level = level) Then
    rounds_played = rounds_played + 1
End If
    
If (rounds_played > 3) Then
    rounds_played = 1
    level = level + 1
End If

If (level > 3) Then
    level = 1
End If
If (level = 1) Then
    If (rounds_played = 1) Then
        image_count = 12
    ElseIf (rounds_played = 2) Then
        image_count = 16
    ElseIf (rounds_played = 3) Then
        image_count = 20
    End If
End If

If (level = 2) Then
    If (rounds_played = 1) Then
        image_count = 16
    ElseIf (rounds_played = 2) Then
        image_count = 16
    ElseIf (rounds_played = 3) Then
        image_count = 20
    End If
End If

If (level = 3) Then
    If (rounds_played = 1) Then
        image_count = 20
    ElseIf (rounds_played = 2) Then
        image_count = 20
    ElseIf (rounds_played = 3) Then
        image_count = 20
    End If
End If


previous_level = level
starting_round_moves = moves

Dim previous As image

Dim niz(50) As Integer

Dim X As Integer ' nesumican broj
Dim brojac As Integer

Dim k As Integer

If (box.Count > 1) Then
    For k = 1 To (box.Count - 1)
        Unload box(k)
    Next k
End If

'''''''''''''''''''''''''''''''''''''''''
' Uklanjamo svih 20 slika jer postoji   '
' mogucnost runde 16 -> 20              '
' kako ne bi 4 slike ostale, za svaku   '
' rundu uklanjamo svih 20 slika s tim   '
' da cemo ignorisat ako slika ne postoji'
                                        '
Dim n As Integer                        '
For n = 1 To 20                         '
    removeImage (n)                     '
Next n                                  '
                                        '
'''''''''''''''''''''''''''''''''''''''''


For brojac = 1 To image_count
Dim i As Integer
Dim position As Integer

Randomize
    
Start:
    Randomize (Timer)
    X = Int(image_count * Rnd + 1)
    For i = 1 To image_count
      If X = niz(i) Then
        GoTo Start 'ako postoji broj ponovo odaberi
      End If
    Next i
    
    niz(brojac) = X ' smjesti odabrani broj u niz

    Load box(X)

    Set slika = Me.Controls.Add("VB.Image", "slika" & X)
    
    slika.Left = 1000
    slika.Top = 1000
    
    ordered_list.AddItem (X)
    
    box(X).Left = 1000
    box(X).Top = 1000
    
    slika.Picture = LoadPicture(App.Path & "\" & level & "\" & X & ".jpg")
    slika.Stretch = True

    Dim space_right As Integer
    Dim space_down As Integer
    Dim width As Integer
    Dim height As Integer
    
    space_right = 2550
    space_down = 2050
    
    width = 2500
    height = 2000
    
    If (previous Is Nothing = False) Then
        slika.Left = previous.Left + space_right
        slika.Top = previous.Top
        
        box(X).Left = previous.Left + space_right
        box(X).Top = previous.Top
        
        last_top_value = previous.Top
    End If
    
    If (image_count = 12 Or image_count = 16) Then
        If (position = 4 Or position = 8 Or position = 12) Then
            slika.Left = 1000
            slika.Top = previous.Top + space_down
            box(X).Left = 1000
            box(X).Top = previous.Top + space_down
        End If
    ElseIf (image_count = 20) Then
        If (position = 5 Or position = 10 Or position = 15) Then
            slika.Left = 1000
            slika.Top = previous.Top + space_down
            box(X).Left = 1000
            box(X).Top = previous.Top + space_down
        End If
    End If

    
    position = position + 1
    
    Set previous = slika
        
    slika.width = width
    slika.height = height
    
    box(X).width = width
    box(X).height = height

    box(X).BackColor = RGB(82, 187, 196)
    
    slika.Visible = False
    box(X).Visible = False
    
Next brojac

If (image_count = 12 Or image_count = 16) Then
    Frame1.Left = 500
    Label2.Left = 440
ElseIf (image_count = 20) Then
    Frame1.Left = 1500
    Label2.Left = 1440
End If

Frame1.Top = last_top_value + 3000
Label2.Top = last_top_value + 2940
Frame1.Visible = False
Label2.Visible = False


End Function

Private Sub Label10_Click()
Dim a As String

a = MsgBox("You will be able to re-play round two times per level!" & vbCrLf & " " & vbCrLf & "Do you want to use reply right now?", vbYesNo, "Replays")

If (a = "6") Then
    If (re_plays > 0) Then
        hideAttributes
        re_plays = re_plays - 1
        rounds_played = rounds_played - 1
        moves = starting_round_moves
        failed_once = False
    
        loadGame
        showAttributes
    Else
        MsgBox "You don't have enough replays", vbOKOnly, "Replays"
    End If
    
End If

End Sub

Private Sub Label12_Click()
MsgBox "You receive 1.5 points per guess and you lose 0.80 points every second fail! Points can't go negative.", vbOKOnly, "Points"
End Sub

Private Sub Label4_Click()
MsgBox "Levels information" & vbCrLf & vbCrLf & " " & vbCrLf & "First Level" & vbCrLf & " [-] 3 rounds" & vbCrLf & " [-] Images count: 12, 16, 20" & vbCrLf & " " & vbCrLf & "Second Level" & vbCrLf & " [-] 3 rounds" & vbCrLf & " [-] Images count: 16, 16, 20" & vbCrLf & " " & vbCrLf & "Third Level" & vbCrLf & " [-] 3 rounds" & vbCrLf & " [-] Images count: 20, 20, 20", vbOKOnly, "Levels"
End Sub

Private Sub Timer1_Timer()
    box(timer_index_1).Visible = True
    box(timer_index_2).Visible = True
    igra = False
End Sub

Private Sub Timer2_Timer()
 
    start_countdown = start_countdown - 1
    
    If (Label3.Caption = "Loading") Then
        Label3.Caption = "Loading."
    ElseIf (Label3.Caption = "Loading.") Then
        Label3.Caption = "Loading.."
    ElseIf (Label3.Caption = "Loading..") Then
        Label3.Caption = "Loading..."
    ElseIf (Label3.Caption = "Loading...") Then
        Label3.Caption = "Loading"
    End If

    If (start_countdown >= 0) Then
        Timer2.Enabled = True
    Else
        Label3.Visible = False
        Timer2.Enabled = False
        showAttributes
    End If
End Sub

Private Sub Timer3_Timer()
    Label1.Caption = level
    Label5.Caption = moves
    Label7.Caption = rounds_played & "/3"
    Label9.Caption = re_plays & "/2"
    Label11.Caption = points
Timer3.Enabled = True
End Sub

Private Sub Timer4_Timer()
    If (attributes_interval >= image_count) Then
        Frame1.Visible = hide_attributes = False
        Label2.Visible = hide_attributes = False
        Timer4.Enabled = False
        loaded = True
    Else
        Timer4.Enabled = True

        Dim value As Integer
        
        If (hide_attributes = True) Then
            value = Int(ordered_list.List(attributes_interval))
            
            box(value).Visible = False
            hideImage (value)
        Else
            value = Int(ordered_list.List(attributes_interval))
            
            box(value).Visible = True
            showImage (value)
        End If
    End If
    
    attributes_interval = attributes_interval + 1
      
End Sub
