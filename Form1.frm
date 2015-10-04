VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000004&
   ClientHeight    =   3720
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14535
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   14535
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtChkLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   0
      Left            =   7800
      TabIndex        =   14
      Text            =   "Choice 1"
      Top             =   1635
      Width           =   4095
   End
   Begin VB.CheckBox chkCheckbox 
      Caption         =   "Check1"
      Height          =   195
      Index           =   0
      Left            =   7440
      TabIndex        =   13
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox txtAnswer 
      Height          =   285
      Index           =   0
      Left            =   7440
      TabIndex        =   12
      Top             =   1200
      Width           =   4455
   End
   Begin VB.TextBox txtTitleDescription 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   240
      TabIndex        =   9
      Text            =   "Form description"
      Top             =   480
      Width           =   6735
   End
   Begin VB.TextBox txtTitle 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   240
      TabIndex        =   8
      Text            =   "Untitled forms"
      Top             =   120
      Width           =   6255
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   6975
      Begin VB.CommandButton cmdAddMoreChk 
         Caption         =   "Add More Choices"
         Height          =   315
         Left            =   3600
         TabIndex        =   15
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtQuestion 
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   600
         Width           =   4695
      End
      Begin VB.TextBox txtQuestionDescripton 
         Height          =   405
         Left            =   1800
         TabIndex        =   2
         Top             =   1080
         Width           =   4695
      End
      Begin VB.ComboBox cboInputType 
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   1800
         List            =   "Form1.frx":0013
         TabIndex        =   1
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Question Title"
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Help Text"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   1155
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Question Type"
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   1620
         Width           =   1455
      End
   End
   Begin VB.Label lblQuestionDescripton 
      Caption         =   "Label4"
      Height          =   255
      Index           =   0
      Left            =   7440
      TabIndex        =   11
      Top             =   720
      Width           =   6975
   End
   Begin VB.Label lblQuestion 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   7440
      TabIndex        =   10
      Top             =   360
      Width           =   6855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i, Y, toping As Long

Private Sub cboInputType_Click()
    If cboInputType.ListIndex = 2 Then
        cmdAddMoreChk.Visible = True
        cmdAddMoreChk.Enabled = False
    Else
        cmdAddMoreChk.Visible = False
    End If
End Sub

Private Sub cmdAddMoreChk_Click()
        Load chkCheckbox(i + 1)
        With chkCheckbox(i + 1)
            .Left = 240
            .Top = toping + 350
            toping = toping + 350
            .Visible = True
        End With
        Load txtChkLabel(i + 1)
        With txtChkLabel(i + 1)
            .Left = 560
            .Top = toping - 28
            .Visible = True
        End With
        Form1.Height = Height + 350
        Frame1.Top = Y + 350
        Y = Y + 350
        i = i + 1
End Sub

Private Sub Form_Load()
    i = 0
    Y = 720
    toping = 450
    lblQuestion(0).Visible = False
    lblQuestionDescripton(0).Visible = False
    Form1.Width = 7515
    cboInputType.ListIndex = 0
    cmdAddMoreChk.Visible = False
End Sub

Private Sub cmdAdd_Click()
    Frame1.Top = Y + 750
    Form1.Height = Height + 750
    Y = Frame1.Top
      
    If txtQuestionDescripton.Text <> "" Then
        Load lblQuestion(i + 1)
        With lblQuestion(i + 1)
            .Caption = txtQuestion.Text
            .Left = 240
            .Top = toping + 450
            toping = toping + 450
            .Visible = True
            
        End With
    

        Load lblQuestionDescripton(i + 1)
        With lblQuestionDescripton(i + 1)
            .Caption = txtQuestionDescripton.Text
            .Left = 240
            .Top = toping + 250
            toping = toping + 250
            .Visible = True
            Frame1.Top = Y + 250
            Form1.Height = Height + 250
            Y = Frame1.Top
        End With
     
        
    Else
        Load lblQuestion(i + 1)
        With lblQuestion(i + 1)
            .Caption = txtQuestion.Text
            .Left = 240
            .Top = toping + 450
            toping = toping + 450
            .Visible = True
            
        End With
        
    End If
    
    If cboInputType.ListIndex = 0 Then
    
        Load txtAnswer(i + 1)
        With txtAnswer(i + 1)
            .Left = 240
            .Top = toping + 250
            toping = toping + 250
            .Visible = True
            
        End With
    ElseIf cboInputType.ListIndex = 1 Then

    ElseIf cboInputType.ListIndex = 2 Then
        Load chkCheckbox(i + 1)
        With chkCheckbox(i + 1)
            .Left = 240
            .Top = toping + 250
            toping = toping + 250
            .Visible = True
        End With
        Load txtChkLabel(i + 1)
        With txtChkLabel(i + 1)
            .Left = 560
            .Top = toping - 28
            .Visible = True
        End With
        cmdAddMoreChk.Visible = True
        cmdAddMoreChk.Enabled = True
    End If
    txtQuestion.Text = ""
    txtQuestionDescripton.Text = ""
    i = i + 1
End Sub



Private Sub txtTitle_Click()
    txtTitle.Text = ""
End Sub

Private Sub txtTitleDescription_Click()
    txtTitleDescription.Text = ""
End Sub
