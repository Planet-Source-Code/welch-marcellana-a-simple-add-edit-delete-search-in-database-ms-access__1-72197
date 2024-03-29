VERSION 5.00
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSearch 
      Height          =   330
      Left            =   1200
      TabIndex        =   1
      Top             =   660
      Width           =   3255
   End
   Begin VB.ComboBox cmbSearch 
      Height          =   345
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
   Begin WelchButton.lvButtons_H cmdSearch 
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Search"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Enter Text:"
      Height          =   225
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Search by:"
      Height          =   225
      Left            =   240
      TabIndex        =   2
      Top             =   300
      Width           =   855
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'CODED BY:  Welch Regime Marcellana
'I hope that my code will help you
'JOIN IN MY FORUM SITE, IT'S FREE TO REGISTER!!.
'Post topic about VB Tutorials, Love/Relationships, Careers/At the Job,
'Movie, Music etc.
'www.thesacrificiallamb.com
'This is a new website and currently looking for members.
'Your registration is very much appreciated :)  Thank you.

Private Sub cmdSearch_Click()
  If Me.cmbSearch.Text = "" Or Me.txtSearch.Text = "" Then
    MsgBox "All fields are required!", vbExclamation, "Error"
    Exit Sub
  End If
  
  Select Case LCase(Me.cmbSearch.Text)
    Case "name"
      SQL = "SELECT tbl_info.info_name, tbl_info.id_info, tbl_info.info_name, tbl_info.info_address " & _
            "From tbl_info WHERE (((tbl_info.info_name) Like '" & Me.txtSearch.Text & "%')) order by tbl_info.info_name asc;"
      Unload Me
      Call goSearch(SQL)

    Case "address"
      SQL = "SELECT tbl_info.info_name, tbl_info.id_info, tbl_info.info_name, tbl_info.info_address " & _
            "From tbl_info WHERE (((tbl_info.info_address) Like '" & Me.txtSearch.Text & "%')) order by tbl_info.info_name asc;"
      Unload Me
      Call goSearch(SQL)
  End Select
End Sub

Private Sub Form_Load()
  With Me.cmbSearch
    .AddItem "Name"
    .AddItem "Address"
  End With
End Sub

Public Sub goSearch(theSQL As String)
  Dim totRes As Long
  
  Form1.ListView.ListItems.Clear
  I = 0
  frmLoading.Show
  frmLoading.lblSub.Caption = "Searching..."
  totRes = countResults(theSQL)
  Call dbConnect
    RS.Open theSQL, Conn, adOpenDynamic
      If Not RS.EOF Then
        RS.MoveFirst
        Do While Not RS.EOF
          I = I + 1
          With Form1.ListView.ListItems
            Set Item = .Add(, , RS!id_info)
              Item.SubItems(1) = RS!info_name
              Item.SubItems(2) = RS!info_address
          End With
          frmLoading.lblSub.Caption = "Displaying Results:  " & I & " of " & totRes
          RS.MoveNext
          DoEvents
        Loop
      End If
    RS.Close
  Conn.Close
  Set Conn = Nothing
  Unload frmLoading
End Sub

Public Function countResults(theSQL2 As String) As Long
  Call dbConnect
     RS.Open SQL, Conn, adOpenDynamic
       If Not RS.EOF Then
         RS.MoveFirst
         Do While Not RS.EOF
           frmLoading.lblSub.Caption = "Total records found:  " & countResults
           countResults = countResults + 1
           RS.MoveNext
           DoEvents
         Loop
       End If
     RS.Close
  Conn.Close
  Set Conn = Nothing
End Function
