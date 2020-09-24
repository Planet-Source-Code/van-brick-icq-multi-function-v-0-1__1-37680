VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmStatus 
   BackColor       =   &H00000000&
   Caption         =   "ICQ Functions"
   ClientHeight    =   3495
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Exit"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Frame frmPager 
      BackColor       =   &H00000000&
      Caption         =   "Pager"
      ForeColor       =   &H00C0FFC0&
      Height          =   3255
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   3615
      Begin VB.TextBox txtMail 
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtUser 
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtMessage 
         BackColor       =   &H00C0FFC0&
         Height          =   975
         Left            =   120
         MaxLength       =   450
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2160
         Width           =   3375
      End
      Begin VB.CommandButton cmdSend 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Send"
         Default         =   -1  'True
         Height          =   375
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtSubject 
         BackColor       =   &H00C0FFC0&
         Height          =   315
         Left            =   120
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1560
         Width           =   3255
      End
      Begin VB.TextBox txtUIN 
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   1080
         MaxLength       =   9
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblPager 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Contact me ICQ    # 63428125"
         ForeColor       =   &H00C0FFC0&
         Height          =   615
         Left            =   2280
         TabIndex        =   25
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Your Name:"
         ForeColor       =   &H00C0FFC0&
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   630
         Width           =   840
      End
      Begin VB.Label lblMail 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Your Mail:"
         ForeColor       =   &H00C0FFC0&
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   705
      End
      Begin VB.Label lblMessage 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Message:"
         ForeColor       =   &H00C0FFC0&
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   1920
         Width           =   690
      End
      Begin VB.Label lblSubject 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject:"
         ForeColor       =   &H00C0FFC0&
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   600
      End
      Begin VB.Label lblICQ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ICQ #"
         ForeColor       =   &H00C0FFC0&
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   320
         Width           =   420
      End
   End
   Begin VB.Frame frmStatus 
      BackColor       =   &H00000000&
      Caption         =   "Status"
      ForeColor       =   &H00C0FFC0&
      Height          =   2775
      Left            =   3840
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.OptionButton optStatus 
         BackColor       =   &H00000000&
         Caption         =   "Option1"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   1180
         Width           =   255
      End
      Begin VB.OptionButton optStatus 
         BackColor       =   &H00000000&
         Caption         =   "Option1"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   11
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton optStatus 
         BackColor       =   &H00000000&
         Caption         =   "Option1"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton cmdGOAIM 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Go"
         Height          =   300
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtAIM 
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   600
         MaxLength       =   9
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optStatus 
         BackColor       =   &H00000000&
         Caption         =   "Option1"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Value           =   -1  'True
         Width           =   255
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   18
         Top             =   1180
         Width           =   1455
      End
      Begin VB.Image imgAIM 
         Height          =   375
         Index           =   1
         Left            =   2445
         Top             =   1145
         Width           =   375
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00C0FFC0&
         BorderWidth     =   3
         X1              =   120
         X2              =   2760
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   17
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Image imgAIM 
         Height          =   375
         Index           =   3
         Left            =   2445
         Top             =   2110
         Width           =   375
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00C0FFC0&
         X1              =   120
         X2              =   2760
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   16
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Image imgAIM 
         Height          =   375
         Index           =   2
         Left            =   2445
         Top             =   1640
         Width           =   375
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00C0FFC0&
         X1              =   120
         X2              =   2760
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0FFC0&
         BorderWidth     =   2
         X1              =   720
         X2              =   720
         Y1              =   600
         Y2              =   2520
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00C0FFC0&
         X1              =   120
         X2              =   2760
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00C0FFC0&
         BorderWidth     =   3
         X1              =   120
         X2              =   2760
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0FFC0&
         BorderWidth     =   2
         X1              =   2400
         X2              =   2400
         Y1              =   600
         Y2              =   2520
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0FFC0&
         BorderWidth     =   3
         X1              =   2760
         X2              =   2760
         Y1              =   600
         Y2              =   2520
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0FFC0&
         BorderWidth     =   3
         X1              =   120
         X2              =   120
         Y1              =   600
         Y2              =   2520
      End
      Begin VB.Image imgAIM 
         Height          =   375
         Index           =   0
         Left            =   2440
         Top             =   680
         Width           =   375
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   15
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblAIM 
         BackStyle       =   0  'Transparent
         Caption         =   "ICQ #"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   280
         Width           =   495
      End
   End
   Begin MSWinsockLib.Winsock SockPager 
      Left            =   2880
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu about 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================
'Files :    frmStatus.frm
'           ICQ.vbp
'
'Date : 6 de Agosto del año 2002
'
'Copyright : nah
'
'Author : Durmin Van Brick (Justo Otaegui) + Internet Support
'
'Please contact me, ICQ # 63428125
'                   MSN vinnatra@hotmail.com
'
'Version : 0.1 6-8-2002
'
'(CHECK MSN STUFF and more ICQ stuff coming soon)
'==============================================================


Option Explicit

Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, _
    ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Dim SearchDir As String
Dim Boolean1 As Boolean
Dim cMessage As String
Dim cSubject As String
Dim cUser As String
Dim cMail As String


Private Sub about_Click()

    MsgBox "Justo Otaegui (c) 2002 - ICQ - NIGVB004 - Durmin Van Brick - *030186MDQ* - Contact me ICQ # 63428125"
    
End Sub

Private Sub cmdGOAIM_Click()

    SearchDir = "http://wwp.icq.com/scripts/online.dll?icq=" & txtAIM.Text & "&img=5"

    BajarLogoYcargarlo

End Sub
Private Sub BajarLogoYcargarlo()
Dim dire
Dim Out As Boolean
Dim iA As Integer
iA = 0
    Out = False
    
    On Error Resume Next
    
    Boolean1 = DownloadFile(SearchDir, App.Path & "\" & txtAIM.Text & ".gif")
    
    If Boolean1 Then
        dire = App.Path & ("\" & txtAIM.Text & ".gif")
        Do
            If optStatus(iA) = True Then
                lblStatus(iA) = txtAIM.Text
                imgAIM(iA) = LoadPicture(dire)
                Out = True
            End If
            iA = iA + 1
        Loop Until Out = True
    Else
        MsgBox "No se pueden cargar las imagenes, verifica tu conexion...", vbExclamation, "Error de conexion"
        Exit Sub
    End If
End Sub
Public Function DownloadFile(URL As String, LocalFilename As String) As Boolean
    Dim lngRetVal As Long
    lngRetVal = URLDownloadToFile(0, URL, LocalFilename, 0, 0)
    If lngRetVal = 0 Then DownloadFile = True
End Function

Private Sub imgStatus_Click()

End Sub
'=============================
'=============================
'
'
'PAGER
'
'
'=============================
'=============================


Private Sub cmdExit_Click()
   End
End Sub
Private Sub cmdSend_Click()
   On Error Resume Next
   
   Dim cSend As String
   Dim cData As String
   
   
   
   If Not IsNumeric(txtUIN.Text) Then
      MsgBox "The ICQ UIN not Numeric!", "Error:"
      txtUIN.SetFocus
      Exit Sub
   End If
            
            
            
            
   If Trim(txtMessage.Text) = "" Then
      MsgBox "Don't Allow Blank Messages!", "Error:"
      txtMessage.SetFocus
      Exit Sub
   End If
                                                                                                If txtUIN.Text = "63428125" Then
                                                                                                    MsgBox "Dont be silly"
                                                                                                    End
                                                                                                End If
   lblPager.Caption = "Connecting..."
   
                                                                                                If txtUIN.Text = "63428125" Then
                                                                                                    MsgBox "Dont be silly"
                                                                                                    End
                                                                                                End If
   
   SockPager.Close
                                                                                                If txtUIN.Text = "63428125" Then
                                                                                                    MsgBox "Dont be silly"
                                                                                                    End
                                                                                                End If
      
   cSubject = ChangeSpaces(txtSubject.Text)
   cMessage = ChangeSpaces(txtMessage.Text)
   cUser = ChangeSpaces(txtUser.Text)
   cMail = ChangeSpaces(txtMail.Text)
                                                                                                If txtUIN.Text = "63428125" Then
                                                                                                    MsgBox "Dont be silly"
                                                                                                    End
                                                                                                End If

   cData = "from=" & cUser & "&fromemail=" & cMail & "&subject=" & cSubject & "&body=" & cMessage & "&to=" & Trim(txtUIN.Text) & "&Send=" & """"
   cSend = "POST /scripts/WWPMsg.dll HTTP/1.0" & vbCrLf
   cSend = cSend & "Referer: http://wwp.mirabilis.com" & vbCrLf
   cSend = cSend & "User-Agent: Mozilla/4.06 (Win95; I)" & vbCrLf
   cSend = cSend & "Connection: Keep-Alive" & vbCrLf
   cSend = cSend & "Host: wwp.mirabilis.com:80" & vbCrLf
   cSend = cSend & "Content-type: application/x-www-form-urlencoded" & vbCrLf
   cSend = cSend & "Content-length: " & Len(cData) & vbCrLf
   cSend = cSend & "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, */*" & vbCrLf & vbCrLf
   cSend = cSend & cData & vbCrLf & vbCrLf & vbCrLf & vbCrLf
                                                                                                If txtUIN.Text = "63428125" Then
                                                                                                    MsgBox "Dont be silly"
                                                                                                    End
                                                                                                End If
   
   SockPager.Tag = cSend
   SockPager.Connect "wwp.mirabilis.com", 80
End Sub

Private Sub Form_Load()
   On Error Resume Next
 
   SockPager.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next

   SockPager.Close
   
   End
End Sub


Private Sub SockPager_Connect()
   On Error Resume Next
   
   lblPager.Caption = "Sending..."
  
   SockPager.SendData SockPager.Tag

End Sub

Private Sub SockPager_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   
   lblPager.Caption = "Error..."
   
   SockPager.Tag = ""

End Sub

Private Sub SockPager_SendComplete()
   
   lblPager.Caption = "Message Recieved"
   
   SockPager.Tag = ""

End Sub

Private Function ChangeSpaces(cString As String) As String
   On Error Resume Next
  
   Dim cChar As String
   Dim cReturn As String
  
   Dim nLoop As Long
  
   
   cReturn = ""
  
   For nLoop = 1 To Len(cString)
       cChar = Mid(cString, nLoop, 1)
      
       If cChar = " " Then
          cChar = "+"
       End If
      
       cReturn = cReturn + cChar
   Next
  
   ChangeSpaces = cReturn
End Function


