VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVBUTTONS.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BlackBook - Get Organized!"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13500
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   13500
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Contact Info / Add a Contact"
      Height          =   8295
      Left            =   8760
      TabIndex        =   9
      Top             =   120
      Width           =   4575
      Begin LVbuttons.LaVolpeButton cmdClear 
         Height          =   375
         Left            =   3000
         TabIndex        =   20
         Top             =   7800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   4
         TX              =   "Clear Info"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   12632256
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmMain.frx":038A
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton cmdSave 
         Height          =   375
         Left            =   1560
         TabIndex        =   19
         Top             =   7800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   4
         TX              =   "Save Contact"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   12632256
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmMain.frx":03A6
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton cmdDelete 
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   7800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   4
         TX              =   "Delete Contact"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   12632256
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmMain.frx":03C2
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         Height          =   3495
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   4080
         Width           =   4335
      End
      Begin VB.TextBox txtZip 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         MaxLength       =   5
         TabIndex        =   7
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox txtState 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   6
         Top             =   3360
         Width           =   375
      End
      Begin VB.TextBox txtCity 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   1920
         Width           =   4335
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   2640
         Width           =   4335
      End
      Begin VB.TextBox txtPhone 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   4335
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label7 
         Caption         =   "Misc. Notes:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "City / State / Zip Code"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Email Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   3495
      End
      Begin VB.Label Label4 
         Caption         =   "Street Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Phone Number:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Full Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Complete Contact List"
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      Begin MSComctlLib.ListView lvNames 
         Height          =   7695
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   13573
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Contact Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Phone Number"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Click on any name to view all of their information in the frame to the right!"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   7920
         Width           =   6495
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
txtAddress = ""
txtCity = ""
txtEmail = ""
txtInfo = ""
txtName = ""
txtPhone = ""
txtState = ""
txtZip = ""
End Sub

Private Sub cmdDelete_Click()
Dim i As Integer

XML.DeleteNode "/contacts/" & Left(txtName, InStr(txtName, " ") - 1)
XML.Save App.Path & "\contact.xml"
XML.OpenXML App.Path & "\contact.xml"

For i = 1 To lvNames.ListItems.Count
    If lvNames.ListItems(i).Text = Left(txtName, InStr(txtName, " ") - 1) Then
        lvNames.ListItems.Remove i
        GoTo Clear:
    End If
Next i

Clear:
cmdClear_Click
End Sub

Private Sub cmdSave_Click()
SaveContactInfo Left(txtName, InStr(txtName, " ") - 1)
End Sub

Private Sub Form_Load()
lvNames.ColumnHeaders(1).Width = lvNames.Width / 2
lvNames.ColumnHeaders(2).Width = (lvNames.Width / 2) - 45

XML.OpenXML App.Path & "\contact.xml"

LoadContactList
End Sub

Private Sub lvNames_Click()
On Error Resume Next
LoadContactInfo lvNames.SelectedItem.Text
End Sub

Private Sub txtName_Change()
If Right(txtName, 1) <> " " Then txtName = txtName & " "
txtName.SelStart = Len(txtName) - 1
End Sub
