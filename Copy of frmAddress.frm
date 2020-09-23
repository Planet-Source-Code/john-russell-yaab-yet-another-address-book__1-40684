VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAddress 
   Caption         =   "Address Book"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10890
   Icon            =   "frmAddress.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   10890
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1200
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddress.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddress.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddress.frx":0CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddress.frx":1138
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddress.frx":158A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddress.frx":19DC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   7905
      Width           =   10890
      _ExtentX        =   19209
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10980
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "13:10"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "07/11/2002"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvContact 
      Height          =   6495
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   11456
      _Version        =   393217
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
   End
   Begin TabDlg.SSTab tbContact 
      Height          =   6375
      Left            =   3840
      TabIndex        =   1
      Top             =   960
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   11245
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Contact"
      TabPicture(0)   =   "frmAddress.frx":1E2E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblBirthday"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label9"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label11"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label12"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label13"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label7"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label8"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label10"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "mskHomeZip"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "mskHomePhone"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "mskHomeFax"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "mskHomeCellPhone"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "mskBirthday"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtFirstName"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtMiddleInitial"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtLastName"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtHomeStreet"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtHomeCity"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtHomeState"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtHomeEmail"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtCompany"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).ControlCount=   26
      TabCaption(1)   =   "Call Log"
      TabPicture(1)   =   "frmAddress.frx":1E4A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtNotes"
      Tab(1).Control(1)=   "lvCalls"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "DB Statistics"
      TabPicture(2)   =   "frmAddress.frx":1E66
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(1)=   "lblDateCreated"
      Tab(2).Control(2)=   "lblLastUpdated"
      Tab(2).Control(3)=   "lblRecordCount"
      Tab(2).Control(4)=   "Label14"
      Tab(2).Control(5)=   "Label15"
      Tab(2).Control(6)=   "Label16"
      Tab(2).ControlCount=   7
      Begin VB.TextBox txtCompany 
         Height          =   285
         Left            =   240
         TabIndex        =   50
         Tag             =   "1"
         Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         Top             =   1080
         Width           =   5535
      End
      Begin VB.Frame Frame1 
         Caption         =   "Database Stats"
         Height          =   1935
         Left            =   -74760
         TabIndex        =   36
         Top             =   3720
         Width           =   5655
         Begin VB.Label lblDiskReads 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1320
            TabIndex        =   48
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lblDiskWrites 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1320
            TabIndex        =   47
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblReadCache 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1320
            TabIndex        =   46
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblReadAhead 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4200
            TabIndex        =   45
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lblLocksPlaced 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4200
            TabIndex        =   44
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblReleaseLocks 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4200
            TabIndex        =   43
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label17 
            Caption         =   "Disk Reads"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label18 
            Caption         =   "Disk Writes"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label19 
            Caption         =   "Read Cache"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label20 
            Caption         =   "Read Ahead"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2880
            TabIndex        =   39
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label21 
            Caption         =   "Locks Placed"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2880
            TabIndex        =   38
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label22 
            Caption         =   "Locks Released"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2880
            TabIndex        =   37
            Top             =   1440
            Width           =   1215
         End
      End
      Begin VB.TextBox txtNotes 
         Height          =   2295
         Left            =   -74640
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   3720
         Width           =   6495
      End
      Begin MSComctlLib.ListView lvCalls 
         Height          =   2895
         Left            =   -74640
         TabIndex        =   28
         Top             =   600
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   5106
         View            =   2
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox txtHomeEmail 
         Height          =   285
         Left            =   240
         MaxLength       =   50
         TabIndex        =   10
         Tag             =   "1"
         Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         Top             =   4920
         Width           =   4695
      End
      Begin VB.TextBox txtHomeState 
         Height          =   285
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   9
         Tag             =   "1"
         Text            =   "XX"
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox txtHomeCity 
         Height          =   285
         Left            =   2640
         MaxLength       =   12
         TabIndex        =   8
         Tag             =   "1"
         Text            =   "XXXXXXXXXXXX"
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox txtHomeStreet 
         Height          =   285
         Left            =   240
         MaxLength       =   20
         TabIndex        =   7
         Tag             =   "1"
         Text            =   "XXXXXXXXXXXXXXXXXXXX"
         Top             =   3120
         Width           =   2295
      End
      Begin VB.TextBox txtLastName 
         Height          =   285
         Left            =   3600
         MaxLength       =   20
         TabIndex        =   6
         Tag             =   "1"
         Text            =   "XXXXXXXXXXXXXXXXXXXX"
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox txtMiddleInitial 
         Height          =   285
         Left            =   2640
         MaxLength       =   1
         TabIndex        =   5
         Tag             =   "1"
         Text            =   "X"
         Top             =   2040
         Width           =   255
      End
      Begin VB.TextBox txtFirstName 
         Height          =   285
         Left            =   240
         MaxLength       =   15
         TabIndex        =   4
         Tag             =   "1"
         Text            =   "XXXXXXXXXXXXXXX"
         Top             =   2040
         Width           =   1695
      End
      Begin MSMask.MaskEdBox mskBirthday 
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Tag             =   "1"
         Top             =   5520
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskHomeCellPhone 
         Height          =   285
         Left            =   4320
         TabIndex        =   12
         Tag             =   "1"
         Top             =   4080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   13
         Mask            =   "(###)###-####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskHomeFax 
         Height          =   285
         Left            =   2280
         TabIndex        =   13
         Tag             =   "1"
         Top             =   4080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   13
         Mask            =   "(###)###-####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskHomePhone 
         Height          =   285
         Left            =   240
         TabIndex        =   14
         Tag             =   "1"
         Top             =   4080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   13
         Mask            =   "(###)###-####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskHomeZip 
         Height          =   285
         Left            =   4680
         TabIndex        =   15
         Tag             =   "1"
         Top             =   3120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "#####-####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Company"
         Height          =   195
         Left            =   240
         TabIndex        =   49
         Tag             =   "CCom"
         Top             =   720
         Width           =   660
      End
      Begin VB.Label lblDateCreated 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -73080
         TabIndex        =   35
         Top             =   1080
         Width           =   3615
      End
      Begin VB.Label lblLastUpdated 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -73080
         TabIndex        =   34
         Top             =   1680
         Width           =   3615
      End
      Begin VB.Label lblRecordCount 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -73080
         TabIndex        =   33
         Top             =   2280
         Width           =   3615
      End
      Begin VB.Label Label14 
         Caption         =   "Database Created"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   32
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Last Updated"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   31
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "Contact Records"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   30
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Zip"
         Height          =   255
         Left            =   4680
         TabIndex        =   27
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "City"
         Height          =   255
         Left            =   2640
         TabIndex        =   26
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Home Street"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label Label13 
         Caption         =   "Last Name"
         Height          =   255
         Left            =   3600
         TabIndex        =   24
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Label12 
         Caption         =   "M. I."
         Height          =   255
         Left            =   2640
         TabIndex        =   23
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "First Name"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "State"
         Height          =   255
         Left            =   4200
         TabIndex        =   21
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Cell Phone"
         Height          =   255
         Left            =   4320
         TabIndex        =   20
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Home Fax"
         Height          =   255
         Left            =   2280
         TabIndex        =   19
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Home Phone"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Home Email Address"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   4680
         Width           =   3255
      End
      Begin VB.Label lblBirthday 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1920
         TabIndex        =   16
         Tag             =   "1"
         Top             =   5520
         Width           =   3015
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10890
      _ExtentX        =   19209
      _ExtentY        =   1429
      ButtonWidth     =   1323
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancel"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Quit"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuAddNew 
         Caption         =   "&Add New"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
   End
End
Attribute VB_Name = "frmAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim contactNode As Node ' used to populate tree control
Dim rsNotesTable As Recordset
Dim rsCallType As Recordset

Dim iCurrentState As Integer ' current state of program, ie. adding, editing etc...
Dim lCurrentContactKey As Long ' unique ID of contact
Dim sCurrentContactName As String ' contact's name
Dim bFieldsPopulated As Boolean  'flag to see if the fields have data and whether they should be cleared

Private Sub Form_Activate()

    Static bLoadedAlready As Boolean ' False by default
    sbStatus.Panels.Item(2).Text = "Loading...."

    If (Not bLoadedAlready) Then
         Call InitializeForm ' Should be called only once for any session
         bLoadedAlready = True
    
    End If

    sbStatus.Panels.Item(2).Text = "Ready."

End Sub

Private Sub Form_Load()

    If Not openTheDatabase() Then
        MsgBox "Sorry - the database could not be found. Check for CONTACTS.MDB"
        End ' Terminate the program unconditionally
    End If

    Call clearFields
    bFieldsPopulated = False ' As we have just loaded the program, the fields should be empty
    iCurrentState = NOW_IDLE

End Sub


Public Sub InitializeForm()

    Screen.MousePointer = vbHourglass
    iCurrentState = NOW_IDLE
    sbStatus.Panels.Item(2).Text = "Loading...."
    tbContact.Tab = 0 ' make first tab default shown
    DoEvents ' update visual components. Ensures that all visual changes are shown immediately
    Call clearFields
    Call lockFields(True)
    Call updateTree
    Call updateForm
    Call setUpListView
       
    tbContact.Enabled = False
    
    Screen.MousePointer = vbDefault
    sbStatus.Panels.Item(2).Text = "Ready."

End Sub


Public Sub clearFields()

    Dim indx As Integer
    Dim tempMask As String
    
    With Me.Controls
        For indx = 0 To .Count - 1
            If Me.Controls(indx).Tag = "1" Then
                
                If (TypeOf Me.Controls(indx) Is TextBox) Then
                
                    Me.Controls(indx).Text = ""
                    
                ElseIf (TypeOf Me.Controls(indx) Is MaskEdBox) Then
                
                    tempMask = Me.Controls(indx).Mask
                    Me.Controls(indx).Mask = ""
                    Me.Controls(indx).Text = ""
                    Me.Controls(indx).Mask = tempMask
                 Else
                    Me.Controls(indx).Caption = ""
                End If
            
            End If
                  
        Next
     
    End With
    
    DoEvents
    
End Sub

Public Sub lockFields(bDoLock As Boolean)
    Dim indx As Integer

    For indx = 0 To Me.Controls.Count - 1
        If Me.Controls(indx).Tag = "1" Then
            If (TypeOf Me.Controls(indx) Is TextBox) Then
                If (bDoLock = True) Then
                    Me.Controls(indx).Locked = True
                    Me.Controls(indx).BackColor = vbWhite
                Else
                    Me.Controls(indx).Locked = False
                    Me.Controls(indx).BackColor = vbYellow
                End If
            ElseIf (TypeOf Me.Controls(indx) Is MaskEdBox) Then
                If (bDoLock = True) Then
                    Me.Controls(indx).Enabled = False
                    Me.Controls(indx).BackColor = vbWhite
                Else
                    Me.Controls(indx).Enabled = True
                    Me.Controls(indx).BackColor = vbYellow
                End If
            End If
        End If
    Next
DoEvents
End Sub


Public Sub updateTree()

    Dim indx As Integer
    Dim rsAllNames As Recordset
    Dim sqlNames As String
    Dim sContactName As String
    Dim currentAlpha As String

    tvContact.Nodes.Clear ' Clear any nodes in tree

    sqlNames = "SELECT ContactID, LastName, FirstName, MiddleInitial "
    sqlNames = sqlNames & "FROM Contact ORDER BY"
    sqlNames = sqlNames & " LastName, FirstName, MiddleInitial "

    Set rsAllNames = dbContact.OpenRecordset(sqlNames) ' open recordset

    If (rsAllNames.RecordCount > 0) Then ' Are there any contacts? If so, go to first record
        rsAllNames.MoveFirst
    End If

    For indx = Asc("A") To Asc("Z")
        currentAlpha = Chr(indx)
    
        ' Add the chatacter to the treeview control.
        ' So we add a  node to the treeviews nodes collection.
        ' currentAlpa is used to represwent the unique key (A-Z)  to identify the node
        ' and the text that will be whown in the control
        
        Set contactNode = tvContact.Nodes.Add _
            (, , currentAlpha, currentAlpha)
  
        If (Not rsAllNames.EOF) Then
            Do While UCase$(Left(rsAllNames!LastName, 1)) = currentAlpha
                With rsAllNames
                    sContactName = !LastName & ", "
                    sContactName = sContactName & !FirstName
                    If (Not IsNull(!MiddleInitial)) Then
                        sContactName = sContactName & " " & !MiddleInitial & "."
                    End If
                End With

                DoEvents
                
                ' Add the contact under the letter (A-Z) in treeview control
                ' as a 'child' node of the (A-Z) node
                ' NB.  for some reason, VB does not like strict umerics converted to a string
                ' So we concatenate the string "ID" with the contactID
                Set contactNode = tvContact.Nodes.Add(currentAlpha, _
                tvwChild, "ID" & CStr(rsAllNames!ContactID), sContactName)
                rsAllNames.MoveNext
                If (rsAllNames.EOF) Then
                    Exit Do
                End If
            Loop
        End If
    Next

    sbStatus.Panels.Item(1).Text = "There are " & _
    rsAllNames.RecordCount & " contacts in the database."

    rsAllNames.Close

    DoEvents

End Sub

Public Sub updateForm()

    ' This approach isolates all the messy details of setting up buttons and controls
    ' into a single routine. Once working, you can forget about it.
    
    Select Case iCurrentState
        Case NOW_ADDING, NOW_EDITING
            If (iCurrentState = NOW_ADDING) Then
                sbStatus.Panels.Item(2).Text = "Adding..."
                Call clearFields
            Else
                sbStatus.Panels.Item(2).Text = "Editing..."
            End If
            tbContact.Enabled = True
            tbContact.Tab = 0              '-- make the 1st tab current
            tbContact.TabEnabled(1) = False '-disable the 2nd and 3rd tabs
            tbContact.TabEnabled(2) = False
            tvContact.Enabled = False
            lockFields (False)        '-- unlock fields and set background
            txtCompany.SetFocus    '-- set focus to first name field
            Toolbar1.Buttons(bAdd).Enabled = False
            Toolbar1.Buttons(bCancel).Enabled = True
            Toolbar1.Buttons(bSave).Enabled = True
            Toolbar1.Buttons(bDelete).Enabled = False
            Toolbar1.Buttons(bEdit).Enabled = False
            Toolbar1.Buttons(bQuit).Enabled = False
        Case NOW_IDLE
            sbStatus.Panels.Item(2).Text = "Ready."
            Toolbar1.Buttons(bAdd).Enabled = True
            Toolbar1.Buttons(bCancel).Enabled = False
            Toolbar1.Buttons(bSave).Enabled = False
            Toolbar1.Buttons(bQuit).Enabled = True
            If (Len(txtLastName)) Then
                Toolbar1.Buttons(bDelete).Enabled = True
                Toolbar1.Buttons(bEdit).Enabled = True
            Else
                Toolbar1.Buttons(bDelete).Enabled = False
                Toolbar1.Buttons(bEdit).Enabled = False
            End If
            tvContact.Enabled = True
            tbContact.TabEnabled(1) = True
            tbContact.TabEnabled(2) = True
        Case NOW_DELETING
            sbStatus.Panels.Item(2).Text = "Deleting...."
            Toolbar1.Buttons(bAdd).Enabled = False
            Toolbar1.Buttons(bCancel).Enabled = False
            Toolbar1.Buttons(bSave).Enabled = False
            Toolbar1.Buttons(bDelete).Enabled = False
            Toolbar1.Buttons(bEdit).Enabled = False
            Toolbar1.Buttons(bQuit).Enabled = False
        Case NOW_SAVING
            sbStatus.Panels.Item(2).Text = "Saving...."
            tvContact.Enabled = True
            Toolbar1.Buttons(bAdd).Enabled = False
            Toolbar1.Buttons(bCancel).Enabled = False
            Toolbar1.Buttons(bSave).Enabled = False
            Toolbar1.Buttons(bDelete).Enabled = False
            Toolbar1.Buttons(bEdit).Enabled = False
            Toolbar1.Buttons(bQuit).Enabled = False
            If (Len(mskBirthday)) Then
                lblBirthday = Format$(mskBirthday, "mmmm dd, yyyy")
            End If
    End Select

    DoEvents

End Sub


Public Sub setUpListView()

    ' Here we are just adding columns. Data - in the form of ListItems will be added later
    ' We are passing just 2  of the 5 possible parameters into the Add method (text and width)
    ' Width of each control is divided by 3 so each column takes up a third of the screen space
    ' Finally, list view control shows itself in report format
    
    
    Dim clmHdr As ColumnHeader
    Static bBeenHereBefore As Boolean ' as this routine is triggered through the InitializeForm event
    ' each time a new contact is entered, prevent this from happening
    
    If bBeenHereBefore = False Then
    
        Set clmHdr = lvCalls.ColumnHeaders. _
             Add(, , "Date / Time", lvCalls.Width \ 3)

        Set clmHdr = lvCalls.ColumnHeaders. _
             Add(, , "Type of Call", lvCalls.Width \ 3)
             
        Set clmHdr = lvCalls.ColumnHeaders. _
             Add(, , "Call Identifier", lvCalls.Width \ 3)
        bBeenHereBefore = True
    End If
    lvCalls.View = lvwReport

End Sub


Private Sub lvCalls_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ' When the user clicks on a column, such ad date of call, we will sort the items.
    Dim nSortCol As Integer
        
    ' When a ColumnHeader object is clicked, the list view
    ' control is sorted by the SubItems of that column.
    ' Set the SortKey to the index of the ColumnHeader - 1

    nSortCol = ColumnHeader.Index - 1
    
    If (lvCalls.SortKey = nSortCol) Then
        lvCalls.SortOrder = 1 - lvCalls.SortOrder
    Else
        lvCalls.SortKey = nSortCol
        lvCalls.SortOrder = lvwAscending
    End If
    
    '-- Do the sort now
    lvCalls.Sorted = True
End Sub

Private Sub lvCalls_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Bring up popup menu. Check button pressed when the list view is clicked.
    ' If there are no calls, there's nothing to delete so disable mnuDelete menu option.
    If Button = vbRightButton Then
    
        If (rsCallType.RecordCount < 1) Then
            mnuDelete.Enabled = False
        Else
            mnuDelete.Enabled = True
        End If
        'Display popup menu using VB command PopupMenu and passing it name of the menu
        
        PopupMenu mnuPopup
         
    End If
    
End Sub

Private Sub mnuAddNew_Click()
    frmCall.sContactName = sCurrentContactName
    frmCall.lContactNumber = lCurrentContactKey
    frmCall.Show vbModal
    Call populateListView
End Sub

Private Sub mnuDelete_Click()
    Dim indx As Integer
    Dim rsDeleteCall As Recordset
    Dim sDeleteCall As String

    indx = MsgBox("Are you sure you wish to delete this call from " & _
              lvCalls.ListItems(lvCalls.SelectedItem.Index) & "?", _
              vbYesNo + vbQuestion, progname)


    If (indx <> vbYes) Then Exit Sub

        sDeleteCall = "DELETE * FROM Notes WHERE CallCounter = " & _
              lvCalls.ListItems(lvCalls.SelectedItem.Index).SubItems(2)

        dbContact.Execute (sDeleteCall)
    Call populateListView


End Sub

Private Sub tbContact_DblClick()
    If (tbContact.Tab = 2) Then
        lblDateCreated = Format$(rsContactTable.DateCreated, _
        "dddd mmmm dd, yyyy hh:mm AMPM")
        lblLastUpdated = Format$(rsContactTable.LastUpdated, _
        "dddd mmmm dd, yyyy hh:mm AMPM")
        lblRecordCount = "Contacts in Database: " & _
        rsContactTable.RecordCount
        lblDiskReads = ISAMStats(0)
        lblDiskWrites = ISAMStats(1)
        lblReadCache = ISAMStats(2)
        lblReadAhead = ISAMStats(3)
        lblLocksPlaced = ISAMStats(4)
        lblReleaseLocks = ISAMStats(5)
        End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    
        Case bAdd
            iCurrentState = NOW_ADDING
            Call updateForm
            
        Case bCancel
            
            If (bFieldsPopulated = True) Then
                Call populateFields
            End If
            Call lockFields(True)
            iCurrentState = NOW_IDLE
            Call updateForm
            
        Case bSave
        
            '-- Here we are saving either a new or edited entry --
            If (Not validateEntry()) Then
                Exit Sub
            End If
            postContact
        
        Case bDelete
        
            Dim indx As Integer
            Dim sMsg As String
            Dim sDeleteSQL As String
            sMsg = "Delete " & tvContact.SelectedItem & _
            " and all related call logs?"
            indx = MsgBox(sMsg, vbYesNo + vbCritical, progname)
            If (indx <> vbYes) Then Exit Sub
            sDeleteSQL = "DELETE * FROM Contact WHERE ContactID = " _
            & lCurrentContactKey
            ' Cascade delete should take care of any reelaated calls
            dbContact.Execute (sDeleteSQL)

    Call InitializeForm

        Case bEdit
        
            iCurrentState = NOW_EDITING
            updateForm
        
        Case bQuit
        
            rsContactTable.Close
            dbContact.Close
            Set rsContactTable = Nothing
            Set dbContact = Nothing
            Unload Me
            
    End Select
    
End Sub

Private Sub tvContact_NodeClick(ByVal Node As MSComctlLib.Node)

    ' If user clicks on a letter such as 'A' instead of a name, we dont want to take any action.
    ' To determine when this occurs, we just check the length of the Key property of the node
    ' that was clicked. If it's only one character long, we exit the routine.
    ' Otherwise, if user clicked on a contact name, we get the ContactID from the Key property
    ' of the node.
    
    If (Len(Node.Key) = 1) Then Exit Sub

    '-- Here we retrieve the contact the user clicked on --
    lCurrentContactKey = CLng(Mid$(Node.Key, 3, Len(Node.Key)))
    With rsContactTable
        .Index = "PrimaryKey"
        .Seek "=", lCurrentContactKey
        If Not .NoMatch Then
            bFieldsPopulated = True
            sCurrentContactName = tvContact.SelectedItem
            Call populateFields
            Call populateListView
            tbContact.Enabled = True
        Else
            MsgBox ("Not found! That's odd because it should be there!!?!")
            
    End If
End With
End Sub


Public Sub populateFields()
    Dim sBirthDay As String

    ' Now that we have a valid record in rsContactTable, as identified in the cCurrentContactName
    ' information, let's display the fields on the form.
    '-- Here we retrieve the fields from the database and --
    '-- populate the fields in the user interface.        --

    Call clearFields

    ' Update each field on the tabcontrol with the appropriate fields from the current record

    With rsContactTable
        If (Not IsNull(!Company)) Then txtCompany = !Company
        If (Not IsNull(!LastName)) Then txtLastName = !LastName
        If (Not IsNull(!MiddleInitial)) Then
            txtMiddleInitial = !MiddleInitial
        End If
        If (Not IsNull(!FirstName)) Then txtFirstName = !FirstName
        If (Not IsNull(!HomeStreet)) Then
            txtHomeStreet = !HomeStreet
        End If
        If (Not IsNull(!HomeCity)) Then
            txtHomeCity = !HomeCity
        End If
        If (Not IsNull(!HomeState)) Then
            txtHomeState = !HomeState
        End If
        If (Not IsNull(!HomeZip)) Then
            mskHomeZip = !HomeZip
        End If
        If (Not IsNull(!HomePhone)) Then
            mskHomePhone = !HomePhone
        End If
        If (Not IsNull(!HomeFax)) Then
            mskHomeFax = !HomeFax
        End If
        If (Not IsNull(!HomeEmail)) Then
            txtHomeEmail = !HomeEmail
        End If
        If (Not IsNull(!HomeCellPhone)) Then
            mskHomeCellPhone = !HomeCellPhone
        End If
        If (Not IsNull(!Birthday)) Then
            sBirthDay = !Birthday
            convertDate sBirthDay
            mskBirthday = sBirthDay
            lblBirthday = Format$(!Birthday, "dddd dd mmmm, yyyy")
        End If
        DoEvents
        
        ' Update all of the form buttons
        Call updateForm

End With


End Sub

Public Sub populateListView()
    ' Once the fields for the contact are displayed, we want to see if there are any
    ' calls logged for that contact.
    
Dim itemToAdd As ListItem
Dim noteSQL As String

' Clear any calls in the list view  from a previous contact. TxtNotes textbox holds text of any previous
' call is cleared also. Lock the control so user can't accidentally overwrite data

lvCalls.ListItems.Clear
txtNotes = ""
txtNotes.Locked = True

' Construct SQL string to retieve call records for user in descending order so that latest call
' is first.

noteSQL = "SELECT DISTINCTROW Notes.DateOfCall,"
noteSQL = noteSQL & "Notes.CallTypeID, Notes.NotesOnPhoneCall, "
noteSQL = noteSQL & " Notes.CallCounter, CallType.CallDescription,"
noteSQL = noteSQL & " Notes.ContactID "
noteSQL = noteSQL & " FROM Notes "
noteSQL = noteSQL & " INNER JOIN CallType ON Notes.CallTypeID ="
noteSQL = noteSQL & " CallType.CallTypeID "
noteSQL = noteSQL & " WHERE Notes.ContactID = " & _
    lCurrentContactKey
noteSQL = noteSQL & " ORDER BY Notes.DateOfCall DESC"

Set rsCallType = dbContact.OpenRecordset(noteSQL)

If (rsCallType.RecordCount > 0) Then
   rsCallType.MoveFirst
    While Not rsCallType.EOF
       Set itemToAdd = lvCalls.ListItems.Add(, , _
          Format$(rsCallType!DateOfCall, "dddd mmmm dd, yyyy"))
       itemToAdd.SubItems(1) = rsCallType!CallDescription
       itemToAdd.SubItems(2) = CStr(rsCallType!CallCounter)
       rsCallType.MoveNext
   Wend
   sbStatus.Panels.Item(1).Text = "There are " & _
    rsCallType.RecordCount & " calls logged for " & _
    sCurrentContactName
Else
   Set itemToAdd = lvCalls.ListItems.Add(, , "No calls logged")
   sbStatus.Panels.Item(1).Text = "No calls logged for " _
    & sCurrentContactName
End If

lvCalls.SelectedItem = lvCalls.ListItems(1)
Call lvCalls_ItemClick(lvCalls.SelectedItem)
DoEvents
End Sub

Public Sub convertDate(sBirthDay As String)
    
    Dim sYear As String
    ' First, check lenght of sBirthday. A correctly formatted date should be 10 characters:
    ' 2 = day, 2 = month, 4 = year and 2 = '/' separators
    
    Select Case Len(sBirthDay)

    Case 10 'needed to keep centuries correct prior to 1900 and after 2029.
        Exit Sub

    Case 9
        If Mid$(sBirthDay, 2, 1) = "/" Then
            sBirthDay = "0" & sBirthDay
        Else
            sBirthDay = Left(sBirthDay, 3) & "0" & Mid$(sBirthDay, 4, 6)
        End If
        Exit Sub

    Case 8
        Select Case Mid$(sBirthDay, 2, 1)
            Case "/"
            sBirthDay = "0" & Left(sBirthDay, 2) & "0" & Right(sBirthDay, 6)
        Exit Sub
        Case Else
            End Select

    Case 7
        Select Case Mid$(sBirthDay, 2, 1)
            Case "/"
                sBirthDay = "0" & Left(sBirthDay, 7)
            Case Else
                sBirthDay = Left(sBirthDay, 3) & "0" & Right(sBirthDay, 4)
            End Select

    Case 6
        Select Case Mid$(sBirthDay, 2, 1)
            Case Is = "/"
                sBirthDay = "0" & Left(sBirthDay, 2) & "0" & Right(sBirthDay, 4)
            Case Else
        End Select

    Case Else
    End Select

    sYear = Right(sBirthDay, 2)

    If sYear >= 30 Then
        sBirthDay = Mid$(sBirthDay, 1, 6) & "19" & sYear
    Else
        sBirthDay = Mid$(sBirthDay, 1, 6) & "20" & sYear
    End If
    
End Sub


Private Sub lvCalls_ItemClick(ByVal Item As MSComctlLib.ListItem)
If (rsCallType.RecordCount > 0) Then
    rsCallType.MoveFirst
    '-- Find the record that has the ID --
    rsCallType.FindFirst "CallCounter = " & _
                 lvCalls.ListItems(Item.Index).SubItems(2)
     txtNotes = rsCallType!NotesOnPhoneCall
End If

End Sub



Public Function validateEntry() As Boolean

    ' When we wish to save a new contact, or a current record just edited, before saving
    ' any data, ensure that there is at least a name for the contact
    ' Perform 3 tests:
    ' Both first and last name must be entered
    ' make sure date is valid in dd/mm/yyyy format
    
    Dim indx As Integer

    validateEntry = True
    
    sbStatus.Panels.Item(2).Text = "Validating..."
     If (Len(txtCompany) < 1) Then
        tbContact.Tab = 0
        indx = MsgBox("Please enter the company name.", _
          vbInformation + vbOKOnly, progname)
        txtCompany.SetFocus
        validateEntry = False
        Exit Function
    End If
    
    If (Len(txtFirstName) < 1) Then
        tbContact.Tab = 0
        indx = MsgBox("Please enter the first name of the contact.", _
          vbInformation + vbOKOnly, progname)
        txtFirstName.SetFocus
        validateEntry = False
        Exit Function
    End If

    If (Len(txtLastName) < 1) Then
        tbContact.Tab = 0
        indx = MsgBox("Please enter the last name of the contact.", _
          vbInformation + vbOKOnly, progname)
        txtLastName.SetFocus
        validateEntry = False
        Exit Function
    End If

    mskBirthday.PromptInclude = False
    If (Len(mskBirthday.Text) > 0) Then
        mskBirthday.PromptInclude = True
        If (Not IsDate(mskBirthday)) Then
            tbContact.Tab = 0
            indx = MsgBox("Please enter a valid birthdate dd/mm/yyyy.", _
            vbInformation + vbOKOnly, progname)
            mskBirthday.SetFocus
            validateEntry = False
            Exit Function
        End If
    End If
    mskBirthday.PromptInclude = False

End Function

Public Sub postContact()
    ' When user wants to save new or edited record, the database is updated with any new information
    Dim rsMaxIDNumber As Recordset
    Dim sqlMaxID As String
    Dim lNewContactID As Long

    Screen.MousePointer = vbHourglass
    sbStatus.Panels.Item(2).Text = "Posting Contact...."

    If (iCurrentState = NOW_ADDING) Then
        rsContactTable.AddNew
    Else
        With rsContactTable
            .MoveFirst
            .Index = "PrimaryKey"
            .Seek "=", lCurrentContactKey
            If Not .NoMatch Then
                rsContactTable.Edit
            Else
                MsgBox ("Ohhhh Nooo")
            End If
        End With
    End If

    With rsContactTable
        If (Len(txtCompany)) Then !Company = txtCompany
        If (Len(txtFirstName)) Then !FirstName = txtFirstName
        If (Len(txtMiddleInitial)) Then !MiddleInitial = _
            txtMiddleInitial
        If (Len(txtLastName)) Then !LastName = txtLastName
        If (Len(txtHomeStreet)) Then !HomeStreet = txtHomeStreet
        If (Len(txtHomeCity)) Then !HomeCity = txtHomeCity
        If (Len(txtHomeState)) Then !HomeState = txtHomeState
        If (Len(mskHomeZip)) Then !HomeZip = mskHomeZip
        If (Len(mskHomePhone)) Then !HomePhone = mskHomePhone
        If (Len(mskHomeFax)) Then !HomeFax = mskHomeFax
        If (Len(mskHomeCellPhone)) Then !HomeCellPhone = _
            mskHomeCellPhone
        If (Len(txtHomeEmail)) Then !HomeEmail = txtHomeEmail
            mskBirthday.PromptInclude = False
        If (Len(mskBirthday.Text) > 0) Then
            mskBirthday.PromptInclude = True
            !Birthday = mskBirthday
            lblBirthday = Format$(!Birthday, "dddd mmmm dd, yyyy")
        End If
        mskBirthday.PromptInclude = True
        .Update

    End With

    DoEvents

    If (iCurrentState = NOW_ADDING) Then
        ' Force the tree view to be refreshed with the new contact and set up the form
        Call InitializeForm
    Else
        iCurrentState = NOW_IDLE
        Call lockFields(True)
        Call updateForm
        End If

    sbStatus.Panels.Item(2).Text = "Ready."
    Screen.MousePointer = vbDefault

End Sub



