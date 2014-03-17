VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{3D800911-77E3-43DE-82EA-7FC87C713180}#1.1#0"; "cPopMenu6.ocx"
Begin VB.Form Form1 
   Caption         =   "Demo menambahkan icon pada menu standar VB"
   ClientHeight    =   4755
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin cPopMenu6.PopMenu PopMenu1 
      Left            =   3480
      Top             =   1680
      _ExtentX        =   1058
      _ExtentY        =   1058
      HighlightCheckedItems=   0   'False
      TickIconIndex   =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1680
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":059A
            Key             =   "open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0B34
            Key             =   "close"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":10CE
            Key             =   "save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1668
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1C02
            Key             =   "print"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":219C
            Key             =   "mail"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2736
            Key             =   "fax"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2CD0
            Key             =   "powerpoint"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
      Begin VB.Menu mnuSpr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As..."
      End
      Begin VB.Menu mnuSpr2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintPreview 
         Caption         =   "Print Preview"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuSpr3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSendTo 
         Caption         =   "Send To"
         Begin VB.Menu mnuMailRecipient 
            Caption         =   "Mail Recipient"
         End
         Begin VB.Menu mnuMailRecipientReview 
            Caption         =   "Mail Recipient (for Review)"
         End
         Begin VB.Menu mnuOnlineMeetingParticipant 
            Caption         =   "Online Meeting Participant"
         End
         Begin VB.Menu mnuFaxRecipient 
            Caption         =   "Fax Recipient..."
         End
         Begin VB.Menu mnuSpr4 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMicrosoftPowerPoint 
            Caption         =   "Microsoft PowerPoint"
         End
      End
      Begin VB.Menu mnuSpr5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
' MMMM  MMMMM  OMMM   MMMO    OMMM    OMMM    OMMMMO     OMMMMO    OMMMMO  '
'  MM    MM   MM MM    MMMO  OMMM    MM MM    MM   MO   OM    MO  OM    MO '
'  MM  MM    MM  MM    MM  OO  MM   MM  MM    MM   MO   OM    MO       OMO '
'  MMMM     MMMMMMMM   MM  MM  MM  MMMMMMMM   MMMMMO     OMMMMO      OMO   '
'  MM  MM        MM    MM      MM       MM    MM   MO   OM    MO   OMO     '
'  MM    MM      MM    MM      MM       MM    MM    MO  OM    MO  OM   MM  '
' MMMM  MMMM    MMMM  MMMM    MMMM     MMMM  MMMM  MMMM  OMMMMO   MMMMMMM  '
'                                                                          '
' K4m4r82's Laboratory                                                     '
' http://coding4ever.wordpress.com
'***************************************************************************

Option Explicit

Private Function getIconIndex(ByVal key As String) As Long
    getIconIndex = ImageList1.ListImages.Item(key).Index - 1
End Function

Private Sub setIcon(ByVal key As String, ByVal menuName As String)
    Dim iconIndex As Long
    
    iconIndex = getIconIndex(key)
    PopMenu1.ItemIcon(menuName) = iconIndex
End Sub

Private Sub Form_Load()
    With PopMenu1
        .ImageList = ImageList1
        .OfficeXpStyle = True
        .SubClassMenu Me
        
        Call setIcon("new", "mnuFile")
        Call setIcon("open", "mnuOpen")
        Call setIcon("close", "mnuClose")
        Call setIcon("save", "mnuSave")
        Call setIcon("preview", "mnuPrintPreview")
        Call setIcon("print", "mnuPrint")
        Call setIcon("mail", "mnuMailRecipient")
        Call setIcon("fax", "mnuFaxRecipient")
        Call setIcon("powerpoint", "mnuMicrosoftPowerPoint")
     End With
End Sub
