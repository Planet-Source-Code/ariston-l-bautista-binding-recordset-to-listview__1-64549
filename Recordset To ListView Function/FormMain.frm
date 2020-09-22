VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recordset To Listview"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc 
      Height          =   375
      Left            =   240
      Top             =   4200
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3360
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMain.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView 
      Height          =   3855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   6800
      Arrange         =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''***************************************************************************
'' File Name: Recordset To Listview
'' Purpose: Automatically Bind Any Recordset To Listview
'' Required Files:  1. Windows Common Control 6.0 (SP6) For The Listview Control
''                  2. Adodc Control
''
'' Programmer: Ariston L. Bautista   E-mail: aristonbautista@gmail.com
'' Date Created: March 06, 2006
'' Last Modified: March 06, 2006
'' Credits: NONE, ALL CODES ARE CODED BY Ariston L. Bautista
''          Give Me Credits When Using This Code
''**************************************************************************


Private Sub Form_Load()
DatabaseConnect Adodc                    'Conenction To Database
RecordsetToListview Adodc, ListView     'Binding Recordset To Database
End Sub

Function DatabaseConnect(AdoControl As Adodc)
DatabasePath = App.Path & "\Database.mdb"
AdoControl.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source= " & DatabasePath & ""
AdoControl.RecordSource = "Select * From Login"
AdoControl.Refresh
End Function



Function RecordsetToListview(AdodcControl As Adodc, ListViewControl As ListView)
Dim Column_Header As ColumnHeader   'Handler Of Column Header Name
Dim List_Item As ListItem           'Handler Of Sub List Items
Dim FieldCount As Integer           'Handler Of No. Of Fields Of Recordset
Dim FieldLoopNo As Integer          'Handler Of How Many Fields To Loop
Dim FieldStringLenght As Integer    'Handler Of Each Field Name String Lenght

'List View Setting, You Can Edit This According To Your Settings
ListViewControl.View = lvwReport
ListViewControl.GridLines = True
ListViewControl.Font.Name = "Tahoma"
ListViewControl.Font.Size = 10
ListViewControl.FullRowSelect = True

'Check How Many Fields Are On The Table
FieldCount = AdodcControl.Recordset.Fields.Count - 1

'This Loop Will Add The Table Field Name To ListView Column Headers
For FieldLoopNo = 0 To FieldCount
    'This Will Get The String Lenght Of The Field Name So That I Can Automatically Adjust My Column Header Width
    FieldStringLenght = Len(AdodcControl.Recordset.Fields(FieldLoopNo).Name) + Int(10)
    Set Column_Header = ListViewControl.ColumnHeaders.Add(, , AdodcControl.Recordset.Fields(FieldLoopNo).Name, TextWidth(String(FieldStringLenght, "0")))
Next FieldLoopNo

'This Loop Will Add Data In The ListView 1st Column
While Not AdodcControl.Recordset.EOF
    Set List_Item = ListViewControl.ListItems.Add(, , AdodcControl.Recordset(0))
    
    'This Set The Icons, You Can Change This If You Like
    List_Item.Icon = 1
    List_Item.SmallIcon = 1

    'This Loop Will Add Data In The ListView 2nd Column Up To The Last Column
    For FieldLoopNo = 1 To FieldCount
    List_Item.SubItems(FieldLoopNo) = AdodcControl.Recordset(FieldLoopNo)
    Next FieldLoopNo
    
    AdodcControl.Recordset.MoveNext
    Wend
End Function






















