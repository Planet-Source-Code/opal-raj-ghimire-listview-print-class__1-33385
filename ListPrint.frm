VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List Print Demo"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10455
   FillColor       =   &H00FFFF00&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   10455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "Format"
      Height          =   375
      Left            =   5850
      TabIndex        =   12
      Top             =   4590
      Width           =   1140
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Preview"
      Height          =   375
      Left            =   5850
      TabIndex        =   11
      Top             =   2970
      Width           =   1140
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print"
      Height          =   375
      Left            =   7200
      TabIndex        =   10
      Top             =   2970
      Width           =   1140
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Print"
      Height          =   375
      Left            =   7200
      TabIndex        =   9
      Top             =   4005
      Width           =   1140
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Preview"
      Height          =   375
      Left            =   5850
      TabIndex        =   8
      Top             =   4050
      Width           =   1140
   End
   Begin VB.CheckBox ChkPic 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Picture"
      Height          =   240
      Left            =   8685
      TabIndex        =   7
      Top             =   4140
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   375
      Left            =   7200
      TabIndex        =   6
      Top             =   3487
      Width           =   1140
   End
   Begin VB.CheckBox ChkBor 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Border"
      Height          =   240
      Left            =   8685
      TabIndex        =   4
      Top             =   3750
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox ChkVer 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Vertical Lines"
      Height          =   240
      Left            =   8685
      TabIndex        =   3
      Top             =   3360
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox ChkHor 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Horizonal Lines"
      Height          =   240
      Left            =   8685
      TabIndex        =   2
      Top             =   2970
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Preview"
      Height          =   375
      Left            =   5850
      TabIndex        =   1
      Top             =   3510
      Width           =   1140
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4770
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListPrint.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListPrint.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListPrint.frx":0A88
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListPrint.frx":0FCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListPrint.frx":1510
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListPrint.frx":1A54
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListPrint.frx":1F98
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListPrint.frx":24DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListPrint.frx":2A20
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListPrint.frx":2F64
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListPrint.frx":34A8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LstView 
      Height          =   2445
      Left            =   5850
      TabIndex        =   0
      Top             =   270
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   4313
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      GridLines       =   -1  'True
      PictureAlignment=   1
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000004&
      BorderStyle     =   3  'Dot
      X1              =   5265
      X2              =   9900
      Y1              =   5175
      Y2              =   5175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   690
      Left            =   5895
      TabIndex        =   5
      Top             =   5355
      Width           =   3930
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000000&
      BorderStyle     =   3  'Dot
      X1              =   5490
      X2              =   5490
      Y1              =   180
      Y2              =   5850
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim M As ListViewPrinter


Private Sub ChkBor_Click()
M.DrawBorder = IIf(ChkBor.Value, True, False)
End Sub

Private Sub ChkHor_Click()
M.DrawHorizontalLines = IIf(ChkHor.Value, True, False)
End Sub

Private Sub ChkPic_Click()
M.HasPicture = IIf(ChkPic.Value, True, False)
End Sub

Private Sub ChkVer_Click()
M.DrawVerticalLines = IIf(ChkVer.Value, True, False)
End Sub


Private Sub Command1_Click()
Dim Page As Integer
Dim sngTotalPage As Single
M.NumOfRowsPerPage = 10
 
sngTotalPage = LstView.ListItems.Count / M.NumOfRowsPerPage
If sngTotalPage - Int(sngTotalPage) <> 0 Then sngTotalPage = Int(sngTotalPage) + 1

Form1.ScaleMode = vbPixels 'this must be done, the container [form1 in this case] must be in vbpixels scalemode
Printer.ScaleMode = vbTwips
Printer.PaperSize = vbPRPSA5
Printer.Orientation = vbPRORPortrait
Printer.Font = LstView.Font.Name
Printer.FontSize = LstView.Font.Size

While Not M.LastRowPrinted
        Page = Page + 1
        M.SetRows
        Printer.CurrentX = 700
        Printer.CurrentY = 900: Printer.FontSize = 18: Printer.FontName = "Times New Roman"
        Printer.Print "My Report"
        Printer.FontSize = 8: Printer.FontName = "MS SANS SERIF"
        
        M.PrintHead Printer
        M.PrintBody Printer
        
        Printer.CurrentY = 4400
        Printer.CurrentX = 700
        Printer.Print "Date: " + Str(Date)
        Printer.CurrentX = 700
        Printer.Print "Time: " + Str(Time)
        Printer.CurrentX = 700
        Printer.Print "Page: " + Str(Page) + " of " + Str(sngTotalPage)
        Printer.NewPage
Wend
       Printer.EndDoc
M.LastRowPrinted = False
Form1.ScaleMode = vbTwips
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label1 = "Paper size A5 (Half of A4)  Orientation Portait  Paper [4 Pages]"
End Sub

Private Sub Command2_Click()
Enable False
Cls
Form1.Font = LstView.Font.Name
Form1.FontSize = LstView.Font.Size
With M
           .RowFrom = 1
           .RowTo = 15
           .PrintHead Form1
           .PrintBody Form1

End With
Enable True
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label1 = "Prints the given range of rows (1 to 15)   "
End Sub

Private Sub Command3_Click()
Form1.ScaleMode = vbPixels 'this must be done, the container [form1 in this case] must be in vbpixels scalemode
Printer.ScaleMode = vbTwips
Printer.PaperSize = vbPRPSA5
Printer.Orientation = vbPRORPortrait
Printer.Font = LstView.Font.Name
Printer.FontSize = LstView.Font.Size

With M
               
        .RowFrom = 1
        .RowTo = 15
        .PrintHead Printer
        .PrintBody Printer
        Printer.EndDoc

End With
Form1.ScaleMode = vbTwips
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label1 = "Paper size A5 (Half of A4)  Orientation Portait. [1 Page] "
End Sub

Private Sub Command4_Click()
Dim Page As Integer
Dim sngTotalPage As Single
Enable False
Form1.Font = LstView.Font.Name
Form1.FontSize = LstView.Font.Size


M.NumOfRowsPerPage = 10
 
sngTotalPage = LstView.ListItems.Count / M.NumOfRowsPerPage
If sngTotalPage - Int(sngTotalPage) <> 0 Then sngTotalPage = Int(sngTotalPage) + 1

While Not M.LastRowPrinted
        Form1.Cls
        Page = Page + 1
        M.SetRows
        Form1.CurrentX = 700
        Form1.CurrentY = 900: FontSize = 18: FontName = "Times New Roman": ForeColor = vbRed
        Print "My Report": ForeColor = vbBlack
        Form1.FontSize = 8: Form1.FontName = "MS SANS SERIF"
        
        M.PrintHead Form1
        M.PrintBody Form1
        
        Form1.CurrentY = 4400
        Form1.CurrentX = 700
        Form1.Print "Date: " + Str(Date)
        Form1.CurrentX = 700
        Form1.Print "Time: " + Str(Time)
        Form1.CurrentX = 700
        Form1.Print "Page: " + Str(Page) + " of " + Str(sngTotalPage)

        Sleep 1500
Wend

M.LastRowPrinted = False
Enable True
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label1 = "Prints several pages with the header and footer on each page "
End Sub

Private Sub Command5_Click()
        Dim Page As Integer
        Enable False
        Form1.Font = LstView.Font.Name
        Form1.FontSize = LstView.Font.Size

        Cls 'Preparing First page
        Page = Page + 1
        CurrentX = 700
        CurrentY = 900: FontSize = 18: FontName = "Times New Roman": ForeColor = vbRed
        Print "Weekly Report": ForeColor = vbBlack
        FontSize = 8: FontName = "MS SANS SERIF"
        CurrentX = 700: Print "M/s ";: FontBold = True: Print "Abcdefg, A Literary Journal ": FontBold = False
        CurrentX = 700: Print "Mitrapark, Kathmandu, Nepal"
        CurrentX = 700:      CurrentY = 2050
        Print "Report # 9681": CurrentY = 2050
        CurrentX = 3000: Print "Date:" + Str(Now)
        CurrentX = 700: CurrentY = 2250
        Print "Code #. 12441": CurrentY = 2250
        CurrentX = 3000: Print "A/c #: 223-55-889-420"
        CurrentX = 700:      CurrentY = 2700
        FontBold = True: FontSize = 9: ForeColor = &H808080
        Print "Detail of  CREDIT TRANSACTION"
        FontBold = False: FontSize = 8: ForeColor = vbBlack
        '======End of 1st page details=========
      
        M.PosY = 2950
        M.NumOfRowsPerPage = 8 ' 8 rows on the 1st page
        M.SetRows
        M.PrintHead Form1
        M.PrintBody Form1
        CurrentY = 5400
        CurrentX = 700
        Print "Page: " + Str(Page)
        Sleep 4000
        M.NumOfRowsPerPage = 16 '16 rows for the remaining pages
        M.PosY = 900
        
        '===========Printing remaining Pages=======
        While Not M.LastRowPrinted
        Cls
                    M.SetRows
                    M.PrintHead Form1
                    M.PrintBody Form1
                    
                    Page = Page + 1
                    If Not M.LastRowPrinted Then
                                CurrentY = 5400:        CurrentX = 700
                                Print "Page: " + Str(Page)
                                Sleep 3000
                    End If
                   
        Wend
If M.LastRowPrinted Then
                'Planning to print 4 lines at the end of report(Olny if there is space)
             
                If (5400 - M.CurrentY) > (4 * Form1.TextHeight("X")) Then
             
                            CurrentY = M.CurrentY
                            CurrentX = 700: Print ""
                            CurrentX = 700: Print "End of the detailed transcation list "
                            CurrentX = 700: Print "You can view this report on our "
                            CurrentX = 700: Print "web page  www.www.www.";: FontBold = True: Print " Thank you !  ": FontBold = False
                            CurrentY = 5400:        CurrentX = 700
                            Print "Page: " + Str(Page)
                Else
                            'Print the page number
                            CurrentY = 5400:        CurrentX = 700
                            Print "Page: " + Str(Page)
                End If
End If
M.LastRowPrinted = False
'restoring x and y for other previews
M.PosX = 700    'Value in Twips
M.PosY = 1400  'Value in Twips
Enable True
End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label1 = "Prints several pages with a bit details in the first and the last pages "

End Sub

Private Sub Command6_Click()
Dim Page As Integer
Form1.ScaleMode = vbPixels 'this must be done, the container [form1 in this case] must be in vbpixels scalemode
Printer.ScaleMode = vbTwips
Printer.PaperSize = vbPRPSA5
Printer.Orientation = vbPRORPortrait
Printer.Font = LstView.Font.Name
Printer.FontSize = LstView.Font.Size
        
        Page = Page + 1
        Printer.CurrentX = 700
        Printer.CurrentY = 900: Printer.FontSize = 18: Printer.FontName = "Times New Roman": Printer.ForeColor = vbRed
        Printer.Print "Weekly Report": ForeColor = vbBlack
        Printer.FontSize = 8: Printer.FontName = "MS SANS SERIF"
        Printer.CurrentX = 700: Printer.Print "M/s ";: Printer.FontBold = True: Printer.Print "Abhibyakti, A Literary Journal ": Printer.FontBold = False
        Printer.CurrentX = 700: Printer.Print "Mitrapark, Kathmandu, Nepal"
        Printer.CurrentX = 700:      Printer.CurrentY = 2050
        Printer.Print "Report # 9681": Printer.CurrentY = 2050
        Printer.CurrentX = 3000: Printer.Print "Date:" + Str(Now)
        Printer.CurrentX = 700: Printer.CurrentY = 2250
        Printer.Print "Code #. 12441": Printer.CurrentY = 2250
        Printer.CurrentX = 3000: Printer.Print "A/c #: 223-55-889-420"
        Printer.CurrentX = 700:      Printer.CurrentY = 2700
        Printer.FontBold = True: Printer.FontSize = 9: Printer.ForeColor = &H808080
        Printer.Print "Detail of  CREDIT TRANSACTION"
        Printer.FontBold = False: Printer.FontSize = 8: Printer.ForeColor = vbBlack
        '======End of 1st page details=========
        M.PosY = 2950
        M.NumOfRowsPerPage = 8 ' 8 rows on the 1st page
        M.SetRows
        M.PrintHead Printer
        M.PrintBody Printer
        Printer.CurrentY = 5400
        Printer.CurrentX = 700
        Printer.Print "Page: " + Str(Page)
        Printer.NewPage
        M.NumOfRowsPerPage = 16 '16 rows for remaining pages
        M.PosY = 900
        '===========Printing remaining Pages=======
        While Not M.LastRowPrinted
      
                    M.SetRows
                    M.PrintHead Printer
                    M.PrintBody Printer
                    
                    Page = Page + 1
                    If Not M.LastRowPrinted Then
                                Printer.CurrentY = 5400:        Printer.CurrentX = 700
                                Printer.Print "Page: " + Str(Page)
                                Printer.NewPage
                    End If
                   
        Wend
If M.LastRowPrinted Then
                'Planning to print 4 lines at the end of report(Olny if there is room for)
                If (5400 - M.CurrentY) > (4 * Printer.TextHeight("X")) Then
             
                            Printer.CurrentY = M.CurrentY
                            Printer.CurrentX = 700: Printer.Print ""
                            Printer.CurrentX = 700: Printer.Print "End of the detailed transcation list "
                            Printer.CurrentX = 700: Printer.Print "You can view this report on our"
                            Printer.CurrentX = 700: Printer.Print "web page  www.www.www";: Printer.FontBold = True: Printer.Print " Thank you !  ": Printer.FontBold = False
                            Printer.CurrentY = 5400:       Printer.CurrentX = 700
                            Printer.Print "Page: " + Str(Page)
                            Printer.EndDoc
                Else
                            'Print the page number
                            Printer.CurrentY = 5400:        Printer.CurrentX = 700
                            Printer.Print "Page: " + Str(Page)
                            Printer.EndDoc
                End If
End If
M.LastRowPrinted = False
'restoring x and y for other previews
Form1.ScaleMode = vbTwips
M.PosX = 700    'Value in Twips
M.PosY = 1400  'Value in Twips

End Sub



Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label1 = "Paper size A5 (Half of A4)  Orientation Portait  [3 Pages]"
End Sub

Private Sub Command7_Click()

LstView.ListItems(1).ForeColor = vbBlue
LstView.ListItems(1).ListSubItems(2).Bold = True

LstView.ListItems(1).ListSubItems(2).ForeColor = vbRed ' This line does not work! Why?
'LstView.ListItems(1).ListSubItems(2).ForeColor = vbBlue ' This works!

LstView.Refresh
End Sub
   
Private Sub Command7_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label1 = "Turns 1st List Item  to Blue and 2nd Subitem  Bold and Red "
End Sub






Private Sub Form_Load()

Dim IconNum As Integer

Dim itemX As ListItem
Dim clmx As ColumnHeader
Dim i As Integer
Set clmx = LstView.ColumnHeaders.Add(, , "List Item", , , 1)
Set clmx = LstView.ColumnHeaders.Add(, , "First Sub", , , 2)
Set clmx = LstView.ColumnHeaders.Add(, , "Second Sub", , , 3)
'LstView.ColumnHeaders(2).Alignment = lvwColumnRight
Randomize Timer
For i = 1 To 35
        IconNum = Int((11 - 1 + 1) * Rnd + 1)
        Set itemX = LstView.ListItems.Add(, , "List Item #" + Str(i), , IconNum)
        itemX.SubItems(1) = "1st sub #" + Str(i)
        itemX.SubItems(2) = "2nd sub #" + Str(i)
        'IconNum = Int((11 - 1 + 1) * Rnd + 1)

        'itemX.ListSubItems.Add 1, , "1st Sub " + Str(i), IconNum
        'IconNum = Int((11 - 1 + 1) * Rnd + 1)

        'itemX.ListSubItems.Add 2, , "2nd Sub " + Str(i), IconNum

Next
'================================Initial Values
Set M = New ListViewPrinter
Set M.ListViewName = LstView
M.DrawHorizontalLines = True
M.DrawVerticalLines = True
M.DrawBorder = True
M.BorderDistance = 2
M.PosX = 700    'Value in Twips
M.PosY = 1400  'Value in Twips
M.HasPicture = True

'===============================

End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label1 = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set M = Nothing
End Sub
Private Sub Enable(bX As Boolean)
If bX Then MousePointer = 0 Else MousePointer = 11
Command1.Enabled = bX
Command2.Enabled = bX
Command3.Enabled = bX
Command4.Enabled = bX
Command5.Enabled = bX
Command6.Enabled = bX
End Sub


