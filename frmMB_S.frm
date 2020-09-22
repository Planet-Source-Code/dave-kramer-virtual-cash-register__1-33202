VERSION 5.00
Begin VB.Form frmMasterCashRegister 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Master Burger - Cash Register"
   ClientHeight    =   4290
   ClientLeft      =   2925
   ClientTop       =   1980
   ClientWidth     =   5715
   DrawStyle       =   5  'Transparent
   FillColor       =   &H00C0C000&
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4290
   ScaleWidth      =   5715
   Begin VB.Frame V 
      Caption         =   "Delivery"
      Height          =   1695
      Left            =   120
      TabIndex        =   27
      Top             =   1800
      Width           =   2535
      Begin VB.OptionButton optDelivery 
         Caption         =   "Eat In"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   31
         Top             =   1320
         Width           =   1335
      End
      Begin VB.OptionButton optDelivery 
         Caption         =   "Zone 3"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optDelivery 
         Caption         =   "Zone 2"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton optDelivery 
         Caption         =   "Zone 1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.VScrollBar vbsSoftDrink 
      Height          =   375
      Left            =   1560
      Max             =   0
      Min             =   25
      TabIndex        =   26
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.VScrollBar vbsFrenchFries 
      Height          =   375
      Left            =   1560
      Max             =   0
      Min             =   25
      TabIndex        =   25
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.VScrollBar vbsBurger 
      Height          =   375
      Left            =   1560
      Max             =   0
      Min             =   25
      TabIndex        =   24
      Top             =   360
      Value           =   25
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox chkSoftDrink 
      Caption         =   "Soft Drink"
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CheckBox chkFries 
      Caption         =   "French Fries"
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   840
      Width           =   1335
   End
   Begin VB.CheckBox chkBurger 
      Caption         =   "Burger"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   735
      Left            =   1680
      Picture         =   "FRMMB_S.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3600
      Width           =   865
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   735
      Left            =   840
      Picture         =   "FRMMB_S.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Reciept"
      Height          =   735
      Left            =   0
      Picture         =   "FRMMB_S.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3600
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      BorderWidth     =   5
      X1              =   5505
      X2              =   3480
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label lblPriceSoftDrink 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   2640
      TabIndex        =   20
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblPriceFrenchFries 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   2640
      TabIndex        =   19
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblPriceBurger 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   2640
      TabIndex        =   18
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Price"
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
      Left            =   2640
      TabIndex        =   17
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblDelivery 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   4560
      TabIndex        =   15
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Delivery: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   14
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   4560
      TabIndex        =   11
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label lblSalesTax 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   4560
      TabIndex        =   10
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblExtSoftDrink 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   4560
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblExtFrenchFries 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblExtBurger 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblQtySoftDrink 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblQtyFrenchFries 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblQtyBurger 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   2520
      X2              =   2520
      Y1              =   0
      Y2              =   4320
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ext. Price"
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
      Left            =   4560
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Quantity"
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
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total:  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sales Tax: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
End
Attribute VB_Name = "frmMasterCashRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const csngPriceBurger As Single = 1.44
Private Const csngPriceFrenchFries As Single = 0.74
Private Const csngPriceSoftDrink As Single = 0.66
Private Const csngPercentTax As Single = 0.07










Private Sub chkBurger_Click()
    If chkBurger.Value = vbChecked Then
        vbsBurger.Visible = True
        vbsBurger.Value = 1
        lblPriceBurger.Visible = True
        lblQtyBurger.Visible = True
        lblExtBurger.Visible = True
    Else
        vbsBurger.Visible = False
        vbsBurger.Value = 0
        lblPriceBurger.Visible = False
        lblQtyBurger.Visible = False
        lblExtBurger.Visible = False
    End If
    
     
End Sub

Private Sub chkFries_Click()
     If chkFries.Value = vbChecked Then
        vbsFrenchFries.Visible = True
        vbsFrenchFries.Value = 1
        lblPriceFrenchFries.Visible = True
        lblQtyFrenchFries.Visible = True
        lblExtFrenchFries.Visible = True
    Else
        vbsFrenchFries.Visible = False
        vbsFrenchFries.Value = 0
        lblPriceFrenchFries.Visible = False
        lblQtyFrenchFries.Visible = False
        lblExtFrenchFries.Visible = False
    End If
End Sub

Private Sub chkSoftDrink_Click()
     If chkSoftDrink.Value = vbChecked Then
        vbsSoftDrink.Visible = True
        vbsSoftDrink.Value = 1
        lblPriceSoftDrink.Visible = True
        lblQtySoftDrink.Visible = True
        lblExtSoftDrink.Visible = True
    Else
        vbsSoftDrink.Visible = False
        vbsSoftDrink.Value = 0
        lblPriceSoftDrink.Visible = False
        lblQtySoftDrink.Visible = False
        lblExtSoftDrink.Visible = False
    End If
End Sub



Private Sub cmdClear_Click()
    chkBurger.Value = vbUnchecked
    chkFries.Value = Unchecked
    chkSoftDrink.Value = Unchecked
    optDelivery(3).Value = True
End Sub

'
' Display a message box to confirm that the user wants
' to exit the program.
Private Sub cmdExit_Click()
    Dim pintReturn As Integer
    pintReturn = MsgBox("Really Exit?", _
        vbYesNo + vbQuestion, "Master Burger")
    If pintReturn = vbYes Then
        Unload frmMasterCashRegister
    End If
End Sub

Private Sub cmdPrint_Click()
Printer.Print lblExtBurger
Printer.Print lblExtFrenchFries
Printer.Print lblDelivery
Printer.Print lblSalesTax
Printer.Print lblTotal
'frmmastercashregister.printform

End Sub

Private Sub lblDelivery_Change()
    ComputeTotals
    
End Sub



Private Sub lblExtBurger_Change()
    ComputeTotals
End Sub



Private Sub lblExtFrenchFries_Change()
ComputeTotals
End Sub



Private Sub lblExtSoftDrink_Change()
    ComputeTotals
End Sub



Private Sub optDelivery_Click(Index As Integer)
    Select Case Index
        Case 0    'Zone 1
            lblDelivery.Caption = Format(1.5, "fixed")
        Case 1   'Zone 2
            lblDelivery.Caption = Format(2.5, "fixed")
        Case 2    'Zone 3
            lblDelivery.Caption = Format(3.5, "fixed")
        Case 3    ' Eat in
            lblDelivery.Caption = Format(0, "fixed")
        End Select
        
End Sub

Private Sub Form_Load()
    optDelivery(3).Value = True
    lblPriceBurger = csngPriceBurger
    lblPriceFrenchFries = csngPriceFrenchFries
    lblPriceSoftDrink = csngPriceSoftDrink
    
End Sub

Private Sub vbsBurger_Change()
    Dim psngextburger As Single
    lblQtyBurger.Caption = vbsBurger.Value
    psngextburger = lblQtyBurger.Caption * _
        csngPriceBurger
    lblExtBurger.Caption = Format(psngextburger, "fixed")
    
End Sub

Private Sub vbsFrenchFries_Change()
    Dim psngextfrenchfries As Single
    lblQtyFrenchFries.Caption = vbsFrenchFries
    psngextfrenchfries = lblQtyFrenchFries.Caption * csngPriceFrenchFries
    lblExtFrenchFries.Caption = Format(psngextfrenchfries, "fixed")
    
    
End Sub

Private Sub vbsSoftDrink_Change()
    Dim psngextsoftdrink As Single
    lblQtySoftDrink.Caption = vbsSoftDrink
    psngextsoftdrink = lblQtySoftDrink.Caption * csngPriceSoftDrink
    lblExtSoftDrink.Caption = Format(psngextsoftdrink, "fixed")
    
End Sub

Private Sub ComputeTotals()
    Dim psngsubtotal As Single
    Dim psngsalestax As Single
    psngsubtotal = Val(lblExtBurger.Caption) + _
    Val(lblExtFrenchFries.Caption) + _
    Val(lblExtSoftDrink.Caption)
psngsalestax = csngPercentTax * psngsubtotal
lblSalesTax.Caption = Format(psngsalestax, "fixed")
lblTotal.Caption = Format(psngsubtotal + psngsalestax _
    + lblDelivery.Caption, "fixed")
    
End Sub
