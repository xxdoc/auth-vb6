VERSION 5.00
Begin VB.Form frmStart 
   Caption         =   "BTF Auth Example"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2550
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   2550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BTFCustomers 
      Caption         =   "Find a customer"
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton BTFSale 
      Caption         =   "Authorize Sale/Hold/Credit"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CommandButton BTFVoid 
      Caption         =   "Void a transaction"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   1560
      Width           =   2535
   End
   Begin VB.CommandButton BTFDepositCollect 
      Caption         =   "Deposit Collect"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   2040
      Width           =   2535
   End
   Begin VB.CommandButton Exit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label BTFLabel 
      BackColor       =   &H8000000D&
      Caption         =   "BlueTarp Authorization"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub BTFCustomers_Click()
    frmBTFCustomer.Show
End Sub


Private Sub BTFSale_Click()
    frmBTFSale.Show
End Sub


Private Sub BTFVoid_Click()
    frmBTFVoidTransaction.Show
End Sub

Private Sub BTFDepositCollect_Click()
    frmBTFDepositCollect.Show
End Sub


Private Sub Exit_Click()
    Unload Me
End Sub
