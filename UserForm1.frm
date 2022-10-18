VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   8616.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6564
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    CommandButton1.Caption = "Añadir"
    CommandButton2.Caption = "Eliminar"
    CommandButton3.Caption = "Volver"
    Frame2.Visible = True
    
End Sub

Private Sub limpiar_Frame(ByVal numfilas As Integer, ByVal mostrar As Boolean, ByVal txtBtnProduc As String, ByVal txtBtnVentas As String, ByVal txtBtnAnalisis As String)
    CommandButton1.Caption = txtBtnProduc
    CommandButton2.Caption = txtBtnVentas
    CommandButton3.Caption = txtBtnAnalisis
    'textbox6 corresponde al ID de productos'
    TextBox6 = numfilas
    TextBox1 = ""
    TextBox2 = ""
    TextBox3 = ""
    TextBox4 = ""
End Sub



