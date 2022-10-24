VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   8664.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6624
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    
    If (CommandButton1.Caption = "Productos") Then
        CommandButton1.Caption = "Añadir"
        CommandButton2.Caption = "Modificar"
        CommandButton3.Caption = "Volver"
    Else
        If CommandButton1.Caption = "Añadir" Then
            Frame2.Visible = True
            CommandButton2.Enabled = False
            CommandButton3.Enabled = False
            CommandButton1.Enabled = False
        End If
        
    End If
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

Private Sub CommandButton2_Click()
    If CommandButton2.Caption = "Modificar" Then
        Frame2.Visible = True
        Frame2.Caption = "Modificar Producto"
    Else
        If CommandButton2.Caption = "Ventas" Then
            CommandButton1.Caption = "Añadir"
            CommandButton2.Caption = "Modificar"
            CommandButton3.Caption = "Eliminar"
            CommandButton7.Caption = "Volver"
            CommandButton7.Visible = True
            Frame2.Visible = True
        End If
    End If
End Sub
