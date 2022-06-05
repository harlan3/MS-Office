VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Calculator"
   ClientHeight    =   6600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4785
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim num1 As Double
Dim num2 As Double
Dim op As String
Dim result As Double

Private Sub CommandButton1_Click()
If TextBox1.Text = "0" Then
TextBox1.Text = "7"
Else
TextBox1.Text = TextBox1.Text & "7"
End If
End Sub

Private Sub CommandButton2_Click()
If TextBox1.Text = "0" Then
TextBox1.Text = "8"
Else
TextBox1.Text = TextBox1.Text & "8"
End If
End Sub

Private Sub CommandButton3_Click()
If TextBox1.Text = "0" Then
TextBox1.Text = "9"
Else
TextBox1.Text = TextBox1.Text & "9"
End If
End Sub

Private Sub CommandButton4_Click()
num1 = TextBox1.Text
TextBox1.Text = 0
op = "+"
End Sub

Private Sub CommandButton5_Click()
If TextBox1.Text = "0" Then
TextBox1.Text = "4"
Else
TextBox1.Text = TextBox1.Text & "4"
End If
End Sub

Private Sub CommandButton6_Click()
If TextBox1.Text = "0" Then
TextBox1.Text = "5"
Else
TextBox1.Text = TextBox1.Text & "5"
End If
End Sub

Private Sub CommandButton7_Click()
If TextBox1.Text = "0" Then
TextBox1.Text = "6"
Else
TextBox1.Text = TextBox1.Text & "6"
End If
End Sub

Private Sub CommandButton8_Click()
num1 = TextBox1.Text
TextBox1.Text = 0
op = "-"
End Sub

Private Sub CommandButton9_Click()
If TextBox1.Text = "0" Then
TextBox1.Text = "1"
Else
TextBox1.Text = TextBox1.Text & "1"
End If
End Sub

Private Sub CommandButton10_Click()
If TextBox1.Text = "0" Then
TextBox1.Text = "2"
Else
TextBox1.Text = TextBox1.Text & "2"
End If
End Sub

Private Sub CommandButton11_Click()
If TextBox1.Text = "0" Then
TextBox1.Text = "3"
Else
TextBox1.Text = TextBox1.Text & "3"
End If
End Sub

Private Sub CommandButton12_Click()
num1 = TextBox1.Text
TextBox1.Text = 0
op = "x"
End Sub

Private Sub CommandButton13_Click()
If TextBox1.Text = "0" Then
TextBox1.Text = "0"
Else
TextBox1.Text = TextBox1.Text & "0"
End If
End Sub

Private Sub CommandButton14_Click()
If TextBox1.Text = "0" Then
TextBox1.Text = "."
Else
TextBox1.Text = TextBox1.Text & "."
End If
End Sub

Private Sub CommandButton15_Click()
num2 = TextBox1.Text
If op = "+" Then
result = num1 + num2
TextBox1.Text = result
ElseIf op = "-" Then
result = num1 - num2
TextBox1.Text = result
ElseIf op = "x" Then
result = num1 * num2
TextBox1.Text = result
ElseIf op = "/" Then
result = num1 / num2
TextBox1.Text = result
End If
End Sub

Private Sub CommandButton16_Click()
num1 = TextBox1.Text
TextBox1.Text = 0
op = "/"
End Sub

Private Sub CommandButton17_Click()
TextBox1.Text = 0
End Sub
