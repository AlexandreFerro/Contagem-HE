Attribute VB_Name = "Módulo3"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"

' Macro1 Macro
' Conta os dias do mês em que ocorreram as HE

Sheets("Escala1").Select

i2 = 30

For x = 16 To 47

For y = 5 To 34

Cells(x, y).Activate


' Procura as palavra HEA,HEB,HEC,HE/A,HE/B,HE/C,A/HE,B/HE,C/HE na planilha
If Cells(x, y).Value = "HEA" Or Cells(x, y).Value = "HEB" Or Cells(x, y).Value = "HEC" Or Cells(x, y).Value = "HE/A" Or Cells(x, y).Value = "HE/B" Or Cells(x, y).Value = "HE/C" Or Cells(x, y).Value = "A/HE" Or Cells(x, y).Value = "B/HE" Or Cells(x, y).Value = "C/HE" Then

Nome = Cells(x, 3).Value
Matrícula = Cells(x, 4).Value
Coluna = ActiveCell.Column
Per = Cells(15, Coluna).Value
Hor = Cells(x, y).Value

Sheets("Formulário").Select

' Doc final recebe dados
    Cells(i2, 2).Value = Nome
    Columns("A:A").EntireColumn.AutoFit

' Cola matrícula

     Cells(i2, 1).Value = Matrícula
     Columns("A:A").EntireColumn.AutoFit

' Cola Lotação
    Cells(i2, 3).Value = "TAKP"
    Columns("A:A").EntireColumn.AutoFit
    
' Cola Período
Cells(i2, 4).Value = Per & "/09/2016"
Columns("A:A").EntireColumn.AutoFit

If Hor = "HEA" Then
    
    Horario = "7 as 15"
    Cells(i2, 5).Value = Horario
    Columns("A:A").EntireColumn.AutoFit
    
End If

If Hor = "HEB" Then
    
    Horario = "15 as 23"
    Cells(i2, 5).Value = Horario
    Columns("A:A").EntireColumn.AutoFit
    
End If

If Hor = "HEC" Then
    
    Horario = "23 as 7"
    Cells(i2, 5).Value = Horario
    Columns("A:A").EntireColumn.AutoFit
    
End If

If Hor = "HE/A" Then
    
    Horario = "3 as 7"
    Cells(i2, 5).Value = Horario
    Columns("A:A").EntireColumn.AutoFit
    
End If

If Hor = "HE/B" Then
    
    Horario = "11 as 15"
        Cells(i2, 5).Value = Horario
    Columns("A:A").EntireColumn.AutoFit
    
End If

If Hor = "HE/C" Then
    
    Horario = "19 as 23"
    Cells(i2, 5).Value = Horario
    Columns("A:A").EntireColumn.AutoFit
    
End If

If Hor = "A/HE" Then
    
    Horario = "15 as 19"
    Cells(i2, 5).Value = Horario
    Columns("A:A").EntireColumn.AutoFit
    
End If

If Hor = "B/HE" Then
    
    Horario = "23 as 3"
    Cells(i2, 5).Value = Horario
    Columns("A:A").EntireColumn.AutoFit
    
End If

If Hor = "C/HE" Then
    
    Horario = "7 as 11"
    Cells(i2, 5).Value = Horario
    Columns("A:A").EntireColumn.AutoFit

    
End If


i2 = i2 + 1

End If

Sheets("Escala1").Select

Next y
Next x

Sheets("Formulário").Select

End Sub

