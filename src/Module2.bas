Attribute VB_Name = "Module2"
Option Explicit

Sub iroget()
   MsgBox ActiveCell.Interior.ColorIndex
End Sub

Sub irokae()

Const Beforeiro1 = 20
Const Beforeiro2 = 2

Const Afteriro = 1

Dim Selectcells As Range
Dim TCell As Range
Set Selectcells = Selection

For Each TCell In Selectcells
   If TCell.Interior.ColorIndex = Beforeiro1 Or _
      TCell.Interior.ColorIndex = Beforeiro2 Then
      TCell.Interior.ColorIndex = Afteriro
   End If
Next TCell

End Sub

Sub MapSakusei()

Dim i As Integer
Dim i2 As Integer
Application.ScreenUpdating = False

For i = 1 To 224
   For i2 = 1 To 288
      Cells(1000 + i2, i).Value = Cells(1300 + i2, i).Interior.ColorIndex
   Next i2
Next i

End Sub

Sub test2()
 Dim a As Variant
 a = Worksheets("mapdata").Range(Cells(1, 1), Cells(288, 224))
End Sub

Sub MsgIndex()
   MsgBox Selection.Interior.ColorIndex
End Sub
