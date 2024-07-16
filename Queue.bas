Attribute VB_Name = "Queue"
Option Explicit
Private dictJobDay As New Dictionary
Private Sub Queue()
Dim i&, Q&
Dim arrDataQuality As Variant
Dim arrDataDPP As Variant
Dim CheckDate As Variant
Dim arrTmp As Variant
Dim arrSP As Variant
Dim dictLDayUDay As New Dictionary

arrDataQuality = [DataQuality].Columns(1).value2
arrDataDPP = [DataDPP].Columns(3).Resize(, 2).value2

For i = LBound(arrDataQuality) To UBound(arrDataQuality)
    dictJobDay(arrDataQuality(i, 1)) = dictJobDay(arrDataQuality(i, 1)) + 0
Next i

For i = LBound(arrDataDPP) To UBound(arrDataDPP)
    CheckDate = CheckDateTime(arrDataDPP(i, 1), arrDataDPP(i, 2))
    If CheckDate(0) Then
        dictLDayUDay(CheckDate(1) & "|" & CheckDate(2)) = dictLDayUDay(CheckDate(1) & "|" & CheckDate(2)) + 1
    End If
Next i

arrTmp = dictLDayUDay.Keys

For i = LBound(arrTmp) To UBound(arrTmp)
    arrSP = Split(arrTmp(i), "|")
    CheckDate = CalcDate(arrSP(0), arrSP(1))
    If CheckDate(0) Then
    
        For Q = CheckDate(1) To CheckDate(2)
            If dictJobDay.Exists(Q) Then dictJobDay(Q) = dictJobDay(Q) + dictLDayUDay(arrTmp(i))
        Next Q

    End If
Next i

[DataQuality].Columns(2) = arr_1xTranspose(dictJobDay.Items)
dictJobDay.RemoveAll

End Sub
Private Function CheckDateTime(ByVal L_DateTime As Double, ByVal U_DateTime As Double) As Variant
Dim LDate As Long
Dim UDate As Long
Dim LTime As Double
Dim UTime As Double

LDate = Int(L_DateTime)
UDate = Int(U_DateTime)
LTime = L_DateTime - LDate
UTime = U_DateTime - UDate

If LDate = UDate Then
    If LTime > 0.375 Then CheckDateTime = Array(False, 0, 0): Exit Function
    If LTime < 0.375 And UTime < 0.375 Then CheckDateTime = Array(False, 0, 0): Exit Function
End If

If LTime > 0.375 And UTime < 0.375 Then CheckDateTime = Array(True, LDate + 1, UDate - 1): Exit Function
If UTime < 0.375 Then CheckDateTime = Array(True, LDate, UDate - 1): Exit Function
If LTime > 0.375 Then CheckDateTime = Array(True, LDate + 1, UDate): Exit Function

CheckDateTime = Array(True, LDate, UDate)

End Function
Private Function CalcDate(ByVal LDate As Long, ByVal UDate As Long) As Variant
Dim CheckLBound As Boolean
Dim CheckUBound As Boolean
CheckLBound = True: CheckUBound = True

While CheckUBound = True
    If Not dictJobDay.Exists(UDate) Then
        UDate = UDate - 1
        CheckUBound = True
    Else
        CheckUBound = False
    End If
    If UDate < LDate Then CheckUBound = False: CalcDate = Array(False, LDate, UDate): Exit Function
Wend

While CheckLBound = True
    If Not dictJobDay.Exists(LDate) Then
        LDate = LDate + 1
        CheckLBound = True
    Else
        CheckLBound = False
    End If
Wend
CalcDate = Array(True, LDate, UDate)

End Function
Private Function arr_1xTranspose(arr As Variant) As Variant
Dim i&, G&
Dim arrResult As Variant

ReDim arrResult(LBound(arr) To UBound(arr), 1 To 1)
For i = LBound(arrResult) To UBound(arrResult)
    arrResult(i, 1) = arr(i)
Next i

arr_1xTranspose = arrResult

End Function
