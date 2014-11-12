Option Explicit
Option Base 1

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BatchConvertCSVtoXLSX - 16/10/2014
'
' Zczytuje wszystkie pliki CSV z podanego folderu, zczytuje pierwszy wiersz każdego pliku
' w poszukiwaniu nazwa MEMBERÓW, następnie zapisuje w formacie XLSX z formatem TEXT
' dla kolumn ze stałej DoSprawdzenia, by zachować "zera" na początku.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub BatchConvertCSVtoXLSX()
Dim WB As Workbook
Dim strFile As String, strDir As String, strOut_Dir As String
Dim columnFormats() As Integer
Dim i As Long, x As Long
Dim WS As Excel.Worksheet
Dim tempArr() As String, tempArr2() As String, tempName As String, strFirstLine As String
Dim DoSprawdzenia As Variant

DoSprawdzenia = Array("YOUR_EEID", "YOUR_EEID_ORIG", "YOUR_CODE", "YOUR_LEVEL", "YOUR_GRADE", "YOUR_OTHER")

strDir = "H:\POSDATA_all\MLS\" 'location of csv files
strOut_Dir = "H:\POSDATA_all\xlsx_MLS\" 'location of xlsx files
strFile = Dir(strDir & "*.csv")

' Zapisanie wszystkich plików z katalogu w tablicy
ReDim tempArr(1 To 10)
i = 1
Do While strFile <> ""
    tempArr(i) = strFile
    i = i + 1
    strFile = Dir
Loop
ReDim Preserve tempArr(1 To i - 1)
strFile = vbNullString

Call QuickSort(tempArr, 1, 7)

' Właściwa pętla do zmiany formatu i ustawienia "Cell Format" to TEXT
For i = 1 To UBound(tempArr)

    ' Zczytanie pierwszego wiersza z nazwami memberów z podanego pliku z tempArr(i)
    Open strDir & tempArr(i) For Input As #1
    Line Input #1, strFirstLine
        tempArr2 = Split(Right(strFirstLine, Len(strFirstLine) - 2), vbTab)
        strFirstLine = vbNullString
    Close #1

    ' przypisanie do wybranych formatu TEXT
    ReDim columnFormats(1 To UBound(tempArr2))
        For x = 1 To UBound(tempArr2)
            If IsInArray(tempArr2(x), DoSprawdzenia, False) = True Then
                columnFormats(x) = xlTextFormat
            Else
                columnFormats(x) = xlGeneralFormat
            End If
        Next x

    If Application.Workbooks.Count < 2 Then
        Application.Workbooks.Add
    End If
    
    Set WS = Excel.ActiveSheet
    With WS.QueryTables.Add("TEXT;" & strDir & tempArr(i), WS.Cells(1, 1))
        .FieldNames = True
        .AdjustColumnWidth = False
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = columnFormats
        .Refresh
    End With
    
    With ActiveWorkbook
        Application.DisplayAlerts = False
            If .Sheets.Count >= 3 Then
                .Sheets(1).Name = Left(tempArr(i), Len(tempArr(i)) - 4)
                .Sheets(2).Delete
                .Sheets(2).Delete
            End If
        Application.DisplayAlerts = True
        .SaveAs strOut_Dir & Replace(tempArr(i), ".csv", ".xlsx"), 51
        .Close True
    End With
Next i
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsInArray - 17/10/2014
'
' Sprawdza czy dany string wystepuje w tablicy
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function IsInArray(StringToBeFound As String, ArrayToSeach As Variant, Optional ExactMatch As Boolean = True) As Boolean

Dim i As Long
Dim r As Variant
' default return value if value not found in array
IsInArray = False

If ExactMatch = True Then
    For i = 1 To UBound(ArrayToSeach)
        If StrComp(StringToBeFound, ArrayToSeach(i), vbBinaryCompare) = 0 Then
            IsInArray = True
            Exit For
        End If
    Next i
Else
    For i = 1 To UBound(ArrayToSeach)
        If Len(StringToBeFound) <= Len(ArrayToSeach(i)) Then
            If InStr(CStr(ArrayToSeach(i)), CStr(StringToBeFound)) > 0 Then
                IsInArray = True
                Exit For
            End If
        End If
    Next i
End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' QuickSort - 17/10/2014
'
' Sortowanie tablicy algorytmem QuickSort
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub QuickSort(arr, Lo As Long, Hi As Long)
  Dim varPivot As Variant
  Dim varTmp As Variant
  Dim TmpLow As Long
  Dim tmpHi As Long
  TmpLow = Lo
  tmpHi = Hi
  varPivot = arr((Lo + Hi) \ 2)
  Do While TmpLow <= tmpHi
    Do While arr(TmpLow) < varPivot And TmpLow < Hi
      TmpLow = TmpLow + 1
    Loop
    Do While varPivot < arr(tmpHi) And tmpHi > Lo
      tmpHi = tmpHi - 1
    Loop
    If TmpLow <= tmpHi Then
      varTmp = arr(TmpLow)
      arr(TmpLow) = arr(tmpHi)
      arr(tmpHi) = varTmp
      TmpLow = TmpLow + 1
      tmpHi = tmpHi - 1
    End If
  Loop
  If Lo < tmpHi Then QuickSort arr, Lo, tmpHi
  If TmpLow < Hi Then QuickSort arr, TmpLow, Hi
End Sub
