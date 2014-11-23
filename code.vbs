Option Explicit

Sub Template()

Application.ScreenUpdating = False 'change to False    --- False
Application.Calculation = xlCalculationManual 'change to xlCalculationManual       xlCalculationAutomatic

  Dim wbResults, wbCodeBook, wbTemplate As Workbook
  Dim wsTemplateSheet, wsSummarySheet, wsQuerySheet, wsMacroSheet As Worksheet
  Dim wbCodeBookName, Column, ColumnA, strPath, strFile, strDir As String
  Dim i, lr, j, t, totalCols As Integer
  Dim vCell, rng As Range
  Dim c, Record(1700, 30) As Variant ' first part of record is how many rows, second part is how many columns
  Dim rowCounter As Long
  Dim toggle As Boolean
  
  'wbCodeBook = Application.ActiveWorkbook
  wbCodeBookName = ActiveWorkbook.Name
  
  Set wsSummarySheet = ActiveWorkbook.Sheets("Query")
  Set wsQuerySheet = ActiveWorkbook.Sheets("Template")
  Set wsMacroSheet = ActiveWorkbook.Sheets("Run Macro")

  'Where are you aiming to grab files from?
  strPath = ActiveWorkbook.Worksheets("Run Macro").Range("E11").Value
  'Ensure the path ends in a backslash
  If Right(strPath, 1) <> "\" Then
      strPath = strPath & "\"
  End If
  
  wsSummarySheet.Select
  lr = Cells(Rows.Count, 9).End(xlUp).Row  '9 here is the column length, ie column I
  c = Range("i4:i" & lr)
  For i = 1 To UBound(c, 1) Step 2 ' ie length of rows
    Column = "I"
    ColumnA = "A"
    toggle = False
    For j = 1 To 30  ' ie length of helper method columns '2011, 2012, and 2013
      Record(i, j) = wsSummarySheet.Range(Column & (i + 3)).Value ' Volumes
      Record(i + 1, j) = wsSummarySheet.Range(Column & (i + 4)).Value 'Count
      
      
      If toggle Then
        Column = "A"
        Column = "A" & ColumnA
        ColumnA = Chr(Asc(ColumnA) + 1)
      Else
        Column = Chr(Asc(Column) + 1)
        If Column = "Z" Then toggle = True
      End If
    Next j
    
    
    
  'toggle = True
  
    'Do While toggle
      wsQuerySheet.Select
      '   ## select new workbook so range values work correctly
      Range("b3").Value = Record(i, 1)  ' Account ID
      Range("b2").Value = Record(i, 2)  ' Name
      
      'Volume
      Range("b24").Value = Record(i, 4)   ' 2011 Visa
      Range("b25").Value = Record(i, 5)   ' 2011 MC
      Range("b26").Value = Record(i, 6)   ' 2011 Disc
      Range("b27").Value = Record(i, 7)   ' 2011 eCheck
      Range("b28").Value = Record(i, 8)   ' 2011 Amex
      Range("b29").Value = Record(i, 9)   ' 2011 Scan
      Range("b30").Value = Record(i, 10)  ' 2011 Debit
      Range("b31").Value = ""      'Record(i, 11)  ' 2011 MMF
      Range("b32").Value = ""      'Record(i, 12)  ' 2011 Pay by Cash
      
      Range("b38").Value = Record(i, 13)   ' 2012 Visa
      Range("b39").Value = Record(i, 14)   ' 2012 MC
      Range("b40").Value = Record(i, 15)   ' 2012 Disc
      Range("b41").Value = Record(i, 16)   ' 2012 eCheck
      Range("b42").Value = Record(i, 17)   ' 2012 Amex
      Range("b43").Value = Record(i, 18)   ' 2012 Scan
      Range("b44").Value = Record(i, 19)  ' 2012 Debit
      Range("b45").Value = ""      'Record(i, 20)  ' 2012 MMF
      Range("b46").Value = ""      'Record(i, 21)  ' 2012 Pay by Cash
      
      Range("b52").Value = Record(i, 22)   ' 2013 Visa
      Range("b53").Value = Record(i, 23)   ' 2013 MC
      Range("b54").Value = Record(i, 24)   ' 2013 Disc
      Range("b55").Value = Record(i, 25)   ' 2013 eCheck
      Range("b56").Value = Record(i, 26)   ' 2013 Amex
      Range("b57").Value = Record(i, 27)   ' 2013 Scan
      Range("b58").Value = Record(i, 28)  ' 2013 Debit
      Range("b59").Value = ""      'Record(i, 29)  ' 2013 MMF
      Range("b60").Value = ""      'Record(i, 30)  ' 2013 Pay by Cash
      
      'Count
      Range("e24").Value = Record(i + 1, 4)  ' 2011 Visa
      Range("e25").Value = Record(i + 1, 5)  ' 2011 MC
      Range("e26").Value = Record(i + 1, 6)  ' 2011 Disc
      Range("e27").Value = Record(i + 1, 7)  ' 2011 eCheck
      Range("e28").Value = Record(i + 1, 8)  ' 2011 Amex
      Range("e29").Value = Record(i + 1, 9)  ' 2011 Scan
      Range("e30").Value = Record(i + 1, 10)  ' 2011 Debit
      Range("e31").Value = ""     ' Record(i + 1, 5)  ' 2011 MMF
      Range("e32").Value = ""     ' Record(i + 1, 5)  ' 2011 Pay by Cash
      
      Range("e38").Value = Record(i + 1, 13)   ' 2012 Visa
      Range("e39").Value = Record(i + 1, 14)   ' 2012 MC
      Range("e40").Value = Record(i + 1, 15)   ' 2012 Disc
      Range("e41").Value = Record(i + 1, 16)   ' 2012 eCheck
      Range("e42").Value = Record(i + 1, 17)   ' 2012 Amex
      Range("e43").Value = Record(i + 1, 18)   ' 2012 Scan
      Range("e44").Value = Record(i + 1, 19)  ' 2012 Debit
      Range("e45").Value = ""      'Record(i + 1, 20)  ' 2012 MMF
      Range("e46").Value = ""      'Record(i + 1, 21)  ' 2012 Pay by Cash
      
      Range("e52").Value = Record(i + 1, 22)   ' 2013 Visa
      Range("e53").Value = Record(i + 1, 23)   ' 2013 MC
      Range("e54").Value = Record(i + 1, 24)   ' 2013 Disc
      Range("e55").Value = Record(i + 1, 25)   ' 2013 eCheck
      Range("e56").Value = Record(i + 1, 26)   ' 2013 Amex
      Range("e57").Value = Record(i + 1, 27)   ' 2013 Scan
      Range("e58").Value = Record(i + 1, 28)  ' 2013 Debit
      Range("e59").Value = ""      'Record(i + 1, 29)  ' 2013 MMF
      Range("e60").Value = ""      'Record(i + 1, 30)  ' 2013 Pay by Cash
      
      ' ## save workbook with new name (add account ID to the front of the name)
      Sheets("Template").Select
      Sheets("Template").Copy
      ChDir strPath
      Application.AlertBeforeOverwriting = False
      Application.DisplayAlerts = False
      ActiveWorkbook.SaveAs Filename:= _
      strPath & Record(i, 1) & " - 2template.xlsx" _
      , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
      ActiveWindow.Close
      Application.AlertBeforeOverwriting = True
      Application.DisplayAlerts = True
      ' wbCodeBook.Activate
      
      Range("b2").Value = ""             ' Account ID
      Range("b3").Value = ""             ' Account ID

      Range("b24").Value = ""   ' 2011 Visa
      Range("b25").Value = ""   ' 2011 MC
      Range("b26").Value = ""   ' 2011 Disc
      Range("b27").Value = ""   ' 2011 eCheck
      Range("b28").Value = ""   ' 2011 Amex
      Range("b29").Value = ""   ' 2011 Scan
      Range("b30").Value = ""   ' 2011 Debit
      Range("b31").Value = ""      'Record(i, 11)  ' 2011 MMF
      Range("b32").Value = ""      'Record(i, 12)  ' 2011 Pay by Cash
      Range("e24").Value = ""  ' 2011 Visa
      Range("e25").Value = ""  ' 2011 MC
      Range("e26").Value = ""  ' 2011 Disc
      Range("e27").Value = ""  ' 2011 eCheck
      Range("e28").Value = ""  ' 2011 Amex
      Range("e29").Value = ""  ' 2011 Scan
      Range("e30").Value = ""  ' 2011 Debit
      Range("e31").Value = ""     ' Record(i + 1, 5)  ' 2011 MMF
      Range("e32").Value = ""     ' Record(i + 1, 5)  ' 2011 Pay by Cash
      
      Range("b38").Value = ""   ' 2011 Visa
      Range("b39").Value = ""   ' 2011 MC
      Range("b40").Value = ""   ' 2011 Disc
      Range("b41").Value = ""   ' 2011 eCheck
      Range("b42").Value = ""   ' 2011 Amex
      Range("b43").Value = ""   ' 2011 Scan
      Range("b44").Value = ""   ' 2011 Debit
      Range("b45").Value = ""      'Record(i, 11)  ' 2011 MMF
      Range("b46").Value = ""      'Record(i, 12)  ' 2011 Pay by Cash
      Range("e38").Value = ""  ' 2011 Visa
      Range("e39").Value = ""  ' 2011 MC
      Range("e40").Value = ""  ' 2011 Disc
      Range("e41").Value = ""  ' 2011 eCheck
      Range("e42").Value = ""  ' 2011 Amex
      Range("e43").Value = ""  ' 2011 Scan
      Range("e44").Value = ""  ' 2011 Debit
      Range("e45").Value = ""     ' Record(i + 1, 5)  ' 2011 MMF
      Range("e46").Value = ""     ' Record(i + 1, 5)  ' 2011 Pay by Cash
      
      Range("b52").Value = ""   ' 2011 Visa
      Range("b53").Value = ""   ' 2011 MC
      Range("b54").Value = ""   ' 2011 Disc
      Range("b55").Value = ""   ' 2011 eCheck
      Range("b56").Value = ""   ' 2011 Amex
      Range("b57").Value = ""   ' 2011 Scan
      Range("b58").Value = ""   ' 2011 Debit
      Range("b59").Value = ""      'Record(i, 11)  ' 2011 MMF
      Range("b60").Value = ""      'Record(i, 12)  ' 2011 Pay by Cash
      Range("e52").Value = ""  ' 2011 Visa
      Range("e53").Value = ""  ' 2011 MC
      Range("e54").Value = ""  ' 2011 Disc
      Range("e55").Value = ""  ' 2011 eCheck
      Range("e56").Value = ""  ' 2011 Amex
      Range("e57").Value = ""  ' 2011 Scan
      Range("e58").Value = ""  ' 2011 Debit
      Range("e59").Value = ""     ' Record(i + 1, 5)  ' 2011 MMF
      Range("e60").Value = ""     ' Record(i + 1, 5)  ' 2011 Pay by Cash
      
    '  toggle = False
    'Loop
    
  Next i
  
  wsMacroSheet.Select
  
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Sub
