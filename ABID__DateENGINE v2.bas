'v2
Option Explicit

' =====================================================================
' ENGINE: ABID__DateENGINE (v3 - Multi-Language Resource Loaded)
' =====================================================================
' Global variables accessible from any other macro

' [1] Date Numbers
Public Abid_DEngine_BanglaDate_dayNum_bn As String
Public Abid_DEngine_BanglaDate_dayNum_bn_pre As String ' <--- NEW: Contains day number with suffix
Public Abid_DEngine_BanglaDate_monthNum_en As String
Public Abid_DEngine_BanglaDate_yearNum_ar As String

Public Abid_DEngine_EnglishDate_dayNum_bn As String
Public Abid_DEngine_EnglishDate_monthNum_en As String
Public Abid_DEngine_EnglishDate_yearNum_ar As String

Public Abid_DEngine_ArabicDate_dayNum_bn As String
Public Abid_DEngine_ArabicDate_monthNum_en As String
Public Abid_DEngine_ArabicDate_yearNum_ar As String

' [2] Week Names
Public Abid_DEngine_BanglaDate_weekName_bn As String
Public Abid_DEngine_BanglaDate_weekName_en As String
Public Abid_DEngine_BanglaDate_weekName_ar As String

Public Abid_DEngine_EnglishDate_weekName_bn As String
Public Abid_DEngine_EnglishDate_weekName_en As String
Public Abid_DEngine_EnglishDate_weekName_ar As String

Public Abid_DEngine_ArabicDate_weekName_bn As String
Public Abid_DEngine_ArabicDate_weekName_en As String
Public Abid_DEngine_ArabicDate_weekName_ar As String

' [3] Month Names
Public Abid_DEngine_BanglaDate_monthName_bn As String
Public Abid_DEngine_BanglaDate_monthName_en As String
Public Abid_DEngine_BanglaDate_monthName_ar As String

Public Abid_DEngine_EnglishDate_monthName_bn As String
Public Abid_DEngine_EnglishDate_monthName_en As String
Public Abid_DEngine_EnglishDate_monthName_ar As String

Public Abid_DEngine_ArabicDate_monthName_bn As String
Public Abid_DEngine_ArabicDate_monthName_en As String
Public Abid_DEngine_ArabicDate_monthName_ar As String

' [4] Eras
Public Abid_DEngine_BanglaDate_Era_bn As String
Public Abid_DEngine_BanglaDate_Era_en As String
Public Abid_DEngine_BanglaDate_Era_ar As String

Public Abid_DEngine_EnglishDate_Era_bn As String
Public Abid_DEngine_EnglishDate_Era_en As String
Public Abid_DEngine_EnglishDate_Era_ar As String

Public Abid_DEngine_ArabicDate_Era_bn As String
Public Abid_DEngine_ArabicDate_Era_en As String
Public Abid_DEngine_ArabicDate_Era_ar As String


Public Sub ABID__DateENGINE(ByVal InputDateStr As String)
    Dim dt As Date
    Dim parts() As String
    Dim isLeapYear As Boolean
    
    ' Split the DD/MM/YYYY format
    parts = Split(InputDateStr, "/")
    If UBound(parts) <> 2 Then
        MsgBox "Date Engine Error: Please provide date in DD/MM/YYYY format.", vbCritical
        Exit Sub
    End If
    
    ' Build Date object
    dt = DateSerial(CInt(parts(2)), CInt(parts(1)), CInt(parts(0)))
    
    ' =======================================================
    ' 0. LOAD RESOURCES FROM FILE
    ' =======================================================
    Dim arrBMonths() As String, arrGMonths() As String, arrHMonths() As String
    Dim arrEras() As String, arrWeeks() As String, arrNums() As String
    Dim arrBPrefixes() As String ' <--- NEW: Array to hold Bangla prefixes
    
    If Not LoadDateResources(arrBMonths, arrGMonths, arrHMonths, arrEras, arrWeeks, arrNums, arrBPrefixes) Then
        Exit Sub ' Stop if file cannot be loaded
    End If

    ' =======================================================
    ' 1. ENGLISH (GREGORIAN) DATE CALCULATION
    ' =======================================================
    Dim gDay As Integer, gMonth As Integer, gYear As Integer
    gDay = Day(dt)
    gMonth = Month(dt)
    gYear = Year(dt)

    ' =======================================================
    ' 2. ARABIC (HIJRI) DATE CALCULATION
    ' =======================================================
    Dim hDay As Integer, hMonth As Integer, hYear As Integer
    
    ' Switch VBA's internal calendar to Hijri
    VBA.Calendar = vbCalHijri
    hDay = Day(dt)
    hMonth = Month(dt)
    hYear = Year(dt)
    ' Switch back to Gregorian immediately
    VBA.Calendar = vbCalGreg

    ' =======================================================
    ' 3. BANGLA DATE CALCULATION (Bangladesh Standard)
    ' =======================================================
    Dim bDay As Integer, bMonth As Integer, bYear As Integer
    
    ' Gregorian Leap Year Check
    isLeapYear = (gYear Mod 4 = 0 And gYear Mod 100 <> 0) Or (gYear Mod 400 = 0)
    
    ' Calculate Bangla Year
    If gMonth < 4 Or (gMonth = 4 And gDay < 14) Then
        bYear = gYear - 594
    Else
        bYear = gYear - 593
    End If
    
    ' BD Government fixed constraints
    If gMonth = 4 And gDay >= 14 Then:     bMonth = 1: bDay = gDay - 13
    If gMonth = 5 And gDay <= 14 Then:     bMonth = 1: bDay = gDay + 17
    If gMonth = 5 And gDay >= 15 Then:     bMonth = 2: bDay = gDay - 14
    If gMonth = 6 And gDay <= 14 Then:     bMonth = 2: bDay = gDay + 17
    If gMonth = 6 And gDay >= 15 Then:     bMonth = 3: bDay = gDay - 14
    If gMonth = 7 And gDay <= 15 Then:     bMonth = 3: bDay = gDay + 16
    If gMonth = 7 And gDay >= 16 Then:     bMonth = 4: bDay = gDay - 15
    If gMonth = 8 And gDay <= 15 Then:     bMonth = 4: bDay = gDay + 16
    If gMonth = 8 And gDay >= 16 Then:     bMonth = 5: bDay = gDay - 15
    If gMonth = 9 And gDay <= 15 Then:     bMonth = 5: bDay = gDay + 16
    If gMonth = 9 And gDay >= 16 Then:     bMonth = 6: bDay = gDay - 15
    If gMonth = 10 And gDay <= 16 Then:    bMonth = 6: bDay = gDay + 15
    If gMonth = 10 And gDay >= 17 Then:    bMonth = 7: bDay = gDay - 16
    If gMonth = 11 And gDay <= 15 Then:    bMonth = 7: bDay = gDay + 15
    If gMonth = 11 And gDay >= 16 Then:    bMonth = 8: bDay = gDay - 15
    If gMonth = 12 And gDay <= 15 Then:    bMonth = 8: bDay = gDay + 15
    If gMonth = 12 And gDay >= 16 Then:    bMonth = 9: bDay = gDay - 15
    If gMonth = 1 And gDay <= 14 Then:     bMonth = 9: bDay = gDay + 16
    If gMonth = 1 And gDay >= 15 Then:     bMonth = 10: bDay = gDay - 14
    If gMonth = 2 And gDay <= 13 Then:     bMonth = 10: bDay = gDay + 17
    If gMonth = 2 And gDay >= 14 Then:     bMonth = 11: bDay = gDay - 13
    If gMonth = 3 And gDay <= 14 Then
        If isLeapYear Then
            bMonth = 11: bDay = gDay + 16
        Else
            bMonth = 11: bDay = gDay + 15
        End If
    End If
    If gMonth = 3 And gDay >= 15 Then:     bMonth = 12: bDay = gDay - 14
    If gMonth = 4 And gDay <= 13 Then:     bMonth = 12: bDay = gDay + 17

    ' =======================================================
    ' 4. POPULATE GLOBAL VARIABLES
    ' =======================================================
    Dim wDay As Integer
    wDay = Weekday(dt, vbSunday)

    ' --- NUMBERS --- (1=BN, 2=EN, 3=AR)
    Abid_DEngine_BanglaDate_dayNum_bn = ToLocalizedNum(CStr(bDay), 1, arrNums)
    Abid_DEngine_BanglaDate_monthNum_en = ToLocalizedNum(CStr(bMonth), 2, arrNums)
    Abid_DEngine_BanglaDate_yearNum_ar = ToLocalizedNum(CStr(bYear), 3, arrNums)
    
    ' <--- NEW: BANGLA DATE SUFFIX LOGIC (লা, রা, ঠা, ই, শে) --->
    Dim bSuffix As String
    Select Case bDay
        Case 1
            bSuffix = arrBPrefixes(1) ' লা
        Case 2, 3
            bSuffix = arrBPrefixes(2) ' রা
        Case 4
            bSuffix = arrBPrefixes(3) ' ঠা
        Case 5 To 18
            bSuffix = arrBPrefixes(4) ' ই
        Case 19 To 31
            bSuffix = arrBPrefixes(5) ' শে
        Case Else
            bSuffix = ""
    End Select
    Abid_DEngine_BanglaDate_dayNum_bn_pre = Abid_DEngine_BanglaDate_dayNum_bn & bSuffix

    Abid_DEngine_EnglishDate_dayNum_bn = ToLocalizedNum(CStr(gDay), 1, arrNums)
    Abid_DEngine_EnglishDate_monthNum_en = ToLocalizedNum(CStr(gMonth), 2, arrNums)
    Abid_DEngine_EnglishDate_yearNum_ar = ToLocalizedNum(CStr(gYear), 3, arrNums)

    Abid_DEngine_ArabicDate_dayNum_bn = ToLocalizedNum(CStr(hDay), 1, arrNums)
    Abid_DEngine_ArabicDate_monthNum_en = ToLocalizedNum(CStr(hMonth), 2, arrNums)
    Abid_DEngine_ArabicDate_yearNum_ar = ToLocalizedNum(CStr(hYear), 3, arrNums)

    ' --- WEEK NAMES ---
    Abid_DEngine_BanglaDate_weekName_bn = arrWeeks(wDay, 1)
    Abid_DEngine_BanglaDate_weekName_en = arrWeeks(wDay, 2)
    Abid_DEngine_BanglaDate_weekName_ar = arrWeeks(wDay, 3)

    Abid_DEngine_EnglishDate_weekName_bn = arrWeeks(wDay, 1)
    Abid_DEngine_EnglishDate_weekName_en = arrWeeks(wDay, 2)
    Abid_DEngine_EnglishDate_weekName_ar = arrWeeks(wDay, 3)

    Abid_DEngine_ArabicDate_weekName_bn = arrWeeks(wDay, 1)
    Abid_DEngine_ArabicDate_weekName_en = arrWeeks(wDay, 2)
    Abid_DEngine_ArabicDate_weekName_ar = arrWeeks(wDay, 3)

    ' --- MONTH NAMES ---
    Abid_DEngine_BanglaDate_monthName_bn = arrBMonths(bMonth, 1)
    Abid_DEngine_BanglaDate_monthName_en = arrBMonths(bMonth, 2)
    Abid_DEngine_BanglaDate_monthName_ar = arrBMonths(bMonth, 3)

    Abid_DEngine_EnglishDate_monthName_bn = arrGMonths(gMonth, 1)
    Abid_DEngine_EnglishDate_monthName_en = arrGMonths(gMonth, 2)
    Abid_DEngine_EnglishDate_monthName_ar = arrGMonths(gMonth, 3)

    Abid_DEngine_ArabicDate_monthName_bn = arrHMonths(hMonth, 1)
    Abid_DEngine_ArabicDate_monthName_en = arrHMonths(hMonth, 2)
    Abid_DEngine_ArabicDate_monthName_ar = arrHMonths(hMonth, 3)

    ' --- ERAS ---
    ' Index 1: Bangabda, Index 2: AD, Index 3: Hijri
    Abid_DEngine_BanglaDate_Era_bn = arrEras(1, 1)
    Abid_DEngine_BanglaDate_Era_en = arrEras(1, 2)
    Abid_DEngine_BanglaDate_Era_ar = arrEras(1, 3)

    Abid_DEngine_EnglishDate_Era_bn = arrEras(2, 1)
    Abid_DEngine_EnglishDate_Era_en = arrEras(2, 2)
    Abid_DEngine_EnglishDate_Era_ar = arrEras(2, 3)

    Abid_DEngine_ArabicDate_Era_bn = arrEras(3, 1)
    Abid_DEngine_ArabicDate_Era_en = arrEras(3, 2)
    Abid_DEngine_ArabicDate_Era_ar = arrEras(3, 3)
    
End Sub

' =======================================================
' HELPER FUNCTIONS
' =======================================================

' <--- NEW: Added arrBPrefixes to method parameters
Private Function LoadDateResources(ByRef arrBMonths() As String, ByRef arrGMonths() As String, ByRef arrHMonths() As String, _
                                   ByRef arrEras() As String, ByRef arrWeeks() As String, ByRef arrNums() As String, _
                                   ByRef arrBPrefixes() As String) As Boolean
    Dim filePath As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Build path logic targeting %appdata%
    filePath = Environ("AppData") & "\Abid\ABID_AL_JAMEE\MS_Office\MS_Word\VBA\VBA_need\ABID_DateENGINE_resources.txt"
    
    ' Fallback to check if it was saved with a .txt extension
    If Not fso.FileExists(filePath) Then
        If fso.FileExists(filePath & ".txt") Then
            filePath = filePath & ".txt"
        Else
            MsgBox "Resource file not found at: " & filePath, vbCritical
            LoadDateResources = False
            Exit Function
        End If
    End If
    
    ' Late binding ADODB Stream for UTF-8 reading
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    
    ' Normalize line endings and split
    Dim lines() As String
    lines = Split(Replace(stream.ReadText, vbCrLf, vbLf), vbLf)
    stream.Close
    
    ReDim arrBMonths(1 To 12, 1 To 3)
    ReDim arrGMonths(1 To 12, 1 To 3)
    ReDim arrHMonths(1 To 12, 1 To 3)
    ReDim arrEras(1 To 3, 1 To 3)
    ReDim arrWeeks(1 To 7, 1 To 3)
    ReDim arrNums(0 To 9, 1 To 3)
    ReDim arrBPrefixes(1 To 5) ' <--- NEW: ReDim for Bangla prefixes
    
    Dim currentSection As String
    Dim bMonthIdx As Integer: bMonthIdx = 1
    Dim gMonthIdx As Integer: gMonthIdx = 1
    Dim hMonthIdx As Integer: hMonthIdx = 1
    Dim eraIdx As Integer: eraIdx = 1
    Dim weekIdx As Integer: weekIdx = 1
    Dim numIdx As Integer: numIdx = 0
    Dim i As Long
    Dim ln As String
    Dim parts() As String
    Dim pParts() As String ' <--- NEW: For comma separation
    
    For i = LBound(lines) To UBound(lines)
        ln = Trim(lines(i))
        If Left(ln, 1) = "[" And Right(ln, 1) = "]" Then
            currentSection = ln
        ElseIf ln <> "" Then
            ' <--- NEW: Logic to handle comma-separated prefix string
            If currentSection = "[BANG_DATE_NUM_PREFIX]" Then
                pParts = Split(ln, ",")
                If UBound(pParts) >= 4 Then
                    arrBPrefixes(1) = Trim(pParts(0))
                    arrBPrefixes(2) = Trim(pParts(1))
                    arrBPrefixes(3) = Trim(pParts(2))
                    arrBPrefixes(4) = Trim(pParts(3))
                    arrBPrefixes(5) = Trim(pParts(4))
                End If
            ElseIf InStr(ln, "|") > 0 Then ' --- Existing parsing continues
                parts = Split(ln, "|")
                If UBound(parts) >= 2 Then
                    Select Case currentSection
                        Case "[BANG_MONTHS]"
                            If bMonthIdx <= 12 Then
                                arrBMonths(bMonthIdx, 1) = Trim(parts(0)): arrBMonths(bMonthIdx, 2) = Trim(parts(1)): arrBMonths(bMonthIdx, 3) = Trim(parts(2))
                                bMonthIdx = bMonthIdx + 1
                            End If
                        Case "[GREGORIAN_MONTHS]"
                            If gMonthIdx <= 12 Then
                                arrGMonths(gMonthIdx, 1) = Trim(parts(0)): arrGMonths(gMonthIdx, 2) = Trim(parts(1)): arrGMonths(gMonthIdx, 3) = Trim(parts(2))
                                gMonthIdx = gMonthIdx + 1
                            End If
                        Case "[HIJRI_MONTHS]"
                            If hMonthIdx <= 12 Then
                                arrHMonths(hMonthIdx, 1) = Trim(parts(0)): arrHMonths(hMonthIdx, 2) = Trim(parts(1)): arrHMonths(hMonthIdx, 3) = Trim(parts(2))
                                hMonthIdx = hMonthIdx + 1
                            End If
                        Case "[ERAS]"
                            If eraIdx <= 3 Then
                                arrEras(eraIdx, 1) = Trim(parts(0)): arrEras(eraIdx, 2) = Trim(parts(1)): arrEras(eraIdx, 3) = Trim(parts(2))
                                eraIdx = eraIdx + 1
                            End If
                        Case "[WEEKDAYS]"
                            If weekIdx <= 7 Then
                                arrWeeks(weekIdx, 1) = Trim(parts(0)): arrWeeks(weekIdx, 2) = Trim(parts(1)): arrWeeks(weekIdx, 3) = Trim(parts(2))
                                weekIdx = weekIdx + 1
                            End If
                        Case "[NUMERALS]"
                            If numIdx <= 9 Then
                                arrNums(numIdx, 1) = Trim(parts(0)): arrNums(numIdx, 2) = Trim(parts(1)): arrNums(numIdx, 3) = Trim(parts(2))
                                numIdx = numIdx + 1
                            End If
                    End Select
                End If
            End If
        End If
    Next i
    
    LoadDateResources = True
End Function

Private Function ToLocalizedNum(ByVal strNum As String, ByVal langIdx As Integer, ByRef arrNums() As String) As String
    Dim i As Integer
    ' Loop through 0-9 using the English numerals (index 2) as the find string
    For i = 0 To 9
        strNum = Replace(strNum, arrNums(i, 2), arrNums(i, langIdx))
    Next i
    ToLocalizedNum = strNum
End Function
