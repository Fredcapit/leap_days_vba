Attribute VB_Name = "leap_days_functions"
Const d_min As String = "01.03.1900"
Const d_max As String = "01.01.2900"


Public Function LEAP_DAYS(ByVal val_begin As Long, ByVal val_end As Long, Optional count_first_day = 0, Optional count_last_day = 1) As Long
    
    Dim d_begin, d_end As Date
    
    count_first_day = IIf(count_first_day <> 0, 1, 0)
    count_last_day = IIf(count_last_day <> 0, 1, 0)
    
    
    d_begin = CDate(val_begin)
    d_end = CDate(val_end)
    
    Dim check_error As Variant
    check_error = check_constrains(d_begin, d_end)
    
    If IsError(check_error) Then
        LEAP_DAYS = check_error
        Exit Function
    End If
    
    Dim result As Long
    result = 0
    
    If is_year_leap(d_begin) Then
        result = DateSerial(year(d_begin), 12, 31) - d_begin
    End If
    
    If (year(d_end) - year(d_begin)) > 1 Then
        result = result + (DateSerial(year(d_end), 1, 1) - DateSerial(year(d_begin) + 1, 1, 1) - (year(d_end) - year(d_begin) - 1) * 365) * 366
    End If
    
    If is_year_leap(d_end) Then
        result = result + (d_end - DateSerial(year(d_end), 1, 1) + 1)
    End If
    
    If is_year_leap(d_begin) And count_first_day = 1 Then result = result + 1
    If is_year_leap(d_end) And count_last_day = 0 Then result = result - 1
    
    
    LEAP_DAYS = result
End Function

Public Function NON_LEAP_DAYS(ByVal val_begin As Long, ByVal val_end As Long, Optional count_first_day = 0, Optional count_last_day = 1) As Long
                
    Dim d_begin, d_end As Date
    
    count_first_day = IIf(count_first_day <> 0, 1, 0)
    count_last_day = IIf(count_last_day <> 0, 1, 0)
    
    
    d_begin = CDate(val_begin)
    d_end = CDate(val_end)
    
    Dim check_error As Variant
    check_error = check_constrains(d_begin, d_end)
    
    If IsError(check_error) Then
        NON_LEAP_DAYS = check_error
        Exit Function
    End If
    
    Dim result As Long
    result = 0
    
    If Not is_year_leap(d_begin) Then
        result = DateSerial(year(d_begin), 12, 31) - d_begin
    End If
    
    If (year(d_end) - year(d_begin)) > 1 Then
        result = result + (DateSerial(year(d_end), 1, 1) - DateSerial(year(d_begin) + 1, 1, 1)) - (DateSerial(year(d_end), 1, 1) - DateSerial(year(d_begin) + 1, 1, 1) - (year(d_end) - year(d_begin) - 1) * 365) * 366
    End If
    
    If Not is_year_leap(d_end) Then
        result = result + (d_end - DateSerial(year(d_end), 1, 1) + 1)
    End If
    
    If Not is_year_leap(d_begin) And count_first_day = 1 Then result = result + 1
    If Not is_year_leap(d_end) And count_last_day = 0 Then result = result - 1
    
    
    NON_LEAP_DAYS = result
End Function

Private Function check_constrains(ByVal d_begin As Date, ByVal d_end As Date) As Variant
    
    check_constrains = Null
    
    If Not ((CDate(d_min) <= d_begin) And (d_begin <= d_end)) _
        Or Not ((CDate(d_max) >= d_end) And (d_end >= d_begin)) Then
        check_constrains = CVErr(xlErrNum)
    End If
    
End Function
                        
Private Function is_year_leap(ByVal d_date As Date) As Boolean

    Dim int_year  As Integer
    int_year = year(d_date)
    
    
    Dim mod_4, mod_100, mod_400 As Integer
    
    mod_4 = int_year Mod 4
    mod_100 = int_year Mod 100
    mod_400 = int_year Mod 400
        
            
    Dim result As Boolean
    
    result = False
    result = result Or (mod_400 = 0) Or (mod_4 = 0 And mod_100 > 0)
        
    
    is_year_leap = result
    
End Function
