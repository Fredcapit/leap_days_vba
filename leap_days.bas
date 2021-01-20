Attribute VB_Name = "leap_days"
Const d_min As String = "01.03.1900"
Const d_max As String = "31.12.2899"


Public Function ÂÈÑÎÊÎÑÍÛÕ_ÄÍÅÉ(ByVal val_begin As Long, ByVal val_end As Long) As Long
    
    Dim d_begin, d_end As Date
    
    d_begin = CDate(val_begin)
    d_end = CDate(val_end)
    
    Dim check_error As Variant
    check_error = check_constrains(d_begin, d_end)
    
    If IsError(check_error) Then
        ÂÈÑÎÊÎÑÍÛÕ_ÄÍÅÉ = check_error
        Exit Function
    End If
    
    Dim result As Long
    result = 0
    
    result = first_quartet_leap_year_days(d_begin, d_end) + middle_quartets_leap_year_days(d_begin, d_end) + last_quartet_leap_year_days(d_begin, d_end)
    
    ÂÈÑÎÊÎÑÍÛÕ_ÄÍÅÉ = result
    
End Function


Public Function ÍÅÂÈÑÎÊÎÑÍÛÕ_ÄÍÅÉ(ByVal val_begin As Long, ByVal val_end As Long) As Long

    Dim d_begin, d_end As Date
    
    d_begin = CDate(val_begin)
    d_end = CDate(val_end)
    
    Dim check_error As Variant
    check_error = check_constrains(d_begin, d_end)
    
    If IsError(check_error) Then
        ÍÅÂÈÑÎÊÎÑÍÛÕ_ÄÍÅÉ = check_error
        Exit Function
    End If
        
    Dim result As Long
    result = 0
    
    result = first_quartet_nonleap_days(d_begin, d_end) + _
             middle_quartet_nonleap_year_days(d_begin, d_end) + _
             last_quartet_nonleap_year_days(d_begin, d_end)
    
    ÍÅÂÈÑÎÊÎÑÍÛÕ_ÄÍÅÉ = result
    
End Function

Private Function first_quartet_leap_year_days(ByVal d_begin As Date, ByVal d_end As Date) As Long

    Dim result As Long
    result = 0
    
    Dim year_diff As Long
    Dim quartet_index_diff As Long
    
    year_diff = year(d_end) - year(d_begin)
    quartet_index_diff = quartet_index(year(d_end)) - quartet_index(year(d_begin))
    
    If year_diff = 0 Then
        result = DateDiff("d", d_begin, d_end)
        first_quartet_leap_year_days = result
        Exit Function
    End If
    
    If quartet_index_diff = 0 Then
    
        If is_year_leap(d_begin) Then
            result = DateDiff("d", d_begin, CDate(DateSerial(year(d_begin) + 1, 1, 1)))
            first_quartet_leap_year_days = result
            Exit Function
        
        End If
        
        If is_year_leap(d_end) Then
            result = DateDiff("d", CDate(DateSerial(year(d_end), 1, 1)), d_end)
            first_quartet_leap_year_days = result
            Exit Function
        End If
        
    Else
    
        If is_year_leap(d_begin) Then
            result = DateDiff("d", d_begin, CDate(DateSerial(year(d_begin) + 1, 1, 1)))
            first_quartet_leap_year_days = result
            Exit Function
        Else
        
            If Not is_quartet_noleap(quartet_index(year(d_begin))) Then
                result = 366
                first_quartet_leap_year_days = result
                Exit Function
            End If
            
        End If
        
    End If

    first_quartet_leap_year_days = result
    
End Function

Private Function first_quartet_nonleap_days(ByVal d_begin As Date, ByVal d_end As Date) As Long

    Dim result As Long
    result = 0
    
    Dim year_diff As Long
    Dim quartet_index_diff As Long
    
    year_diff = year(d_end) - year(d_begin)
    quartet_index_diff = quartet_index(year(d_end)) - quartet_index(year(d_begin))
    
    If year_diff = 0 And Not is_year_leap(d_begin) Then
    
        result = DateDiff("d", d_begin, d_end)
        first_quartet_nonleap_days = result
        Exit Function
        
    End If
    
    If quartet_index_diff = 0 Then
        
        If Not is_year_leap(d_begin) Then
            result = result + DateDiff("d", d_begin, CDate(DateSerial(year(d_begin) + 1, 1, 1)))
        End If
        
        If year_diff > 1 Then
            result = result + 365 * (year_diff - 1)
        End If
        
        If Not is_year_leap(d_end) Then
            result = result + DateDiff("d", DateSerial(year(d_end), 1, 1), d_end)
        End If
        
    Else
        If Not is_year_leap(d_begin) Then
            result = result + DateDiff("d", d_begin, CDate(DateSerial(year(d_begin) + 1, 1, 1)))
        End If
        
        Dim q_index As Integer
        
        q_index = year_of_quartet_index(year(d_begin))
         
        If q_index < 3 Then
            result = result + 365 * (3 - q_index)
        End If
        
    End If
    
        
    first_quartet_nonleap_days = result
    
End Function

Private Function year_of_quartet_index(ByVal year As Integer) As Integer
    
    Dim result  As Integer
    
    result = (year Mod 4)
    If (result = 0) Then
        result = 4
    End If
        
    year_of_quartet_index = result
    
End Function
Private Function quartet_index(ByVal year As Integer) As Integer
    
    quartet_index = WorksheetFunction.RoundUp(year / 4, 0)
    
 End Function
Private Function is_quartet_noleap(ByVal quartet_index As Integer) As Boolean

    Dim mod_25, mod_100 As Integer
    
    Dim result As Boolean
    
    mod_25 = quartet_index Mod 25
    mod_100 = quartet_index Mod 100
    
    result = True And (mod_25 = 0 And mod_100 > 0)
    
    is_quartet_noleap = result
    
End Function

Private Function last_quartet_leap_year_days(ByVal d_begin As Date, ByVal d_end As Date) As Long
    
    Dim result As Long
    result = 0
     
    Dim quartet_index_diff As Long
       
    quartet_index_diff = quartet_index(year(d_end)) - quartet_index(year(d_begin))
    
    If quartet_index_diff > 0 Then
    
        If is_year_leap(d_end) Then
            result = DateDiff("d", CDate(DateSerial(year(d_end), 1, 1)), d_end)
        End If
        
    End If
    
     
    last_quartet_leap_year_days = result
    
End Function

Private Function last_quartet_nonleap_year_days(ByVal d_begin As Date, ByVal d_end As Date) As Long
    
    Dim result As Long
    result = 0
    
    Dim quartet_index_diff As Long
       
    quartet_index_diff = quartet_index(year(d_end)) - quartet_index(year(d_begin))
    
    If quartet_index_diff > 0 Then
         
        Dim year_of_quartet As Integer
         
        year_of_quartet = year_of_quartet_index(year(d_end))
        
        If year_of_quartet > 1 Then
            result = result + 365 * (year_of_quartet - 1)
        End If
        
        If Not is_year_leap(d_end) Then
            result = result + DateDiff("d", CDate(DateSerial(year(d_end), 1, 1)), d_end)
        End If
            
        
    End If
    
    last_quartet_nonleap_year_days = result
    
End Function

Private Function middle_quartets_leap_year_days(ByVal d_begin As Date, ByVal d_end As Date) As Long
    
    Dim quartet_count As Long
    
    quartet_count = middle_quartets_count(d_begin, d_end)
    
    If quartet_count = 0 Then
    
        middle_quartets_leap_year_days = 0
        Exit Function
        
    End If
    
    Dim q_begin, q_end As Long
    
    q_begin = quartet_index(year(d_begin))
    q_end = quartet_index(year(d_end)) - 1
    
    Dim quot_25, quot_100 As Integer
    
    quot_25 = WorksheetFunction.Quotient(q_end, 25) - WorksheetFunction.Quotient(q_begin, 25)
    quot_100 = WorksheetFunction.Quotient(q_end, 100) - WorksheetFunction.Quotient(q_begin, 100)
    
    Dim result As Long
    
    result = (quartet_count - quot_25 + quot_100) * 366
    
    middle_quartets_leap_year_days = result
        
End Function

Private Function middle_quartet_nonleap_year_days(ByVal d_begin As Date, ByVal d_end As Date) As Long
    
    Dim quartet_count As Long
    
    quartet_count = middle_quartets_count(d_begin, d_end)
    
    If quartet_count = 0 Then
        middle_quartet_nonleap_year_days = 0
        Exit Function
    End If
    
    Dim q_begin, q_end As Long
    
    q_begin = quartet_index(year(d_begin))
    q_end = quartet_index(year(d_end)) - 1
    
    Dim quot_25, quot_100 As Integer
    
    quot_25 = WorksheetFunction.Quotient(q_end, 25) - WorksheetFunction.Quotient(q_begin, 25)
    quot_100 = WorksheetFunction.Quotient(q_end, 100) - WorksheetFunction.Quotient(q_begin, 100)
       
    Dim result As Long
    
    result = (quartet_count * 3 + quot_25 - quot_100) * 365
    
    middle_quartet_nonleap_year_days = result
    
End Function
Private Function middle_quartets_count(ByVal d_begin As Date, ByVal d_end As Date) As Long
    
    Dim q_begin, q_end, q_Diff As Long
    
    q_begin = quartet_index(year(d_begin))
    q_end = quartet_index(year(d_end))
    
    Dim result As Long
    
    result = q_end - q_begin - 1
    
    If result < 1 Then
        result = 0
    End If
    
    middle_quartets_count = result
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
