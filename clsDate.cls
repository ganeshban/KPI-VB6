VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'the original sorce code of this program is found in sourceforge.net
'the code in php was translated to VBA
'this add-in is copyfree and can be distributed freely in multiple copies
'the user should be responsible for the use of this add in
'the author is not liable for any damage that may occur by using this add-in
'modification to the code is welcome. Any additional data in the code if notified over the internet will be usefull to all
'***********************************************************************************************************
'*******currently can only calculate the date between BS 2000-2089 and AD 1944-2033*************************
'***********************************************************************************************************
'Coder: Nirmal Raj Joshi
'@copyfree nirmaljoshi.com.np,2010
'joshi.amn@gmail.com

Private bs As Variant
Public nep_date As Dictionary
Public eng_date As Dictionary
Public debug_info
Sub initilizeClass()
    ReDim bs(98)
    bs(0) = Array(2000, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31)
    bs(1) = Array(2001, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    bs(2) = Array(2002, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30)
    bs(3) = Array(2003, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)
    bs(4) = Array(2004, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31)
    bs(5) = Array(2005, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    bs(6) = Array(2006, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30)
    bs(7) = Array(2007, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)
    bs(8) = Array(2008, 31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 29, 31)
    bs(9) = Array(2009, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    bs(10) = Array(2010, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30)
    bs(11) = Array(2011, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)
    bs(12) = Array(2012, 31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 30, 30)
    bs(13) = Array(2013, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    bs(14) = Array(2014, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30)
    bs(15) = Array(2015, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)
    bs(16) = Array(2016, 31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 30, 30)
    bs(17) = Array(2017, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    bs(18) = Array(2018, 31, 32, 31, 32, 31, 30, 30, 29, 30, 29, 30, 30)
    bs(19) = Array(2019, 31, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31)
    bs(20) = Array(2020, 31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30)
    bs(21) = Array(2021, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    bs(22) = Array(2022, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 30)
    bs(23) = Array(2023, 31, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31)
    bs(24) = Array(2024, 31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30)
    bs(25) = Array(2025, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    bs(26) = Array(2026, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)
    bs(27) = Array(2027, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31)
    bs(28) = Array(2028, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    bs(29) = Array(2029, 31, 31, 32, 31, 32, 30, 30, 29, 30, 29, 30, 30)
    bs(30) = Array(2030, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)
    bs(31) = Array(2031, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31)
    bs(32) = Array(2032, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    bs(33) = Array(2033, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30)
    bs(34) = Array(2034, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)
    bs(35) = Array(2035, 30, 32, 31, 32, 31, 31, 29, 30, 30, 29, 29, 31)
    bs(36) = Array(2036, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    bs(37) = Array(2037, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30)
    bs(38) = Array(2038, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)
    bs(39) = Array(2039, 31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 30, 30)
    bs(40) = Array(2040, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    bs(41) = Array(2041, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30)
    bs(42) = Array(2042, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)
    bs(43) = Array(2043, 31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 30, 30)
    bs(44) = Array(2044, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    bs(45) = Array(2045, 31, 32, 31, 32, 31, 30, 30, 29, 30, 29, 30, 30)
    bs(46) = Array(2046, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)
    bs(47) = Array(2047, 31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30)
    bs(48) = Array(2048, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    bs(49) = Array(2049, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 30)
    bs(50) = Array(2050, 31, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31)
    bs(51) = Array(2051, 31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30)
    bs(52) = Array(2052, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    bs(53) = Array(2053, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 30)
    bs(54) = Array(2054, 31, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31)
    bs(55) = Array(2055, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    bs(56) = Array(2056, 31, 31, 32, 31, 32, 30, 30, 29, 30, 29, 30, 30)
    bs(57) = Array(2057, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)
    bs(58) = Array(2058, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31)
    bs(59) = Array(2059, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    bs(60) = Array(2060, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30)
    bs(61) = Array(2061, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)
    bs(62) = Array(2062, 30, 32, 31, 32, 31, 31, 29, 30, 29, 30, 29, 31)
    bs(63) = Array(2063, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    bs(64) = Array(2064, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30)
    bs(65) = Array(2065, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)
    bs(66) = Array(2066, 31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 29, 31)
    bs(67) = Array(2067, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    bs(68) = Array(2068, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30)
    bs(69) = Array(2069, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)
    bs(70) = Array(2070, 31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 30, 30)
    bs(71) = Array(2071, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    bs(72) = Array(2072, 31, 32, 31, 32, 31, 30, 30, 29, 30, 29, 30, 30)
    bs(73) = Array(2073, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31)
    bs(74) = Array(2074, 31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30)
    bs(75) = Array(2075, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    bs(76) = Array(2076, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 30)
    bs(77) = Array(2077, 31, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31)
    bs(78) = Array(2078, 31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30)
    bs(79) = Array(2079, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30)
    bs(80) = Array(2080, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 30)
    bs(81) = Array(2081, 31, 31, 32, 32, 31, 30, 30, 30, 29, 30, 30, 30)
    bs(82) = Array(2082, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 30, 30)
    bs(83) = Array(2083, 31, 31, 32, 31, 31, 30, 30, 30, 29, 30, 30, 30)
    bs(84) = Array(2084, 31, 31, 32, 31, 31, 30, 30, 30, 29, 30, 30, 30)
    bs(85) = Array(2085, 31, 32, 31, 32, 30, 31, 30, 30, 29, 30, 30, 30)
    bs(86) = Array(2086, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 30, 30)
    bs(87) = Array(2087, 31, 31, 32, 31, 31, 31, 30, 30, 29, 30, 30, 30)
    bs(88) = Array(2088, 30, 31, 32, 32, 30, 31, 30, 30, 29, 30, 30, 30)
    bs(89) = Array(2089, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 30, 30)
    bs(90) = Array(2090, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 30, 30)
    bs(91) = Array(2091, 30, 31, 31, 31, 31, 31, 30, 30, 29, 30, 30, 30)
    bs(92) = Array(2092, 30, 31, 30, 32, 29, 30, 30, 29, 29, 30, 30, 29)
    bs(93) = Array(2093, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 30, 30)
    bs(94) = Array(2094, 31, 30, 32, 30, 31, 30, 29, 30, 29, 30, 29, 30)
    bs(95) = Array(2095, 31, 31, 32, 31, 31, 31, 30, 29, 30, 30, 30, 30)

    bs(96) = Array(2096, 30, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30)
    bs(97) = Array(2097, 30, 32, 30, 32, 30, 30, 30, 30, 29, 30, 30, 30)
    bs(98) = Array(2098, 31, 31, 29, 31, 29, 31, 29, 29, 29, 30, 29, 30)

    Set nep_date = New Dictionary
        nep_date.Add "year", 0
        nep_date.Add "month", 0
        nep_date.Add "date", 0
        nep_date.Add "day", 0
        nep_date.Add "nmonth", 0
        nep_date.Add "num_day", 0
    Set eng_date = New Dictionary
        eng_date.Add "year", 0
        eng_date.Add "month", 0
        eng_date.Add "date", 0
        eng_date.Add "day", 0
        eng_date.Add "emonth", 0
        eng_date.Add "num_day", 0



End Sub

'        /**
'         * Calculates wheather english year is leap year or not
'         *
'         * @param integer $year
'         * @return boolean
'         */'
        Public Function is_leap_year(year) As Boolean
            Dim a, returnVal
            a = year
            If a Mod 100 = 0 Then

                If a Mod 400 = 0 Then
                    returnVal = True
                Else
                    returnVal = False
                End If


            Else
                If a Mod 4 = 0 Then
                   returnVal = True
                Else
                    returnVal = False
                End If
            End If
            is_leap_year = returnVal

        End Function

        Public Function get_nepali_month(m)
        Dim n_month
            n_month = False

            Select Case m
                Case 1:
                    n_month = "Baishak"
                Case 2:
                    n_month = "Jestha"

                Case 3:
                    n_month = "Ashad"
                Case 4:
                    n_month = "Shrawn"
                Case 5:
                    n_month = "Bhadra"
                Case 6:
                    n_month = "Ashwin"
                Case 7:
                    n_month = "Kartik"
                Case 8:
                    n_month = "Mangshir"
                Case 9:
                    n_month = "Poush"
                Case 10:
                    n_month = "Magh"
                Case 11:
                    n_month = "Falgun"
                Case 12:
                    n_month = "Chaitra"
                End Select


            get_nepali_month = n_month
        End Function

        Public Function get_english_month(m)
            Dim eMonth
            eMonth = False
            Select Case m
                Case 1
                    eMonth = "January"
                Case 2:
                    eMonth = "February"
                Case 3:
                    eMonth = "March"
                Case 4:
                    eMonth = "April"
                Case 5:
                    eMonth = "May"
                Case 6:
                    eMonth = "June"
                Case 7:
                    eMonth = "July"
                Case 8:
                    eMonth = "August"
                Case 9:
                    eMonth = "September"
                Case 10:
                    eMonth = "October"
                Case 11:
                    eMonth = "November"
                Case 12:
                    eMonth = "December"
                End Select

             get_english_month = eMonth

        End Function

        Public Function get_day_of_week(d)
        Dim day
        day = False
            Select Case d
                Case 1:
                    day = "Sunday"
                Case 2:
                    day = "Monday"
                Case 3:
                    day = "Tuesday"
                Case 4:
                    day = "Wednesday"
                Case 5:
                    day = "Thursday"
                Case 6:
                    day = "Friday"
                Case 7:
                    day = "Saturday"
                End Select
            get_day_of_week = day


        End Function
        Public Function is_range_eng(yy, mm, dd)
            Dim returnVal
            returnVal = True
            If (yy < 1944 Or yy > 2041) Then
                debug_info = "Supported only between 1944-2041"
                returnVal = False
            End If
            If (mm < 1 Or mm > 12) Then
                debug_info = "Error! value 1-12 only"
                returnVal = False
            End If
            If (dd < 1 Or dd > 31) Then
                debug_info = "Error! value 1-31 only"
                returnVal = False
            End If
            is_range_eng = returnVal
        End Function

        Public Function is_range_nep(yy, mm, dd)
            Dim returnVal
            returnVal = True

            If (yy < 2000 Or yy > 2098) Then
                debug_info = "Supported only between 2000-2098"
                returnVal = False
            End If
            If (mm < 1 Or mm > 12) Then
                debug_info = "Error! value 1-12 only"
                returnVal = False
            End If

            If (dd < 1 Or dd > 32) Then
                debug_info = "Error! value 1-31 only"
                returnVal = False
            End If
            is_range_nep = returnVal
        End Function

'                /**
'         * currently can only calculate the date between AD 1944-2033...
'         *
'         * @param unknown_type $yy
'         * @param unknown_type $mm
'         * @param unknown_type $dd
'         * @return unknown
'         */
        Public Sub eng_to_nep(yy, mm, dd)
            If (is_range_eng(yy, mm, dd) = False) Then
                debug_info = "Out of range"
            Else

'                // english month data.
            Dim month, lmonth
                month = Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
                lmonth = Array(31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)

            Dim def_eyy, def_nyy, def_nmm, def_ndd
            Dim total_eDays, total_nDays, a, day
            Dim m, Y, I, j
            Dim numDay
                def_eyy = 1944 '                                    //spear head english date...
                def_nyy = 2000: def_nmm = 9: def_ndd = 17 - 1 '     //spear head nepali date...
                total_eDays = 0: total_nDays = 0: a = 0: day = 7 - 1 '    //all the initializations...
                m = 0: Y = 0: I = 0: j = 0
                numDay = 0

                '// count total no. of days in-terms of year
                For I = 0 To yy - def_eyy - 1 ' //total days for month calculation...(english)
                    If (is_leap_year(def_eyy + I) = True) Then
                        For j = 0 To 12 - 1
                            total_eDays = total_eDays + lmonth(j)
                        Next j
                    Else
                        For j = 0 To 12 - 1
                            total_eDays = total_eDays + month(j)
                        Next j
                    End If
                Next I

                '// count total no. of days in-terms of month
                For I = 0 To (mm - 1) - 1
                    If (is_leap_year(yy) = True) Then
                        total_eDays = total_eDays + lmonth(I)
                    Else
                        total_eDays = total_eDays + month(I)
                    End If
                Next I

               ' // count total no. of days in-terms of date
                total_eDays = total_eDays + dd


                I = 0: j = def_nmm
                total_nDays = def_ndd
                m = def_nmm
                Y = def_nyy

                '// count nepali date from array
                Do While (total_eDays <> 0)
                    a = bs(I)(j)
                    total_nDays = total_nDays + 1 '     //count the days
                    day = day + 1 '                            //count the days interms of 7 days
                    If (total_nDays > a) Then
                        m = m + 1
                        total_nDays = 1
                        j = j + 1
                    End If
                    If (day > 7) Then
                        day = 1
                    End If
                    If (m > 12) Then
                        Y = Y + 1
                        m = 1
                    End If
                    If (j > 12) Then
                        j = 1: I = I + 1
                    End If
                    total_eDays = total_eDays - 1

                Loop
                numDay = day

                nep_date("year") = Y
                nep_date("month") = m
                nep_date("date") = total_nDays
                nep_date("day") = get_day_of_week(day)
                nep_date("nmonth") = get_nepali_month(m)
                nep_date("num_day") = numDay
            End If

        End Sub
'
'        /**
'         * currently can only calculate the date between BS 2000-2089
'         *
'         * @param unknown_type $yy
'         * @param unknown_type $mm
'         * @param unknown_type $dd
'         * @return unknown
'         */
        Public Sub nep_to_eng(yy, mm, dd)
            Dim def_eyy, def_emm, def_edd
            Dim def_nyy, def_nmm, def_ndd
            Dim total_eDays, total_nDays, a, day
            Dim m, Y, I, j
            Dim k, numDay
            Dim month, lmonth

            def_eyy = 1943: def_emm = 4: def_edd = 14 - 1   '// init english date.
            def_nyy = 2000: def_nmm = 1: def_ndd = 1 '        // equivalent nepali date.
            total_eDays = 0: total_nDays = 0: a = 0: day = 4 - 1 '     // initializations...
            m = 0: Y = 0: I = 0
            k = 0: numDay = 0

            month = Array(0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
            lmonth = Array(0, 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)

            If (is_range_nep(yy, mm, dd) = False) Then
                debug_info = "Out of range"
            Else

                '// count total days in-terms of year
                For I = 0 To yy - def_nyy - 1
                    For j = 1 To 12
                        total_nDays = total_nDays + bs(k)(j)
                    Next j
                    k = k + 1
                Next I


                '// count total days in-terms of month
                For j = 1 To mm - 1
                    total_nDays = total_nDays + bs(k)(j)
                Next j

                '// count total days in-terms of dat
                total_nDays = total_nDays + dd

                '//calculation of equivalent english date...
                total_eDays = def_edd
                m = def_emm
                Y = def_eyy
                Do While (total_nDays <> 0)
                    If (is_leap_year(Y)) Then

                        a = lmonth(m)

                    Else

                        a = month(m)
                    End If
                    total_eDays = total_eDays + 1
                    day = day + 1
                    If (total_eDays > a) Then
                        m = m + 1
                        total_eDays = 1
                        If (m > 12) Then
                            Y = Y + 1
                            m = 1
                        End If
                    End If
                    If (day > 7) Then day = 1
                    total_nDays = total_nDays - 1
                Loop
                numDay = day

                eng_date("year") = Y
                eng_date("month") = m
                eng_date("date") = total_eDays
                eng_date("day") = get_day_of_week(day)
                eng_date("emonth") = get_english_month(m)
                eng_date("num_day") = numDay

        End If

        End Sub




