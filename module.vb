Function toJalaali(gy As Integer, gm As Integer, gd As Integer)
  toJalaali = Join(d2j(g2d(gy, gm, gd)), "/")
End Function

Function d2jFormatted(jdn)
    d2jFormatted = Join(d2j(jdn), "/")
End Function

Function toGregorian(jy, jm, jd)
  toGregorian = Join(d2g(j2d(jy, jm, jd)), "/")
End Function

Function isValidJalaaliDate(jy As Integer, jm As Integer, jd As Integer)
  isValidJalaaliDate = jy >= -61 And jy <= 3177 And _
          jm >= 1 And jm <= 12 And _
          jd >= 1 And jd <= jalaaliMonthLength(jy, jm)
End Function

Function isLeapJalaaliYear(jy)
    If jalCalLeap(jy) = 0 Then
        isLeapJalaaliYear = True
    Else
        isLeapJalaaliYear = False
    End If
End Function

Function jalaaliMonthLength(jy, jm)
    If jm <= 6 Then
        jalaaliMonthLength = 31
    ElseIf jm <= 11 Then
        jalaaliMonthLength = 30
    ElseIf isLeapJalaaliYear(jy) Then
        jalaaliMonthLength = 30
    Else
        jalaaliMonthLength = 29
    End If
End Function

Function jalCalLeap(jy)
    Dim breaks
    breaks = Array(-61, 9, 38, 199, 426, 686, 756, 818, 1111, 1181, 1210 _
  , 1635, 2060, 2097, 2192, 2262, 2324, 2394, 2456, 3178)
  Dim bl, jp, jm, jump, leap, n, i
  b1 = myLen(breaks)
  jp = breaks(0)

  If jy < jp Or jy >= breaks(bl - 1) Then
    Err.Raise number:=vbObjectError + 513, _
              Description:="Invalid Jalaali year " + jy
  End If
    
  For i = 1 To b1 - 1
    jm = breaks(i)
    jump = jm - jp
    
    If jy < jm Then
      Exit For
    End If
    
    jp = jm
  Next i

  n = jy - jp
  
  If jump - n < 6 Then
    n = n - jump + myDiv(jump + 4, 33) * 33
  End If
  
  leap = myMod(myMod(n + 1, 33) - 1, 4)
  
  If (leap = -1) Then
    leap = 4
  End If
 
  jalCalLeap = leap
End Function

Function jalCal(jy, withoutLeap)
Dim breaks
    breaks = Array(-61, 9, 38, 199, 426, 686, 756, 818, 1111, 1181, 1210 _
  , 1635, 2060, 2097, 2192, 2262, 2324, 2394, 2456, 3178)
  Dim bl, gy, leapJ, jp, jm, jump, leap, leapG, march, n, i
  bl = myLen(breaks)
  gy = jy + 621
  leapJ = -14
  jp = breaks(0)

  If jy < jp Or jy >= breaks(bl - 1) Then
    Err.Raise number:=vbObjectError + 513, _
              Description:="Invalid Jalaali year " + jy
  End If

  For i = 1 To bl - 1
    jm = breaks(i)
    jump = jm - jp

    If jy < jm Then
      Exit For
    End If

    leapJ = leapJ + myDiv(jump, 33) * 8 + myDiv(myMod(jump, 33), 4)
    jp = jm
  Next i

  n = jy - jp

  leapJ = leapJ + myDiv(n, 33) * 8 + myDiv(myMod(n, 33) + 3, 4)
  
  If myMod(jump, 33) = 4 And jump - n = 4 Then
    leapJ = leapJ + 1
  End If

  leapG = myDiv(gy, 4) - myDiv((myDiv(gy, 100) + 1) * 3, 4) - 150

  march = 20 + leapJ - leapG

  If withoutLeap = True Then
    jalCal = Array(0, gy, march)
    
    Exit Function
  End If


  If jump - n < 6 Then
    n = n - jump + myDiv(jump + 4, 33) * 33
  End If

  leap = myMod(myMod(n + 1, 33) - 1, 4)

  If leap = -1 Then
    leap = 4
  End If

  jalCal = Array(leap, gy, march)
End Function

Function j2d(jy, jm, jd)
  Dim r
  r = jalCal(jy, True)

  j2d = g2d(r(1), 3, r(2)) + (jm - 1) * 31 - myDiv(jm, 7) * (jm - 7) + jd - 1
End Function

Function d2j(jdn)
  Dim gy, jy, r, jdn1f, jd, jm, k
  
  gy = d2g(jdn)(0)
  jy = gy - 621
  r = jalCal(jy, False)
  jdn1f = g2d(gy, 3, r(2))
  
  k = jdn - jdn1f
  If k >= 0 Then
    If k <= 185 Then
      jm = 1 + myDiv(k, 31)
      jd = myMod(k, 31) + 1
      d2j = Array(jy, jm, jd)
      
      Exit Function
    Else
      k = k - 186
    End If
  Else
    jy = jy - 1
    k = k + 179
    
    If r(0) = 1 Then
      k = k + 1
    End If
  End If

  jm = 7 + myDiv(k, 30)
  jd = myMod(k, 30) + 1
  
  d2j = Array(jy, jm, jd)
End Function

Function g2d(gy, gm, gd)
  Dim d
  d = myDiv((gy + myDiv(gm - 8, 6) + 100100) * 1461, 4) _
      + myDiv(153 * myMod(gm + 9, 12) + 2, 5) _
      + gd - 34840408
  
  d = d - myDiv(myDiv(gy + 100100 + myDiv(gm - 8, 6), 100) * 3, 4) + 752
  
  g2d = d
End Function

Function d2g(jdn)
  Dim j, i, gd, gm, gy
  j = 4 * jdn + 139361631
  j = j + myDiv(myDiv(4 * jdn + 183187720, 146097) * 3, 4) * 4 - 3908
  i = myDiv(myMod(j, 1461), 4) * 5 + 308
  gd = myDiv(myMod(i, 153), 5) + 1
  gm = myMod(myDiv(i, 153), 12) + 1
  gy = myDiv(j, 1461) - 100100 + myDiv(8 - gm, 6)
  
  d2g = Array(gy, gm, gd)
End Function

Function myDiv(a, b)
  myDiv = a \ b
End Function

Function myMod(a, b)
  myMod = a - myDiv(a, b) * b
End Function

Function myLen(arr)
    myLen = UBound(arr) - LBound(arr) + 1
End Function
