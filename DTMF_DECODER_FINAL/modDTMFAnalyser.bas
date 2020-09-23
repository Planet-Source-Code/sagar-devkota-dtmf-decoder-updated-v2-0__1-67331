Attribute VB_Name = "modDTMFAnalyser"
Public SAMPLING_RATE As Long, MAX_BINS As Integer, GOERTZEL_N As Integer, q1(0 To 7) As Double, q2(0 To 7) As Double, samples() As Integer, freqs(0 To 7) As Integer, coefs(0 To 7) As Double, r(0 To 7) As Double
Public sample_count As Integer, found As Boolean
Public see_digit As Boolean, dgt As String, exec As Boolean
Public Const FFT_SAMPLES            As Long = 1024


Public Sub goertzel(sample As Integer, tx As String)
    Dim q0 As Double
    Dim i As Integer, n As Integer

  For n = 0 To MAX_BINS - 1
      coefs(n) = 2# * Cos(2# * 3.141592654 * freqs(n) / SAMPLING_RATE)
  Next n
  If (sample_count < GOERTZEL_N) Then
    sample_count = sample_count + 1
    
    For i = 0 To MAX_BINS - 1
      q0 = coefs(i) * q1(i) - q2(i) + sample
      q2(i) = q1(i)
      q1(i) = q0
    Next i
  
  Else
   For i = 0 To MAX_BINS - 1
      r(i) = (q1(i) * q1(i)) + (q2(i) * q2(i)) - (coefs(i) * q1(i) * q2(i))
      
      
      q1(i) = 0#
      q2(i) = 0#
    Next i
      Dim row As Integer, col As Integer
  
  Dim peak_count As Integer, max_index As Integer
  Dim maxval As Double, t As Double
  Dim row_col_ascii_codes(0 To 3, 0 To 3) As String
  
    
row_col_ascii_codes(0, 0) = "1"
row_col_ascii_codes(0, 1) = "2"
row_col_ascii_codes(0, 2) = "3"
row_col_ascii_codes(0, 3) = "A"
row_col_ascii_codes(1, 0) = "4"
row_col_ascii_codes(1, 1) = "5"
row_col_ascii_codes(1, 2) = "6"
row_col_ascii_codes(1, 3) = "B"
row_col_ascii_codes(2, 0) = "7"
row_col_ascii_codes(2, 1) = "8"
row_col_ascii_codes(2, 2) = "9"
row_col_ascii_codes(2, 3) = "C"
row_col_ascii_codes(3, 0) = "*"
row_col_ascii_codes(3, 1) = "0"
row_col_ascii_codes(3, 2) = "#"
row_col_ascii_codes(3, 3) = "D"





'  /* Find the largest in the row group. */
  row = 0
  maxval = 0#
  For i = 0 To 3
  
    If (r(i) > maxval) Then
          maxval = r(i)
      row = i
    End If
  Next i
  

  '/* Find the largest in the column group. */
  col = 4
  maxval = 0#

  For i = 4 To 7
  
    If (r(i) > maxval) Then
          maxval = r(i)
      col = i
    End If
  Next i

  '/* Check for minimum energy */

  If (r(row) < 400000#) Then
  
  ' /* 2.0e5 ... 1.0e8 no change */
  
      
  ElseIf (r(col) < 400000#) Then
  
   ' /* energy not high enough */
 Else
  
    see_digit = True
    

   
    If (r(col) > r(row)) Then
     
      ' {     /* Normal twist */
     max_index = col
       If (r(row) < (r(col) * 0.398)) Then see_digit = False
    
    Else '/* if ( r[row] > r[col] ) */
    
      '/* Reverse twist */
      max_index = row
      If (r(col) < (r(row) * 0.158)) Then see_digit = False

    
     End If
     
    
    If (r(max_index) > 1000000000#) Then
      t = r(max_index) * 0.158
    Else
      t = r(max_index) * 0.01
    End If
    

    peak_count = 0
    For i = 0 To 7
        If (r(i) > t) Then
        peak_count = peak_count + 1
        End If
    Next i
    If (peak_count > 2) Then see_digit = False
      

    If (see_digit) Then
    'tx.Text = ""
    
     ' tx.Text = tx.Text & vbCrLf & row_col_ascii_codes(row, (col - 4)) & " Detected"
    tx = tx & row_col_ascii_codes(row, (col - 4))
   
    
      'MsgBox row_col_ascii_codes(row, (col - 4)) & " Detected"
     'found = True
      End If
   End If
    
   
    sample_count = 0
    
 End If
End Sub




