Attribute VB_Name = "Module1"
Sub creditcharges()
Dim i As Long
Dim j As Long
Dim k As Long
Dim total As Double
Dim ws As Worksheet
Dim LastRow As Long

        For Each ws In Worksheets
                    k = 1
                    j = 1
                    total = 0
                    LastRow = 0
                    gr_incr_t = "i"
                    gr_dec_t = "d"
                    gr_val_t = "v"
                    gr_incr = 0
                    gr_dec = 0
                    gr_val = 0
                    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                    ws.Cells(1, 1) = LastRow
                    For i = 2 To LastRow
                                 If ws.Cells(i - 1, 1) <> ws.Cells(i, 1) Then
                                       first_open = ws.Cells(i, 3)
                                End If
                                If ws.Cells(i, 1) = ws.Cells(i + 1, 1) Then
                                       total = total + ws.Cells(i, 7)
                                Else:
                                       total = total + ws.Cells(i, 7)
                                       last_close = ws.Cells(i, 6)
                                       ws.Cells(k + 1, 11) = total
                                       k = k + 1
                                       ws.Cells(k, 10) = ws.Cells(i, 1)
                                       ws.Cells(k, 12) = last_close - first_open 'last close-first open
                                   'if year opening is 0, then division by 0 problem
                                     If first_open = 0 Then
                                         ws.Cells(k, 13) = 0
                                    Else
                                        ws.Cells(k, 13) = FormatPercent(last_close / first_open - 1, 2)
                                    End If
                                    
                                    total = 0
                                End If
                    Next i
                  For j = 2 To k  'Calculation of max and min decrease and max total
                                       If ws.Cells(j, 11) >= gr_val Then
                                            gr_val = ws.Cells(j, 11)
                                            gr_val_t = ws.Cells(j, 10)
                                        End If
                                        If ws.Cells(j, 13) >= gr_incr Then
                                            gr_incr = ws.Cells(j, 13)
                                            gr_incr_t = ws.Cells(j, 10)
                                        End If
                                        If ws.Cells(j, 13) <= gr_dec Then
                                            gr_dec = ws.Cells(j, 13)
                                            gr_dec_t = ws.Cells(j, 10)
                                     
                                        End If
                Next j
            ws.Cells(2, 16) = gr_incr_t
            ws.Cells(3, 16) = gr_dec_t
            ws.Cells(4, 16) = gr_val_t
            ws.Cells(2, 17) = FormatPercent(gr_incr, 2)
            ws.Cells(3, 17) = FormatPercent(gr_dec, 2)
            ws.Cells(4, 17) = gr_val
        
        Next ws

End Sub
