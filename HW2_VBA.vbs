Sub stock_data()

         ' Declare Current as a worksheet object variable.
         Dim ws As Worksheet

         ' Loop through all of the worksheets in the active workbook.
         For Each ws In Worksheets

            'Column headers for output data columns I:L
            ws.Range("I1") = "Ticker"
            ws.Range("J1") = "Yearly Change"
            ws.Range("K1") = "Percent Change"
            ws.Range("L1") = "Total Stock Volume"
            
            'Define counters for output in columns I:L
            Dim j As Integer
            j = 1
            
            Dim k As Integer
            k = 1
              
            'Return row value of last non-empty cell column A
            Dim lRow As Long
            lRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
            For I = 2 To lRow
                
                vol = vol + ws.Cells(I, 7).Value
                
                If ws.Cells(I, 1).Value <> ws.Cells(I - 1, 1).Value Then
                    Dim OP As Double
                    OP = ws.Cells(I, 3).Value
                    k = k + 1
                    Dim Tic As String
                    Tic = ws.Cells(I, 1).Value
                    ws.Cells(k, 9).Value = Tic
                End If
                
                If ws.Cells(I, 1).Value <> ws.Cells(I + 1, 1).Value Then
                    Dim CL As Double
                    CL = ws.Cells(I, 6).Value
                    Dim YC As Double
                    YC = CL - OP
                    Dim PC As Double
                    If OP <> 0 Then
                        PC = (CL / OP) - 1
                    Else: PC = 0
                    End If
                    j = j + 1
                    ws.Cells(j, 10).Value = YC
                        If ws.Cells(j, 10).Value > 0 Then
                            ws.Cells(j, 10).Interior.Color = RGB(0, 200, 0)
                        ElseIf ws.Cells(j, 10).Value <= 0 Then
                            ws.Cells(j, 10).Interior.Color = RGB(200, 40, 40)
                        End If
                    ws.Cells(j, 11).Value = PC
                    ws.Cells(j, 11).NumberFormat = "0.00%"
                    ws.Cells(j, 12).Value = vol
                    vol = 0
                    
                End If
                
            Next I
            
            ws.Range("I:L").Columns.AutoFit
            ws.Range("P1") = "Ticker"
            ws.Range("Q1") = "Value"
            ws.Range("O2") = "Greatest % Increase"
            ws.Range("O3") = "Greatest % Decrease"
            ws.Range("O4") = "Greatest Total Volume"
            
            Dim lRow2 As Long
            lRow2 = ws.Cells(Rows.Count, 11).End(xlUp).Row
            
            Dim g_inc As Double
            Dim g_dec As Double
            Dim g_tv As Double
            
            Dim gi_tic As String
            Dim gd_tic As String
            Dim gtv_tic As String
            
            g_inc = ws.Cells(2, 11).Value
            g_dec = ws.Cells(2, 11).Value
            g_tv = ws.Cells(2, 12).Value
            
            For I = 3 To lRow2
                
                
                If ws.Cells(I, 11).Value > g_inc Then
                    g_inc = ws.Cells(I, 11).Value
                    gi_tic = ws.Cells(I, 9).Value
                End If
                
                If ws.Cells(I, 11).Value < g_dec Then
                    g_dec = ws.Cells(I, 11).Value
                    gd_tic = ws.Cells(I, 9).Value
                End If
                    
                If ws.Cells(I, 12).Value > g_tv Then
                    g_tv = ws.Cells(I, 12).Value
                    gtv_tic = ws.Cells(I, 9).Value
                End If
                
            Next I
            
            ws.Range("P2") = gi_tic
            ws.Range("Q2") = g_inc
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("P3") = gd_tic
            ws.Range("Q3") = g_dec
            ws.Range("Q3").NumberFormat = "0.00%"
            ws.Range("P4") = gtv_tic
            ws.Range("Q4") = g_tv
            ws.Range("O:Q").Columns.AutoFit
            ' This line displays the worksheet name in a message box.
            MsgBox (ws.Name)
         Next

      End Sub


