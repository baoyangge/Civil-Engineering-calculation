
 ' The log 10 function
    Static Function Log10(X)
        Log10 = Log(X) / Log(10#)
    End Function


Sub Generate_layout()

    Set WS_L = Worksheets("Input-data")
    layer_amounts = WS_L.Cells(5, 1).Value
    Sheets.Add(After:=WS_L).Name = "Drawing"
    Set WS_T = Worksheets("Drawing")
    WS_T.Range("A6:A300").RowHeight = 3   'set the height of the selected area

    N_item = WS_L.Cells(2, 3).Value 'the number of items
    'N_depth = WS_L.Cells(7, 1).Value   'The depth of the investigated part

    WS_T.Cells(2, 12).Value = WS_L.Cells(3, 4).Value
    WS_T.Cells(2, 13).Value = WS_L.Cells(3, 5).Value
    WS_T.Cells(2, 12).Font.Bold = True
    WS_T.Cells(2, 12).Font.Color = vbRed
    downline = N_pos + 5

    'WS_T.Columns("L").Interior.ColorIndex = 34
    'WS_T.Columns("o").Interior.ColorIndex = 34
    'WS_T.Columns("p").Interior.ColorIndex = 34
    'WS_T.Columns("q").Interior.ColorIndex = 34
    'WS_T.Columns("r").Interior.ColorIndex = 34
    WS_T.Rows("1:5").Interior.ColorIndex = 0
    WS_T.Columns("o").ColumnWidth = 15

    WS_T.Cells(3, 12).Value = "壓抂怺搙"
    WS_T.Cells(5, 12).Value = "GL-(m)"
    WS_T.Cells(3, 13).Value = "昗崅 TP"
    WS_T.Cells(5, 13).Value = "(m)"
    WS_T.Cells(3, 14).Value = "憌岤"
    WS_T.Cells(5, 14).Value = "(m)"
    WS_T.Cells(3, 15).Value = "搚幙柤"
    WS_T.Cells(3, 16).Value = "抧幙婰崋"
    gravity = WS_L.Cells(14, 2).Value

    For i = 12 To 29
        Set target_cells = WS_T.Range(WS_T.Cells(3, i), WS_T.Cells(5, i))
        target_cells.Interior.ColorIndex = 35
    Next i

    For i = 12 To 14
        WS_T.Cells(5, i).Interior.ColorIndex = 35
    Set target_cells = WS_T.Range(WS_T.Cells(3, i), WS_T.Cells(4, i))
        target_cells.Merge
    Next i

    For i = 15 To 16
        Set target_cells = WS_T.Range(WS_T.Cells(3, i), WS_T.Cells(5, i))
        target_cells.Merge
    Next i

    WS_T.Cells(3, 17).Value = "姺嶼俶抣"
    Set target_cells = WS_T.Range(WS_T.Cells(3, 17), WS_T.Cells(3, 18))
    target_cells.Merge
    WS_T.Cells(4, 17).Value = "暯嬒抣"
    WS_T.Cells(4, 18).Value = "昗弨曃嵎"

    WS_T.Cells(3, 19).Value = "P攇懍搙"
    WS_T.Cells(4, 19).Value = "Vp"
    WS_T.Cells(5, 19).Value = "(m/s)"
    WS_T.Cells(3, 20).Value = "S攇懍搙"
    WS_T.Cells(4, 20).Value = "Vs"
    WS_T.Cells(5, 20).Value = "(m/s)"
    WS_T.Cells(3, 21).Value = "億傾僜儞斾"
    WS_T.Cells(4, 21).Value = "兯"
    WS_T.Cells(3, 22).Value = "幖弫枾搙"
    WS_T.Cells(4, 22).Value = "兿t"
    WS_T.Cells(5, 22).Value = "(g/cm3)"
    
    WS_T.Cells(3, 23).Value = "扨埵懱愊廳検"
    WS_T.Cells(4, 23).Value = "兞t"
    WS_T.Cells(5, 23).Value = "(kN/m3)"
    
    WS_T.Cells(3, 24).Value = "擲拝椡"
    WS_T.Cells(4, 24).Value = "c"
    WS_T.Cells(5, 24).Value = "(kN/m2)"
    
    WS_T.Cells(3, 25).Value = "撪晹杸嶤妏"
    WS_T.Cells(4, 25).Value = "冇"
    WS_T.Cells(5, 25).Value = "(搙)"
    
        
    WS_T.Cells(3, 26).Value = "曄宍學悢"
    WS_T.Cells(4, 26).Value = "E0"
    WS_T.Cells(5, 26).Value = "(MN/m2)"
    
    WS_T.Cells(3, 27).Value = "偣傫抐抏惈學悢"
    WS_T.Cells(4, 27).Value = "G0"
    WS_T.Cells(5, 27).Value = "(MN/m2)"
    
    WS_T.Cells(3, 28).Value = "婎弨偣傫抐傂偢傒"
    WS_T.Cells(4, 28).Value = "兞0.5"
    WS_T.Cells(5, 28).Value = "(%)"
    
    WS_T.Cells(3, 29).Value = "棜楌尭悐棪"
    WS_T.Cells(4, 29).Value = "hmax"
    WS_T.Cells(5, 29).Value = "(%)"
    
    'from bottom depth to transed N value

    For columncount = 12 To 18
        nowthick = WS_L.Cells(6, 4).Value
        pointerup = 6
        countnum = 1
        Do While (countnum <= N_item)

            rawdepth = WS_L.Cells(countnum + 5, 4).Value
            Depth = rawdepth * 10
            ''''calculation of rawthick
            If pointerup = 6 Then
            rawthick = rawdepth
            Else
            rawthick = rawdepth - prevdepth
            End If
            
            
            prevdepth = rawdepth
            Depth = Application.RoundUp(Depth, 0)
            pointerdown = Depth + 5
            Set Mergetargetcells = WS_T.Range(WS_T.Cells(pointerup, columncount), WS_T.Cells(pointerdown, columncount))
            Mergetargetcells.Merge
            
            
            
            
            'Give value & get item name
            Dim itemname As String
            itemname = WS_T.Cells(3, columncount).Value
            
            'switch different items

            Select Case itemname
            Case "壓抂怺搙"
                 WS_T.Cells(pointerup, columncount).Value = WS_L.Cells(countnum + 5, columncount - 8).Value
            Case "昗崅 TP"
                 WS_T.Cells(pointerup, columncount).Value = WS_L.Cells(3, 5).Value - WS_L.Cells(countnum + 5, columncount - 9).Value
            Case "憌岤"
                  WS_T.Cells(pointerup, columncount).Value = rawthick
            Case "搚幙柤"
                  WS_T.Cells(pointerup, columncount).Value = WS_L.Cells(countnum + 5, columncount - 10).Value
            Case "抧幙婰崋"
                  WS_T.Cells(pointerup, columncount).Value = WS_L.Cells(countnum + 5, columncount - 10).Value
            Case "姺嶼俶抣"
                  WS_T.Cells(pointerup, columncount).Value = WS_L.Cells(countnum + 5, columncount - 10).Value
                  WS_T.Cells(pointerup, columncount + 1).Value = WS_L.Cells(countnum + 5, columncount - 9).Value
                  End Select
            
           'refresh the count value
            pointerup = pointerdown + 1
            countnum = countnum + 1
            
        Loop



    Next columncount
    
    
     ''''start the column 19 to 21  vp vs & poison ratio
  ps_item = WS_L.Cells(2, 14).Value 'the number of ps items
    For columncount = 19 To 21
        nowthick = WS_L.Cells(6, 13).Value
        pointerup = 6
        countnum = 1
        Do While (countnum <= ps_item)

            rawdepth = WS_L.Cells(countnum + 5, 13).Value
            Depth = rawdepth * 10
            
            Depth = Application.RoundUp(Depth, 0)
            pointerdown = Depth + 5
            Set Mergetargetcells = WS_T.Range(WS_T.Cells(pointerup, columncount), WS_T.Cells(pointerdown, columncount))
            Mergetargetcells.Merge
            
            
            
            
            
            'Give value & get item name
            itemname = WS_T.Cells(3, columncount).Value

            'switch different items

            Select Case itemname
                Case "P攇懍搙"
                    WS_T.Cells(pointerup, columncount).Value = WS_L.Cells(countnum + 5, columncount - 5).Value
                Case "S攇懍搙"
                    WS_T.Cells(pointerup, columncount).Value = WS_L.Cells(countnum + 5, columncount - 5).Value
                Case "億傾僜儞斾"
                    WS_T.Cells(pointerup, columncount).Value = 0.5 * (WS_L.Cells(countnum + 5, columncount - 7).Value ^ 2 - 2 * WS_L.Cells(countnum + 5, columncount - 6).Value ^ 2) / (WS_L.Cells(countnum + 5, columncount - 7).Value ^ 2 - WS_L.Cells(countnum + 5, columncount - 6).Value ^ 2)
            End Select

            'refresh the count value
            pointerup = pointerdown + 1
            countnum = countnum + 1

        Loop


    Next columncount





    ''''end column 19 to 21



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''start the column 22 the individual pt settings  the 幖弫枾搙 and 23 yt
  pt_item = WS_L.Cells(2, 17).Value 'the number of pt items
  
     For columncount = 22 To 23
        nowthick = WS_L.Cells(6, 16).Value
        pointerup = 6
        countnum = 1
        Do While (countnum <= pt_item)

            rawdepth = WS_L.Cells(countnum + 5, 16).Value
            Depth = rawdepth * 10
            
            Depth = Application.RoundUp(Depth, 0)
            pointerdown = Depth + 5
            Set Mergetargetcells = WS_T.Range(WS_T.Cells(pointerup, columncount), WS_T.Cells(pointerdown, columncount))
            Mergetargetcells.Merge
            'Give value & get item name
            
            
            
                
            'Give value & get item name
            itemname = WS_T.Cells(3, columncount).Value
            
            'switch different items

            Select Case itemname
            Case "幖弫枾搙"
                 WS_T.Cells(pointerup, columncount).Value = WS_L.Cells(countnum + 5, columncount - 5).Value
            
            Case "扨埵懱愊廳検"
                 WS_T.Cells(pointerup, columncount).Value = WS_L.Cells(countnum + 5, columncount - 6).Value * gravity
           

            End Select
        pointerup = pointerdown + 1
            countnum = countnum + 1
        Loop
        Next columncount
        







    ''''end column 22&23 individual pt settings
    
    

''''''''''''''''''''''''''''''''''''column 24 to 26
 For columncount = 24 To 26
        nowthick = WS_L.Cells(6, 4).Value
        pointerup = 6
        countnum = 1

        Do While (countnum <= N_item)
         indexfai = WS_L.Cells(countnum + 5, 11).Value
         indexc = WS_L.Cells(countnum + 5, 10).Value
         indexe = WS_L.Cells(countnum + 5, 12).Value

            rawdepth = WS_L.Cells(countnum + 5, 4).Value
            Depth = rawdepth * 10
            
            Depth = Application.RoundUp(Depth, 0)
            pointerdown = Depth + 5
            Set Mergetargetcells = WS_T.Range(WS_T.Cells(pointerup, columncount), WS_T.Cells(pointerdown, columncount))
            Mergetargetcells.Merge
            
            
            'Give value & get item name
            itemname = WS_T.Cells(3, columncount).Value
            
            'switch different items

            Select Case itemname
            

            Case "擲拝椡"
                Select Case indexc
                Case "1"
                WS_T.Cells(pointerup, columncount).Value = 0
                Case "2"
                WS_T.Cells(pointerup, columncount).Value = 10 * WS_L.Cells(countnum + 5, 7).Value
                Case "3"
                WS_T.Cells(pointerup, columncount).Value = 15.2 * WS_L.Cells(countnum + 5, 7).Value ^ 0.327
                Case "4"
                WS_T.Cells(pointerup, columncount).Value = 25.3 * WS_L.Cells(countnum + 5, 7).Value ^ 0.334
                Case "5"
                WS_T.Cells(pointerup, columncount).Value = 16.2 * WS_L.Cells(countnum + 5, 7).Value ^ 0.606
                End Select
                
                  
            Case "撪晹杸嶤妏"
                Select Case indexfai
                Case "1"
                WS_T.Cells(pointerup, columncount).Value = 15 + (15 * WS_L.Cells(countnum + 5, 7).Value) ^ 0.5
                Case "2"
                WS_T.Cells(pointerup, columncount).Value = 0
                Case "3"
                WS_T.Cells(pointerup, columncount).Value = 5.1 * Log10(WS_L.Cells(countnum + 5, 7).Value) + 29.3
                Case "4"
                WS_T.Cells(pointerup, columncount).Value = 6.82 * Log10(WS_L.Cells(countnum + 5, 7).Value) + 21.5
                Case "5"
                WS_T.Cells(pointerup, columncount).Value = 0.888 * Log10(WS_L.Cells(countnum + 5, 7).Value) + 19.3
                Case "6"
                WS_T.Cells(pointerup, columncount).Value = 15 + (20 * WS_L.Cells(countnum + 5, 7).Value) ^ 0.5
                End Select
                
                
            Case "曄宍學悢"
                Select Case indexe
                Case "1"
                WS_T.Cells(pointerup, columncount).Value = 700 * WS_L.Cells(countnum + 5, 7).Value / 1000
                Case "2"
                WS_T.Cells(pointerup, columncount).Value = 700 * WS_L.Cells(countnum + 5, 7).Value / 1000
                Case "3"
                WS_T.Cells(pointerup, columncount).Value = 2659 * WS_L.Cells(countnum + 5, columncount - 19).Value ^ 0.69 / 1000
                Case "4"
                WS_T.Cells(pointerup, columncount).Value = 2659 * WS_L.Cells(countnum + 5, columncount - 19).Value ^ 0.69 / 1000
                Case "5"
                WS_T.Cells(pointerup, columncount).Value = 2659 * WS_L.Cells(countnum + 5, columncount - 19).Value ^ 0.69 / 1000
                End Select
                

           End Select
            
           'refresh the count value
            pointerup = pointerdown + 1
            countnum = countnum + 1
            
        Loop


    Next columncount



    'change the bottom color

   ' Worksheets("Drawing").Rows("pointerup:9999").Interior.ColorIndex = 0
    'Range("E9999").Activate
 
''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''column 28 and 29
 For columncount = 28 To 29
        nowthick = WS_L.Cells(6, 4).Value
        pointerup = 6
        countnum = 1

        Do While (countnum <= N_item)
         indexfai = WS_L.Cells(countnum + 5, 11).Value
         indexc = WS_L.Cells(countnum + 5, 10).Value
         indexe = WS_L.Cells(countnum + 5, 12).Value

            rawdepth = WS_L.Cells(countnum + 5, 4).Value
            Depth = rawdepth * 10
            
            Depth = Application.RoundUp(Depth, 0)
            pointerdown = Depth + 5
            Set Mergetargetcells = WS_T.Range(WS_T.Cells(pointerup, columncount), WS_T.Cells(pointerdown, columncount))
            Mergetargetcells.Merge
            
            
            'Give value & get item name
            itemname = WS_T.Cells(3, columncount).Value
            
            'switch different items

            Select Case itemname
            Case "婎弨偣傫抐傂偢傒"
                  WS_T.Cells(pointerup, columncount).Value = "-"
            Case "棜楌尭悐棪"
                  WS_T.Cells(pointerup, columncount).Value = "-"
           End Select
            
           'refresh the count value
            pointerup = pointerdown + 1
            countnum = countnum + 1
            
        Loop


    Next columncount



    'change the bottom color

   ' Worksheets("Drawing").Rows("pointerup:9999").Interior.ColorIndex = 0
    'Range("E9999").Activate
 
''''''''''''''''''''''''''''''''''''



''''''''''''''''''''''''''''''''''''column 27, seperated part
  columncount = 27

       ' nowthick = WS_L.Cells(6, 20).Value
        pointerup = 6
        countnumps = 1
        countnumpt = 1
        checka = 1
        checkb = 1

Do While (WS_L.Cells(countnumps + 5, 13).Value Or WS_L.Cells(countnumpt + 5, 16))
        If WS_L.Cells(countnumps + 3, 13).Value = "" Or WS_L.Cells(countnumpt + 3, 16).Value = "" Then
        GoTo endloop
        End If
        
        If checka = "" And checkb = "" Then
        GoTo endloop
        End If
        
        


        If WS_L.Cells(countnumps + 5, 13).Value = "" And WS_L.Cells(countnumpt + 5, 16).Value = "" Then
        GoTo endcalculation
        End If
        
        
        

        If WS_L.Cells(countnumps + 5, 13).Value Then
        
        psvalue = WS_L.Cells(countnumps + 5, 13).Value
        End If
        
        If WS_L.Cells(countnumpt + 5, 16).Value Then
        ptvalue = WS_L.Cells(countnumpt + 5, 16).Value
        End If
        
        

            If psvalue > ptvalue Then
            If ptvalue = rawdepth Then
            countnumpt = countnumpt + 1
            GoTo endcalculation
            Else
            rawdepth = ptvalue
            countnumpt = countnumpt + 1
            End If
            

            ElseIf psvalue < ptvalue Then
            If psvalue = rawdepth Then
            countnumps = countnumps + 1
            GoTo endcalculation
            Else
            rawdepth = psvalue
            countnumps = countnumps + 1
            End If

            ElseIf psvalue = ptvalue Then
            If countnumpt = countnumps Then

            rawdepth = psvalue
            countnumps = countnumps + 1
            Else
            checka = WS_L.Cells(countnumps + 6, 13).Value
            checkb = WS_L.Cells(countnumpt + 6, 16).Value
          
            If checka Or checkb Then
            countnumpt = countnumpt + 1
            GoTo endcalculation
            Else

            rawdepth = psvalue
            End If
            End If
            End If
            
            
             
            



            Depth = rawdepth * 10
            
            Depth = Application.RoundUp(Depth, 0)
            pointerdown = Depth + 5
            Set Mergetargetcells = WS_T.Range(WS_T.Cells(pointerup, columncount), WS_T.Cells(pointerdown, columncount))
            Mergetargetcells.Merge
            
            
            'Give value & get item name
            itemname = WS_T.Cells(3, columncount).Value
            
            'switch different items

            Select Case itemname
                
            Case "偣傫抐抏惈學悢"
            v6 = WS_T.Cells(pointerup, columncount - 5).Value
            t6 = WS_T.Cells(pointerup, columncount - 7).Value
           If t6 = "" Then
           t6 = t6save
           Else
           t6save = t6
           End If
           
            WS_T.Cells(pointerup, columncount).Value = v6 * t6 ^ 2 / 1000

           End Select
            
           'refresh the count value
            pointerup = pointerdown + 1

endcalculation:

      
Loop

endloop:





    'change the bottom color

   ' Worksheets("Drawing").Rows("pointerup:9999").Interior.ColorIndex = 0
    'Range("E9999").Activate
 
''''''''''''''''''''''''''''''''''''



  
WS_T.Rows("6:99999").HorizontalAlignment = Excel.xlCenter





End Sub







