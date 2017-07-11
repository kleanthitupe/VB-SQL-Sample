Private Sub buttonClearInputs_Click()
    ClearAll Me
End Sub

Private Sub buttonClearLast_Click()
    ClearLastInputed Me
End Sub

Private Sub buttonListCattle_Click()
    Dim bredHeifersID As Integer
    Dim bullsID As Integer
    Dim cowsID As Integer
    Dim calvesID As Integer
    Dim yearlingsID As Integer
    
    Dim SH_ID As Long
    Dim F_ID As Long
    
    SH_ID = Me.txtShareholderID.Value
    F_ID = Me.txtCurrentField.Value
        
    'values from the lookup table: Category
    bredHeifersID = DLookup("[CAT_ID]", "Category", "[CAT_Category] = 'Bred Heifers'")
    bullsID = DLookup("[CAT_ID]", "Category", "[CAT_Category] = 'Bulls'")
    cowsID = DLookup("[CAT_ID]", "Category", "[CAT_Category] = 'Cows'")
    calvesID = DLookup("[CAT_ID]", "Category", "[CAT_Category] = 'Calves'")
    yearlingsID = DLookup("[CAT_ID]", "Category", "[CAT_Category] = 'Yearlings'")

    Me.txtAvailableBredHeif.Value = DCount("BOV_ID", "ActiveCattle", "[CAT_ID] = " & bredHeifersID & " AND [F_ID] = " & F_ID & " AND [SH_ID] = " & SH_ID & " ")
    Me.txtAvailableBulls.Value = DCount("BOV_ID", "ActiveCattle", "[CAT_ID] = " & bullsID & " AND [F_ID] = " & F_ID & " AND [SH_ID] = " & SH_ID & " ")
    Me.txtAvailableCows.Value = DCount("BOV_ID", "ActiveCattle", "[CAT_ID] = " & cowsID & " AND [F_ID] = " & F_ID & " AND [SH_ID] = " & SH_ID & " ")
    Me.txtAvailableCalves.Value = DCount("BOV_ID", "ActiveCattle", "[CAT_ID] = " & calvesID & " AND [F_ID] = " & F_ID & " AND [SH_ID] = " & SH_ID & " ")
    Me.txtAvailableYearlings.Value = DCount("BOV_ID", "ActiveCattle", "[CAT_ID] = " & yearlingsID & " AND [F_ID] = " & F_ID & " AND [SH_ID] = " & SH_ID & " ")
    
End Sub

Private Sub buttonManifestOut_Click()
    
    If (IsNull(Me.comboShareholderID.Value) Or IsNull(Me.txtDateOut.Value) Or IsNull(Me.txtManifestNum.Value) Or IsNull(Me.comboFieldID.Value)) Then
        Call MsgBox("Required input: Shareholder, Date, Manifest Number, and Current Field!", , "CALMS")
        Exit Sub
    Else
        prompt = "Are you sure you want to create this Manifest Out?"
    
        If (MsgBox(prompt, 1, "CALMS") = 1) Then
        
        'values from the lookup table: Category
        bredHeifersID = DLookup("[CAT_ID]", "Category", "[CAT_Category] = 'Bred Heifers'")
        bullsID = DLookup("[CAT_ID]", "Category", "[CAT_Category] = 'Bulls'")
        cowsID = DLookup("[CAT_ID]", "Category", "[CAT_Category] = 'Cows'")
        calvesID = DLookup("[CAT_ID]", "Category", "[CAT_Category] = 'Calves'")
        yearlingsID = DLookup("[CAT_ID]", "Category", "[CAT_Category] = 'Yearlings'")
        
        Dim shareholderID As Long
        Dim dateOut As Date
        Dim fromFieldID As Long
        
        Dim countBredHeif As Integer
        Dim countBulls As Integer
        Dim countCows As Integer
        Dim countCalves As Integer
        Dim countYearlings As Integer
    
        countBredHeif = 0: countBulls = 0: countCows = 0: countCalves = 0: countYearlings = 0
    
        fromFieldID = Me.txtCurrentField.Value
        dateOut = Me.txtDateOut.Value
        shareholderID = Me.txtShareholderID.Value
    
    
        Dim rsLocation As DAO.Recordset
        Set rsLocation = CurrentDb.OpenRecordset("Location")
        
        Dim rsManifest As DAO.Recordset
        Set rsManifest = CurrentDb.OpenRecordset("Manifest")
        
        Dim rsBovine As DAO.Recordset
        Set rsBovine = CurrentDb.OpenRecordset("Bovine")
        
        Dim rs As DAO.Recordset
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM MoveCattleQuery WHERE Manifest.SH_ID = " & shareholderID & " AND Location.F_ID = " & fromFieldID & " ;")
        
        rsManifest.AddNew
        rsManifest("SH_ID").Value = shareholderID
        rsManifest("MANIFEST_Date").Value = dateOut
        rsManifest("MANIFEST_In/Out").Value = True
        rsManifest("MANIFEST_Information").Value = Me.txtComments.Value
        rsManifest("MANIFEST_Number").Value = Me.txtManifestNum.Value
        rsManifest.Update
        
        rsManifest.MoveLast
        
        Dim manifestOutId As Long
        Dim totalAUMs As Double
        manifestOutId = rsManifest!MANIFEST_ID
        totalAUMs = 0
        
        
        'Check to see if the recordset actually contains rows
        If Not (rs.EOF And rs.BOF) Then
            rs.MoveFirst 
            Do Until rs.EOF = True
            
                'check if they are Bred Heifers, Bulls etc in these if statements
                If Not IsNull(Me.txtBredHeifNum.Value) Then
                    If (countBredHeif < (Me.txtBredHeifNum.Value) And rs("CAT_ID") = bredHeifersID) Then
    
                        'Perform an edit
                        rs.Edit
                        rs!LOC_Date_OUT = dateOut
                        rs("Manifest_ID_Out") = manifestOutId
                        rs("LOC_Selected") = False 'The other way to refer to a field, same thing as the row above
                        rs("LOC_AUMs_USED") = metabolicFunctAUM(rs("BOV_Weight"), 1, seasonalFactorTotal(rs!LOC_Date_IN, rs!LOC_Date_OUT))
                        totalAUMs = totalAUMs + metabolicFunctAUM(rs("BOV_Weight"), 1, seasonalFactorTotal(rs!LOC_Date_IN, rs!LOC_Date_OUT))
                        rs("LOC_Manifest_Out") = True
                        rs.Update
    
                        'Move to the next record.
                        countBredHeif = countBredHeif + 1
                    End If
                End If
            
                If Not IsNull(Me.txtBullsNum.Value) Then
                    If (countBulls < (Me.txtBullsNum.Value) And rs("CAT_ID") = bullsID) Then
                    
                        'Perform an edit
                        rs.Edit
                        rs!LOC_Date_OUT = dateOut
                        rs("Manifest_ID_Out") = manifestOutId
                        rs("LOC_Selected") = False 'The other way to refer to a field, same thing as the row above
                        rs("LOC_AUMs_USED") = metabolicFunctAUM(rs("BOV_Weight"), 1, seasonalFactorTotal(rs!LOC_Date_IN, rs!LOC_Date_OUT))
                        totalAUMs = totalAUMs + metabolicFunctAUM(rs("BOV_Weight"), 1, seasonalFactorTotal(rs!LOC_Date_IN, rs!LOC_Date_OUT))
                        rs("LOC_Manifest_Out") = True
                        rs.Update
    
                        'Move to the next record.
                        countBulls = countBulls + 1
                    End If
                End If
            
                If Not IsNull(Me.txtCowsNum.Value) Then
                    If (countCows < (Me.txtCowsNum.Value) And rs("CAT_ID") = cowsID) Then
                    
                        'Perform an edit
                        rs.Edit
                        rs!LOC_Date_OUT = dateOut
                        rs("Manifest_ID_Out") = manifestOutId
                        rs("LOC_Selected") = False 'The other way to refer to a field, same thing as the row above
                        rs("LOC_AUMs_USED") = metabolicFunctAUM(rs("BOV_Weight"), 1, seasonalFactorTotal(rs!LOC_Date_IN, rs!LOC_Date_OUT))
                        totalAUMs = totalAUMs + metabolicFunctAUM(rs("BOV_Weight"), 1, seasonalFactorTotal(rs!LOC_Date_IN, rs!LOC_Date_OUT))
                        rs("LOC_Manifest_Out") = True
                        rs.Update
    
                        'Move to the next record.
                        countCows = countCows + 1
                    End If
                End If
            
                If Not IsNull(Me.txtCalvesNum.Value) Then
                    If (countCalves < (Me.txtCalvesNum.Value) And rs("CAT_ID") = calvesID) Then
                        
                        'Perform an edit
                        rs.Edit
                        rs!LOC_Date_OUT = dateOut
                        rs("Manifest_ID_Out") = manifestOutId
                        rs("LOC_Selected") = False 'The other way to refer to a field, same thing as the row above
                        rs("LOC_AUMs_USED") = metabolicFunctAUM(rs("BOV_Weight"), 1, seasonalFactorTotal(rs!LOC_Date_IN, rs!LOC_Date_OUT))
                        totalAUMs = totalAUMs + metabolicFunctAUM(rs("BOV_Weight"), 1, seasonalFactorTotal(rs!LOC_Date_IN, rs!LOC_Date_OUT))
                        rs("LOC_Manifest_Out") = True
                        rs.Update
    
                        'Move to the next record.
                        countCalves = countCalves + 1
                    End If
                End If
            
                If Not IsNull(Me.txtYearlingsNum.Value) Then
                    If (countYearlings < (Me.txtYearlingsNum.Value) And rs("CAT_ID") = yearlingsID) Then
                    
            
                        'Perform an edit
                        rs.Edit
                        rs!LOC_Date_OUT = dateOut
                        rs("Manifest_ID_Out") = manifestOutId
                        rs("LOC_Selected") = False 'The other way to refer to a field, same thing as the row above
                        rs("LOC_AUMs_USED") = metabolicFunctAUM(rs("BOV_Weight"), 1, seasonalFactorTotal(rs!LOC_Date_IN, rs!LOC_Date_OUT))
                        totalAUMs = totalAUMs + metabolicFunctAUM(rs("BOV_Weight"), 1, seasonalFactorTotal(rs!LOC_Date_IN, rs!LOC_Date_OUT))
                        rs("LOC_Manifest_Out") = True
                        rs.Update
    
                        'Move to the next record. 
                        countYearlings = countYearlings + 1
                    End If
                 End If
                rs.MoveNext
            Loop
            ClearLastInputed Me
            'MsgBox (totalAUMs)
            
            Me.txtLastShareholer.Value = Me.comboShareholderID.Column(1)
            Me.txtLastDateOut.Value = Me.txtDateOut.Value
            Me.txtLastManifestOut.Value = Me.txtManifestNum.Value
            Me.txtLastFromField.Value = Me.comboFieldID.Column(0)
            Me.txtLastBredHeifersNum.Value = countBredHeif
            Me.txtLastBullsNum.Value = countBulls
            Me.txtLastCowsNum.Value = countCows
            Me.txtLastCalvesNum.Value = countCalves
            Me.txtLastYearlingsNum.Value = countYearlings
            Me.txtLastComments.Value = Me.txtComments.Value
            If (DCount("LOC_AUMs_USED", "ManifestOutLocAUMs", "Manifest_ID_Out = " & manifestOutId & " ") > 0) Then
                totalAUMs = DSum("LOC_AUMs_USED", "ManifestOutLocAUMs", "Manifest_ID_Out = " & manifestOutId & " ")
            End If
            
            Dim pairGrazing As Double
            Dim monthGraze As Double
                
            pairGrazing = DLookup("[Quantity]", "ShareFactor", "[Factor] = 'Stocking Rate'")
            monthGraze = DLookup("[Quantity]", "ShareFactor", "[Factor] = 'Basic Season'")
            
            Me.txtAUMsLast.Value = totalAUMs
            Me.txtSharesLast.Value = Round(totalAUMs / (pairGrazing * monthGraze), 2)
                    
            rsManifest.Edit
            rsManifest!MANIFEST_AUMs_Man_Out = totalAUMs
            rsManifest!MANIFEST_Shares_Man_Out = Round(totalAUMs / (pairGrazing * monthGraze), 2)
            rsManifest.Update
            
            Dim rsShareholder As DAO.Recordset
            Dim sqlSharesAvailable As String
            sqlSharesAvailable = "SELECT * FROM Shareholder WHERE SH_ID = " & Me.txtShareholderID.Value & ";"
    
            Set rsShareholder = CurrentDb.OpenRecordset(sqlSharesAvailable)

            'code that updates shares table May17
            rsShareholder.Edit
            rsShareholder!SH_Shares_Used_MOut = rsShareholder!SH_Shares_Used_MOut + Round(totalAUMs / (pairGrazing * monthGraze), 2)
            rsShareholder!SH_Num_Avail_Shares = rsShareholder!SH_Num_Avail_Shares - Round(totalAUMs / (pairGrazing * monthGraze), 2)
            rsShareholder.Update
    
            ClearAll Me
            Call ClearGroupWithTag(Me.Form, "clearForNextShareholder")
        Else
            Call MsgBox("There are no cattle with this criteria.", , "CALMS")
        End If
        
        
        Call MsgBox("Manifest Out was entered. Finished moving cattle.", , "CALMS")
    
        rsShareholder.Close
        rsBovine.Close
        rsManifest.Close
        rsLocation.Close
        rs.Close 'Close the recordset
        Set rsShareholder = Nothing 'Clean up
        Set rsBovine = Nothing
        Set rsManifest = Nothing
        Set rsLocation = Nothing
        Set rs = Nothing
        End If
    End If
    totalAUMs = 0
    
End Sub



Private Sub comboFieldID_Change()
    Me.txtCurrentField.Value = Me.comboFieldID.Column(1)
    
    Dim bredHeifersID As Integer
    Dim bullsID As Integer
    Dim cowsID As Integer
    Dim calvesID As Integer
    Dim yearlingsID As Integer
    
    Dim SH_ID As Long
    Dim F_ID As Long
    If Not (IsNull(Me.txtShareholderID.Value)) Then
        SH_ID = Me.txtShareholderID.Value
    
        F_ID = Me.comboFieldID.Column(1)
            
        
        
        
        Dim sql As String
        sql = "SELECT DISTINCT Manifest.[MANIFEST_Number], Manifest.[MANIFEST_ID], Manifest.[MANIFEST_Date] FROM(Field INNER JOIN (Location INNER JOIN Bovine ON Location.BOV_ID = Bovine.BOV_ID) ON Field.F_ID = Location.F_ID) INNER JOIN (Manifest INNER JOIN Shareholder ON Manifest.SH_ID = Shareholder.SH_ID) ON Bovine.Manifest_ID_In = Manifest.MANIFEST_ID WHERE Field.F_Active = Yes AND Shareholder.SH_ID = " & comboShareholderID.Column(0) & " AND Location.LOC_Date_OUT Is Null AND Location.F_ID = " & Me.txtCurrentField.Value & " ORDER BY Manifest.MANIFEST_Number;"
        'sql = "SELECT DISTINCT Manifest.[MANIFEST_Number], Manifest.[MANIFEST_ID], Manifest.[MANIFEST_Date] FROM (Field INNER JOIN (Location INNER JOIN Bovine ON Location.BOV_ID = Bovine.BOV_ID) ON Field.F_ID = Location.F_ID) INNER JOIN (Manifest INNER JOIN Shareholder ON Manifest.SH_ID = Shareholder.SH_ID) ON Bovine.Manifest_ID_In = Manifest.MANIFEST_ID WHERE Field.F_Active = Yes AND Shareholder.SH_ID = " & comboShareholderID.Column(0) & " AND Location.LOC_Date_OUT Is Null AND Location.F_ID = " & Me.txtCurrentField.Value & ";"
        Me.comboManifest.RowSource = sql
        Me.comboManifest.Requery
        
    End If
End Sub


Private Sub comboFieldID_Click()
    
End Sub

Private Sub comboManifest_Change()
    Dim sql As String
    sql = "SELECT * FROM(Field INNER JOIN (Location INNER JOIN Bovine ON Location.BOV_ID = Bovine.BOV_ID) ON Field.F_ID = Location.F_ID) INNER JOIN (Manifest INNER JOIN Shareholder ON Manifest.SH_ID = Shareholder.SH_ID) ON Bovine.Manifest_ID_In = Manifest.MANIFEST_ID WHERE Field.F_Active = Yes AND Shareholder.SH_ID = " & comboShareholderID.Column(0) & " AND Location.LOC_Date_OUT Is Null AND Location.F_ID = " & Me.txtCurrentField.Value & " AND Manifest.[MANIFEST_ID] = " & Me.comboManifest.Column(1) & " ORDER BY Manifest.MANIFEST_Number;"
    Dim rsManifest As DAO.Recordset
    Set rsManifest = CurrentDb.OpenRecordset(sql)


    'Me.txtDateInMan.Value = Me.comboManifest.Column(2)
    Me.txtDateInMan.Value = rsManifest![MANIFEST_Date]
    Me.txtPredictedDateOut.Value = rsManifest![MANIFEST_Predicted_Date_Out]
    Me.txtPredictedAUMs.Value = rsManifest![MANIFEST_AUMs_Predicted]
    Me.txtPredictedShares.Value = rsManifest![MANIFEST_Shares_Predicted]
    
    If Not (IsNull(Me.txtShareholderID.Value) Or IsNull(Me.comboFieldID.Column(1))) Then
        SH_ID = Me.txtShareholderID.Value
    
        F_ID = Me.comboFieldID.Column(1)
        'values taken from the lookup table: Category
        bredHeifersID = DLookup("[CAT_ID]", "Category", "[CAT_Category] = 'Bred Heifers'")
        bullsID = DLookup("[CAT_ID]", "Category", "[CAT_Category] = 'Bulls'")
        cowsID = DLookup("[CAT_ID]", "Category", "[CAT_Category] = 'Cows'")
        calvesID = DLookup("[CAT_ID]", "Category", "[CAT_Category] = 'Calves'")
        yearlingsID = DLookup("[CAT_ID]", "Category", "[CAT_Category] = 'Yearlings'")
    
        Me.txtAvailableBredHeif.Value = DCount("BOV_ID", "ActiveCattle", "[CAT_ID] = " & bredHeifersID & " AND [F_ID] = " & F_ID & " AND [SH_ID] = " & SH_ID & " AND [Manifest_ID_In] = " & Me.comboManifest.Column(1) & "  ")
        Me.txtAvailableBulls.Value = DCount("BOV_ID", "ActiveCattle", "[CAT_ID] = " & bullsID & " AND [F_ID] = " & F_ID & " AND [SH_ID] = " & SH_ID & " AND [Manifest_ID_In] = " & Me.comboManifest.Column(1) & " ")
        Me.txtAvailableCows.Value = DCount("BOV_ID", "ActiveCattle", "[CAT_ID] = " & cowsID & " AND [F_ID] = " & F_ID & " AND [SH_ID] = " & SH_ID & " AND [Manifest_ID_In] = " & Me.comboManifest.Column(1) & " ")
        Me.txtAvailableCalves.Value = DCount("BOV_ID", "ActiveCattle", "[CAT_ID] = " & calvesID & " AND [F_ID] = " & F_ID & " AND [SH_ID] = " & SH_ID & " AND [Manifest_ID_In] = " & Me.comboManifest.Column(1) & " ")
        Me.txtAvailableYearlings.Value = DCount("BOV_ID", "ActiveCattle", "[CAT_ID] = " & yearlingsID & " AND [F_ID] = " & F_ID & " AND [SH_ID] = " & SH_ID & " AND [Manifest_ID_In] = " & Me.comboManifest.Column(1) & " ")
    End If
End Sub

Private Sub comboShareholderID_Change()
    Me.txtShareholderID.Value = Me.comboShareholderID.Column(0)
    Call ClearGroupWithTag(Me.Form, "clearForNextShareholder")
End Sub

Private Sub comboShareholderID_Click()
    Me.comboFieldID.Value = Null
    Me.txtCurrentField.Value = Null
    Dim sql As String
    sql = "SELECT DISTINCT Field.F_Name, Field.F_ID FROM (Field INNER JOIN (Location INNER JOIN Bovine ON Location.BOV_ID = Bovine.BOV_ID) ON Field.F_ID = Location.F_ID) INNER JOIN (Manifest INNER JOIN Shareholder ON Manifest.SH_ID = Shareholder.SH_ID) ON Bovine.Manifest_ID_In = Manifest.MANIFEST_ID WHERE Field.F_Active = Yes AND Shareholder.SH_ID = " & comboShareholderID.Column(0) & " AND Location.LOC_Date_OUT Is Null ORDER BY Field.F_Name;"
    Me.comboFieldID.RowSource = sql
    Me.comboFieldID.Requery
    
End Sub


Function ClearAll(frm As Form)
Dim ctl As Control

For Each ctl In frm.Controls
    If ctl.Tag = "txtsToClear" Then
        Select Case ctl.ControlType
             Case acTextBox
                ctl.Value = Null
             Case acOptionGroup, acComboBox, acListBox
                ctl.Value = Null
             End Select
    End If
Next
End Function


Function ClearLastInputed(frm As Form)
Dim ctl As Control

For Each ctl In frm.Controls
    If ctl.Tag = "clearLastInputed" Then
        Select Case ctl.ControlType
             Case acTextBox
                ctl.Value = Null
             Case acOptionGroup, acComboBox, acListBox
                ctl.Value = Null
             End Select
    End If
Next
End Function

Function metabolicFunctAUM(averageWeightArg As Integer, cattleQtyArg As Integer, seasonFactorArg As Double) As Double
    'This are values that will be taken from the lookup table
    Dim lbCowVar As Double
    Dim lbCalfVar As Double
    Dim pairGrazing As Double
    Dim monthGraze As Double
    Dim adjFactor As Double
    
    lbCowVar = DLookup("[Quantity]", "ShareFactor", "[Factor] = 'Lb. Cow'")
    lbCalfVar = DLookup("[Quantity]", "ShareFactor", "[Factor] = 'Lb. Calf'")
    pairGrazing = DLookup("[Quantity]", "ShareFactor", "[Factor] = 'Stocking Rate'")
    monthGraze = DLookup("[Quantity]", "ShareFactor", "[Factor] = 'Basic Season'")
    
    
    adjFactor = (averageWeightArg ^ 0.75) / (lbCowVar ^ 0.75 + lbCalfVar ^ 0.75)
    metabolicFunctAUM = cattleQtyArg * adjFactor * seasonFactorArg
    'MsgBox ("The AUMs: " & metabolicFunctAUM)
End Function




Function seasonalFactorTotal(startingDateArg As Date, endingDateArg As Date) As Double
     'Declare variables that will hold how many days fall under each month
    Dim intCount As Integer
    Dim JanCount As Integer: Dim FebCount As Integer: Dim MarCount As Integer: Dim AprCount As Integer: Dim MayCount As Integer: Dim JunCount As Integer: Dim JulCount As Integer
    Dim AugCount As Integer: Dim SeptCount As Integer: Dim OctCount As Integer: Dim NovCount As Integer: Dim DecCount As Integer
    
    'initialize variables
    intCount = 0
    JanCount = 0: FebCount = 0: MarCount = 0: AprCount = 0: MayCount = 0: JunCount = 0: JulCount = 0: AugCount = 0: SeptCount = 0: OctCount = 0: NovCount = 0: DecCount = 0
    
    Dim inDateVar As Date
    Dim outDateVar As Date
    
    inDateVar = startingDateArg
    outDateVar = endingDateArg
    
    Dim numDays As Integer
    Dim getMonthVar As Integer
    
    'Use dateDiff function to find out the difference in days between the two dates entered by the user
    
    
    'MsgBox (inDateVar)
    'MsgBox (outDateVar)
    numDays = DateDiff("d", inDateVar, outDateVar) + 1
    'MsgBox ("the number of days is: " & numDays)
    
    'for loop to run and group how many days fall under each month
    For intCount = 0 To numDays - 1
    
        'get the month and compare it with each month and increment the count for each month appropriately
        getMonthVar = Month(inDateVar)
        If getMonthVar = 1 Then
            JanCount = JanCount + 1
        ElseIf getMonthVar = 2 Then
            FebCount = FebCount + 1
        ElseIf getMonthVar = 3 Then
            MarCount = MarCount + 1
        ElseIf getMonthVar = 4 Then
            AprCount = AprCount + 1
        ElseIf getMonthVar = 5 Then
            MayCount = MayCount + 1
        ElseIf getMonthVar = 6 Then
            JunCount = JunCount + 1
        ElseIf getMonthVar = 7 Then
            JulCount = JulCount + 1
        ElseIf getMonthVar = 8 Then
            AugCount = AugCount + 1
        ElseIf getMonthVar = 9 Then
            SeptCount = SeptCount + 1
        ElseIf getMonthVar = 10 Then
            OctCount = OctCount + 1
        ElseIf getMonthVar = 11 Then
            NovCount = NovCount + 1
        ElseIf getMonthVar = 12 Then
            DecCount = DecCount + 1
        End If
        
        inDateVar = DateAdd("d", 1, inDateVar)
        
    Next intCount
    
    'This are values that will be taken from the lookup table
    Dim Spring As Double
    Dim Summer As Double
    Dim Fall As Double
    Dim Winter As Double
    
    Spring = DLookup("[SEASON_Factor]", "SeasonFactor", "[SEASON_Name] = 'Spring'")
    Summer = DLookup("[SEASON_Factor]", "SeasonFactor", "[SEASON_Name] = 'Summer'")
    Fall = DLookup("[SEASON_Factor]", "SeasonFactor", "[SEASON_Name] = 'Fall'")
    Winter = DLookup("[SEASON_Factor]", "SeasonFactor", "[SEASON_Name] = 'Winter'")
    
    Dim SpringFactorCalc As Double
    Dim SummerFactorCalc As Double
    Dim FallFactorCalc As Double
    Dim WinterFactorCalc As Double
    
    SpringFactorCalc = (JunCount / (Spring * 30))
    SummerFactorCalc = (JulCount / (Summer * 31) + AugCount / (Summer * 31))
    FallFactorCalc = (SeptCount / (Fall * 30) + OctCount / (Fall * 31))
    WinterFactorCalc = (NovCount / (Winter * 30) + DecCount / (Winter * 31) + JanCount / (Winter * 31) + FebCount / (Winter * 28) + MarCount / (Winter * 31))
    
    numDays = 0

    seasonalFactorTotal = (SpringFactorCalc + SummerFactorCalc + FallFactorCalc + WinterFactorCalc)
    
    'MsgBox ("The season factor is: " & seasonalFactorTotal)
    
End Function

Private Sub txtBredHeifNum_KeyPress(KeyAscii As Integer)
    LimitNumeric KeyAscii
End Sub

Private Sub txtBullsNum_KeyPress(KeyAscii As Integer)
    LimitNumeric KeyAscii
End Sub

Private Sub txtCalvesNum_KeyPress(KeyAscii As Integer)
    LimitNumeric KeyAscii
End Sub

Private Sub txtCowsNum_KeyPress(KeyAscii As Integer)
    LimitNumeric KeyAscii
End Sub

Private Sub txtManifestNum_KeyPress(KeyAscii As Integer)
    LimitFieldSize KeyAscii, 15
    LimitAlphanumeric KeyAscii
End Sub

Private Sub txtYearlingsNum_KeyPress(KeyAscii As Integer)
    LimitNumeric KeyAscii
End Sub
