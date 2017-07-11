Option Compare Database


Private Sub buttonAddManifestIn_Click()

    Dim manifestNumberInVar As String
    Dim manifestIdInVar As Long ' should check this if it needs to be declared as a Long
    Dim funtionReturn As Integer
    Dim dateManifestIn As Date
    Dim planedDateOut As Date
    Dim fieldIdVar As Long
    Dim shareholderID As Long
    Dim manifestComment As String
    Dim promptForManifestIn As String
    
    Dim metabolicAUMsBredHeifers As Double
    Dim metabolicAUMsBulls As Double
    Dim metabolicAUMsCows As Double
    Dim metabolicAUMsCalves As Double
    Dim metabolicAUMsYearlings As Double
    Dim totalAUMs As Double
    Dim totalSharesNeeded As Double
    Dim seasonFactor As Double
    metabolicAUMsBredHeifers = 0: metabolicAUMsBulls = 0: metabolicAUMsCows = 0: metabolicAUMsCalves = 0: metabolicAUMsYearlings = 0: totalAUMs = 0: totalSharesNeeded = 0
    
    Dim bredHeiferID As Integer: Dim bredHeiferNum As Integer: Dim bredHeiferWeight As Integer
    Dim bullsID As Integer: Dim bullsNum As Integer: Dim bullsWeight As Integer
    Dim cowID As Integer: Dim cowsNum As Integer: Dim cowsWeight As Integer
    Dim calvesID As Integer: Dim calvesNum As Integer: Dim calvesWeight As Integer
    Dim yearlingsID As Integer: Dim yearlingsNum As Integer: Dim yearlingsWeight As Integer
    
    If (IsNull(Me.txtManifestNumber.Value) Or IsNull(Me.txtDateInManifestIn.Value) Or IsNull(Me.txtPredictedDateOut.Value) Or IsNull(Me.comboFieldSelected.Value) Or IsNull(Me.comboShareholder.Value)) Then
        Call MsgBox("Please enter all the required info: Shareholder, Date In, Date Out, Manifest Number, and Field.", , "CALMS")
        Exit Sub
    Else
    
        Dim db As DAO.Database
        Dim rsManifest As DAO.Recordset
        Dim rsShareholder As Recordset
        Dim sqlSharesAvailable As String
        sqlSharesAvailable = "SELECT * from Shareholder WHERE SH_ID = (" & Me.comboShareholder.Column(0) & ");"
        
        Set db = CurrentDb
        Set rsManifest = db.OpenRecordset("Manifest")
        Set rsShareholder = db.OpenRecordset(sqlSharesAvailable)
        
        dateManifestIn = Me.txtDateInManifestIn.Value
        planedDateOut = Me.txtPredictedDateOut.Value
        fieldIdVar = Me.comboFieldSelected.Column(0)
        shareholderID = Me.comboShareholder.Column(0)
        manifestNumberInVar = Me.txtManifestNumber.Value
    
        If (checkDates(dateManifestIn, planedDateOut) = 0) Then
            Exit Sub
        End If
    
        seasonFactor = seasonalFactorTotal(dateManifestIn, planedDateOut)
    
    
        promptForManifestIn = ("Are you sure you want to go ahead with inserting Manifest #" & manifestNumberInVar & "?")
    
        If (MsgBox(promptForManifestIn, 1, "CALMS") = 1) Then
    
            rsManifest.AddNew
                rsManifest("SH_ID").Value = shareholderID
                rsManifest("MANIFEST_Date").Value = dateManifestIn
                rsManifest("MANIFEST_Predicted_Date_Out").Value = Me.txtPredictedDateOut.Value
                If Not (IsNull(Me.txtCommentManifest.Value)) Then
                    manifestComment = Me.txtCommentManifest.Value
                    rsManifest("MANIFEST_Information").Value = manifestComment
                End If
                rsManifest("MANIFEST_In/Out").Value = False
                rsManifest("MANIFEST_Number").Value = manifestNumberInVar
            rsManifest.Update
            
            rsManifest.MoveLast
            manifestIdInVar = rsManifest!MANIFEST_ID
            
            rsManifest.Close
            Set rsManifest = Nothing
            
            
            ClearLastTxts Me
        
            If Not (IsNull(Me.txtBredHeifers.Value) Or IsNull(Me.txtAvgWeightBredHifers.Value)) Then
                'manifestIdInVar = Me.txtManifestNumber.Value
                bredHeiferID = 1
                bredHeiferID = DLookup("[CAT_ID]", "Category", "[CAT_Category] = 'Bred Heifers'")
                bredHeiferNum = Me.txtBredHeifers.Value
                bredHeiferWeight = Me.txtAvgWeightBredHifers.Value
                Me.txtQtyBredHeifLast.Value = bredHeiferNum
                Me.txtBredHeifWeightLast.Value = bredHeiferWeight
            
            
                metabolicAUMsBredHeifers = metabolicFunctAUM(bredHeiferWeight, bredHeiferNum, seasonFactor)
                funtionReturn = addBovineToTable(manifestIdInVar, bredHeiferID, bredHeiferNum, bredHeiferWeight, dateManifestIn, fieldIdVar)
            End If
            
            If Not (IsNull(Me.txtBulls.Value) Or IsNull(Me.txtAvgWeightBulls.Value) Or IsNull(Me.txtManifestNumber.Value) Or IsNull(Me.txtDateInManifestIn.Value) Or IsNull(Me.comboFieldSelected.Value)) Then
                'manifestIdInVar = Me.txtManifestNumber.Value
                bullsID = 2
                bullsID = DLookup("[CAT_ID]", "Category", "[CAT_Category] = 'Bulls'")
                bullsNum = Me.txtBulls.Value
                bullsWeight = Me.txtAvgWeightBulls.Value
                Me.txtQtyBullsLast.Value = bullsNum
                Me.txtBullsWeightLast.Value = bullsWeight
                
                metabolicAUMsBulls = metabolicFunctAUM(bullsWeight, bullsNum, seasonFactor)
                funtionReturn = addBovineToTable(manifestIdInVar, bullsID, bullsNum, bullsWeight, dateManifestIn, fieldIdVar)
            End If
            
            If Not (IsNull(Me.txtCalves.Value) Or IsNull(Me.txtAvgWeightCalves.Value) Or IsNull(Me.txtManifestNumber.Value) Or IsNull(Me.txtDateInManifestIn.Value) Or IsNull(Me.comboFieldSelected.Value)) Then
                'manifestIdInVar = Me.txtManifestNumber.Value
                calvesID = 3
                calvesID = DLookup("[CAT_ID]", "Category", "[CAT_Category] = 'Calves'")
                calvesNum = Me.txtCalves.Value
                calvesWeight = Me.txtAvgWeightCalves.Value
                Me.txtQtyCalvesLast.Value = calvesNum
                Me.txtCalvesWeightLast.Value = calvesWeight
                
                metabolicAUMsCalves = metabolicFunctAUM(calvesWeight, calvesNum, seasonFactor)
                funtionReturn = addBovineToTable(manifestIdInVar, calvesID, calvesNum, calvesWeight, dateManifestIn, fieldIdVar)
            End If
            
            If Not (IsNull(Me.txtCows.Value) Or IsNull(Me.txtAvgWeightCows.Value) Or IsNull(Me.txtManifestNumber.Value) Or IsNull(Me.txtDateInManifestIn.Value) Or IsNull(Me.comboFieldSelected.Value)) Then
                'manifestIdInVar = Me.txtManifestNumber.Value
                cowID = 4
                cowID = DLookup("[CAT_ID]", "Category", "[CAT_Category] = 'Cows'")
                cowsNum = Me.txtCows.Value
                cowsWeight = Me.txtAvgWeightCows.Value
                Me.txtQtyCowsLast.Value = cowsNum
                Me.txtCowsWeightLast.Value = cowsWeight
            
                metabolicAUMsCows = metabolicFunctAUM(cowsWeight, cowsNum, seasonFactor)
                funtionReturn = addBovineToTable(manifestIdInVar, cowID, cowsNum, cowsWeight, dateManifestIn, fieldIdVar)
            End If
            
            If Not (IsNull(Me.txtYearlings.Value) Or IsNull(Me.txtAvgWeightYearlings.Value) Or IsNull(Me.txtManifestNumber.Value) Or IsNull(Me.txtDateInManifestIn.Value) Or IsNull(Me.comboFieldSelected.Value)) Then
                'manifestIdInVar = Me.txtManifestNumber.Value
                yearlingsID = 5
                yearlingsID = DLookup("[CAT_ID]", "Category", "[CAT_Category] = 'Yearlings'")
                yearlingsNum = Me.txtYearlings.Value
                yearlingsWeight = Me.txtAvgWeightYearlings.Value
                Me.txtQtyYearlingsLast.Value = yearlingsNum
                Me.txtYearlingsWeightLast.Value = yearlingsWeight
                
                metabolicAUMsYearlings = metabolicFunctAUM(yearlingsWeight, yearlingsNum, seasonFactor)
                funtionReturn = addBovineToTable(manifestIdInVar, yearlingsID, yearlingsNum, yearlingsWeight, dateManifestIn, fieldIdVar)
            End If
        
            Dim pairGrazing As Double
            Dim monthGraze As Double
            
            pairGrazing = DLookup("[Quantity]", "ShareFactor", "[Factor] = 'Stocking Rate'")
            monthGraze = DLookup("[Quantity]", "ShareFactor", "[Factor] = 'Basic Season'")
            
            totalAUMs = metabolicAUMsBredHeifers + metabolicAUMsBulls + metabolicAUMsCows + metabolicAUMsCalves + metabolicAUMsYearlings
            
            'This is the variable that holds the total shares needed for the cattle in this manifest
            'We can use it to subtract from the number of shares that the shareholder has
            totalSharesNeeded = Round(totalAUMs / (pairGrazing * monthGraze), 2)
            
            Dim sqlUpdateManifest As String
            sqlUpdateManifest = "UPDATE Manifest " _
                    & "SET MANIFEST_AUMs_Predicted = " & totalAUMs & ", " _
                    & "MANIFEST_Shares_Predicted = " & totalSharesNeeded & " " _
                    & "WHERE MANIFEST_ID = " & manifestIdInVar & " ;"
                    
            DoCmd.SetWarnings False
            DoCmd.RunSQL sqlUpdateManifest
            DoCmd.SetWarnings True
            
        '   Code in case we shouldn't allow negative balance of shares
        '    If ((rsShareholder!SH_Total_Shares - rsShareholder!SH_Shares_Used) >= totalSharesNeeded) Then
        '        rsShareholder.Edit
        '        rsShareholder!SH_Shares_Used = rsShareholder!SH_Shares_Used + totalSharesNeeded
        '        rsShareholder.Update
        '    Else
        '        MsgBox ("There is not enough shares for this shareholder. Please make sure there is enough shares and then proceed.")
        '        Exit Sub
        '    End If
    
            rsShareholder.Edit
            rsShareholder!SH_Shares_Used = rsShareholder!SH_Shares_Used + totalSharesNeeded
            'rsShareholder!SH_Num_Avail_Shares = rsShareholder!SH_Num_Avail_Shares - totalSharesNeeded
            rsShareholder.Update
    
            rsShareholder.Close
            Set rsShareholder = Nothing
            db.Close
            Set db = Nothing
            
                
            Me.txtDateInLast.Value = Me.txtDateInManifestIn.Value
            Me.txtDateOutLast.Value = Me.txtPredictedDateOut.Value
            Me.txtNumDaysLast.Value = Me.txtNumOfDays.Value
            Me.txtFieldLast.Value = Me.comboFieldSelected.Column(1)
            Me.txtShareholderLast.Value = Me.comboShareholder.Column(1)
            Me.txtManifestNumLast.Value = Me.txtManifestNumber.Value
            Me.txtTotalAumLast.Value = totalAUMs
            Me.txtTotalSharesLast.Value = totalSharesNeeded
            Me.txtCommentsLast.Value = manifestComment
            
            ClearAll Me
        End If
    End If
End Sub

Function addBovineToTable(manifestIdInArg As Long, categoryIDarg As Integer, numOfCattleArg As Integer, avgWeightArg As Integer, dateInArg As Date, fieldIdArg As Long) As Integer
    
    Dim db As Database
    Dim rs As DAO.Recordset
    Dim rsLocation As DAO.Recordset
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("Bovine")
    Set rsLocation = db.OpenRecordset("Location")
    
    Dim intCount As Integer
    Dim manifestIdIn As String
    Dim categoryIdVar As Integer
    Dim NrOfCattle As Integer
    Dim avgWeightVar As Integer
    Dim idOfLastCattle As Long
    Dim dateInForLocation As Date
    Dim fieldID As Integer

    manifestIdIn = manifestIdInArg
    categoryIdVar = categoryIDarg
    NrOfCattle = numOfCattleArg - 1
    avgWeightVar = avgWeightArg
    dateInForLocation = dateInArg
    fieldID = fieldIdArg
    
    For intCount = 0 To NrOfCattle
    
        rs.AddNew
        rs("Manifest_ID_In").Value = manifestIdIn
        rs("CAT_ID").Value = categoryIdVar
        rs("BOV_Tag_Number").Value = DMax("[BOV_ID]", "Bovine") + 1
        rs("BOV_Weight").Value = avgWeightVar
        rs("BOV_RFID").Value = DMax("[BOV_ID]", "Bovine") + 1
        rs.Update
        
        'rs.MoveLast
        rs.Move 0, rs.LastModified
        idOfLastCattle = rs!BOV_ID
           
        rsLocation.AddNew
        rsLocation("BOV_ID").Value = idOfLastCattle
        rsLocation("LOC_Date_IN").Value = dateInForLocation
        rsLocation("F_ID").Value = fieldID
        rsLocation.Update
        
    Next intCount
    
    addBovineToTable = 1
    
    rs.Close
    rsLocation.Close
    db.Close
    
    Set rs = Nothing
    Set rsLocation = Nothing
    Set db = Nothing
    
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

Function oldFunctAUM(cattleQtyArg As Integer, cattleFactorArg As Double, seasonFactorArg As Double) As Double
    
    oldFunctAUM = cattleQtyArg * cattleFactorArg * seasonFactorArg
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
    
    'Need to code thyis in later. This are values that will be taken from the lookup table
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

Function ClearAll(frm As Form)
Dim ctl As Control

For Each ctl In frm.Controls
    If ctl.Tag = "txtGroupToClear" Then
        Select Case ctl.ControlType
             Case acTextBox
                ctl.Value = Null
             Case acOptionGroup, acComboBox, acListBox
                ctl.Value = Null
             End Select
    End If
Next
End Function

Function ClearLastTxts(frm As Form)
Dim ctl As Control

For Each ctl In frm.Controls
    If ctl.Tag = "txtLastEntered" Then
        Select Case ctl.ControlType
             Case acTextBox
                ctl.Value = Null
             Case acOptionGroup, acComboBox, acListBox
                ctl.Value = Null
             End Select
    End If
Next
End Function

Function checkDates(startDateArg As Date, endDateArg As Date) As Integer
    If (endDateArg <= startDateArg) Then
    checkDates = 0
    Call MsgBox("The End Date is less than or equal to Start Date. Please enter correct dates.", , "CALMS")
    ElseIf (endDateArg > startDateArg) Then
    checkDates = 1
    End If
End Function

Private Sub buttonCalculateShares_Click()
    
    
    Dim manifestNumberInVar As String
    Dim manifestIdInVar As Integer ' should check this if it needs to be declared as a Long
    Dim funtionReturn As Integer
    Dim dateManifestIn As Date
    Dim planedDateOut As Date
    Dim fieldIdVar As Integer
    Dim shareholderID As Integer
    Dim manifestComment As String
    Dim promptForManifestIn As String
    
    Dim metabolicAUMsBredHeifers As Double
    Dim metabolicAUMsBulls As Double
    Dim metabolicAUMsCows As Double
    Dim metabolicAUMsCalves As Double
    Dim metabolicAUMsYearlings As Double
    Dim totalAUMs As Double
    Dim totalSharesNeeded As Double
    Dim seasonFactor As Double
    metabolicAUMsBredHeifers = 0: metabolicAUMsBulls = 0: metabolicAUMsCows = 0: metabolicAUMsCalves = 0: metabolicAUMsYearlings = 0: totalAUMs = 0: totalSharesNeeded = 0
    
    Dim bredHeiferID As Integer: Dim bredHeiferNum As Integer: Dim bredHeiferWeight As Integer
    Dim bullsID As Integer: Dim bullsNum As Integer: Dim bullsWeight As Integer
    Dim cowID As Integer: Dim cowsNum As Integer: Dim cowsWeight As Integer
    Dim calvesID As Integer: Dim calvesNum As Integer: Dim calvesWeight As Integer
    Dim yearlingsID As Integer: Dim yearlingsNum As Integer: Dim yearlingsWeight As Integer
    
    If (IsNull(Me.txtManifestNumber.Value) Or IsNull(Me.txtDateInManifestIn.Value) Or IsNull(Me.txtPredictedDateOut.Value) Or IsNull(Me.comboFieldSelected.Value) Or IsNull(Me.comboShareholder.Value)) Then
        Call MsgBox("Please enter all the required info: Shareholder, Date In, Date Out, Manifest Number, and Field.", , "CALMS")
        Exit Sub
    Else
    
    Dim db As DAO.Database
    Dim rsManifest As DAO.Recordset
    Dim rsShareholder As DAO.Recordset
    Dim sqlSharesAvailable As String
    sqlSharesAvailable = "SELECT * from Shareholder WHERE SH_ID = (" & Me.comboShareholder.Column(0) & ");"
    
    
    Set db = CurrentDb
    Set rsManifest = db.OpenRecordset("Manifest")
    Set rsShareholder = db.OpenRecordset(sqlSharesAvailable)
    
    dateManifestIn = Me.txtDateInManifestIn.Value
    planedDateOut = Me.txtPredictedDateOut.Value
    fieldIdVar = Me.comboFieldSelected.Column(0)
    shareholderID = Me.comboShareholder.Column(0)
    manifestNumberInVar = Me.txtManifestNumber.Value
    
    If (checkDates(dateManifestIn, planedDateOut) = 0) Then
        Exit Sub
    End If
    
    
    
    seasonFactor = seasonalFactorTotal(dateManifestIn, planedDateOut)
    
    If Not (IsNull(Me.txtBredHeifers.Value) Or IsNull(Me.txtAvgWeightBredHifers.Value)) Then
        bredHeiferNum = Me.txtBredHeifers.Value
        bredHeiferWeight = Me.txtAvgWeightBredHifers.Value
        metabolicAUMsBredHeifers = metabolicFunctAUM(bredHeiferWeight, bredHeiferNum, seasonFactor)
    End If
    If Not (IsNull(Me.txtBulls.Value) Or IsNull(Me.txtAvgWeightBulls.Value)) Then
        bullsNum = Me.txtBulls.Value
        bullsWeight = Me.txtAvgWeightBulls.Value
        metabolicAUMsBulls = metabolicFunctAUM(bullsWeight, bullsNum, seasonFactor)
    End If
    If Not (IsNull(Me.txtCalves.Value) Or IsNull(Me.txtAvgWeightCalves.Value)) Then
        calvesNum = Me.txtCalves.Value
        calvesWeight = Me.txtAvgWeightCalves.Value
        metabolicAUMsCalves = metabolicFunctAUM(calvesWeight, calvesNum, seasonFactor)
    End If
    If Not (IsNull(Me.txtCows.Value) Or IsNull(Me.txtAvgWeightCows.Value)) Then
        cowsNum = Me.txtCows.Value
        cowsWeight = Me.txtAvgWeightCows.Value
        metabolicAUMsCows = metabolicFunctAUM(cowsWeight, cowsNum, seasonFactor)
    End If
    If Not (IsNull(Me.txtYearlings.Value) Or IsNull(Me.txtAvgWeightYearlings.Value)) Then
        yearlingsNum = Me.txtYearlings.Value
        yearlingsWeight = Me.txtAvgWeightYearlings.Value
        metabolicAUMsYearlings = metabolicFunctAUM(yearlingsWeight, yearlingsNum, seasonFactor)
    End If
    
    
    Dim pairGrazing As Double
    Dim monthGraze As Double
    
    pairGrazing = DLookup("[Quantity]", "ShareFactor", "[Factor] = 'Stocking Rate'")
    monthGraze = DLookup("[Quantity]", "ShareFactor", "[Factor] = 'Basic Season'")
    
    totalAUMs = metabolicAUMsBredHeifers + metabolicAUMsBulls + metabolicAUMsCows + metabolicAUMsCalves + metabolicAUMsYearlings
    
    'This is the variable that holds the total shares needed for the cattle in this manifest
    'We can use it to subtract from the number of shares that the shareholder has
    totalSharesNeeded = Round(totalAUMs / (pairGrazing * monthGraze), 2)
    
    Dim subtractShares As Double
    Dim addShares As Double
    'this is DSum function for values from the SharesTransfer table
    addShares = Nz(DSum("ST_Shares_Number", "SharesTransfer", "ST_Controller = " & Me.comboShareholder.Column(0) & ""), 0)
    subtractShares = Nz(DSum("ST_Shares_Number", "SharesTransfer", "ST_Owner = " & Me.comboShareholder.Column(0) & ""), 0)
    
    'code that shows the shares available, two ways to show the shares available based on ManifestsIn
    'Me.txtSharesAvailble.Value = -subtractShares + rsShareholder!SH_Total_Shares + addShares - rsShareholder!SH_Shares_Used
    Me.txtSharesAvailble.Value = rsShareholder!SH_Total_Shares + rsShareholder!SH_Transfers_Balance - rsShareholder!SH_Shares_Used
    
    Me.txtAUMsNeeded.Value = totalAUMs
    Me.txtSharesNeeded.Value = totalSharesNeeded
    
    rsShareholder.Close
    rsManifest.Close
    Set rsShareholder = Nothing
    Set rsManifest = Nothing
    db.Close
    Set db = Nothing
    
    End If
End Sub

Private Sub buttonClearAll_Click()
    ClearAll Me
End Sub



Private Sub txtAvgWeightBredHifers_KeyPress(KeyAscii As Integer)
    LimitNumeric KeyAscii
End Sub

Private Sub txtAvgWeightBulls_KeyPress(KeyAscii As Integer)
    LimitNumeric KeyAscii
End Sub

Private Sub txtAvgWeightCalves_KeyPress(KeyAscii As Integer)
    LimitNumeric KeyAscii
End Sub

Private Sub txtAvgWeightCows_KeyPress(KeyAscii As Integer)
    LimitNumeric KeyAscii
End Sub

Private Sub txtAvgWeightYearlings_KeyPress(KeyAscii As Integer)
    LimitNumeric KeyAscii
End Sub

Private Sub txtBredHeifers_KeyPress(KeyAscii As Integer)
    LimitNumeric KeyAscii
End Sub

Private Sub txtBulls_KeyPress(KeyAscii As Integer)
    LimitNumeric KeyAscii
End Sub

Private Sub txtCalves_KeyPress(KeyAscii As Integer)
    LimitNumeric KeyAscii
End Sub

Private Sub txtCows_KeyPress(KeyAscii As Integer)
    LimitNumeric KeyAscii
End Sub




Private Sub txtDateInManifestIn_AfterUpdate()
    If checkMonthAprilMay(Me.txtDateInManifestIn) = False Then
        Me.txtDateInManifestIn.SetFocus
        Exit Sub
    End If
End Sub

'Private Sub txtDateInManifestIn_Change()
'    If checkMonthAprilMay(Me.txtDateInManifestIn) = False Then
'        Me.txtDateInManifestIn.SetFocus
'        Exit Sub
'    End If
'End Sub
'
'Private Sub txtDateInManifestIn_Exit(Cancel As Integer)
'    If checkMonthAprilMay(Me.txtDateInManifestIn) = False Then
'        Me.txtDateInManifestIn.SetFocus
'        Exit Sub
'    End If
'End Sub

'Private Sub txtDateInManifestIn_LostFocus()
'    If checkMonthAprilMay(Me.txtDateInManifestIn) = False Then
'        Me.txtDateInManifestIn.SetFocus
'        Exit Sub
'    End If
'End Sub

Private Sub txtManifestNumber_KeyPress(KeyAscii As Integer)
    LimitFieldSize KeyAscii, 15
    LimitAlphanumeric KeyAscii
End Sub

Private Sub txtPredictedDateOut_AfterUpdate()
    If checkMonthAprilMay(Me.txtPredictedDateOut) = False Then
        Exit Sub
    End If
    
    If (Not IsNull(Me.txtPredictedDateOut.Value) And Not IsNull(Me.txtDateInManifestIn.Value)) Then
        Me.txtNumOfDays = DateDiff("d", Me.txtDateInManifestIn.Value, Me.txtPredictedDateOut.Value) + 1
    End If
End Sub

Private Sub txtYearlings_KeyPress(KeyAscii As Integer)
    LimitNumeric KeyAscii
End Sub
