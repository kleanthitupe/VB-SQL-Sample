Option Compare Database
    'global variables
    Public oldAumTotal As Double
    Public metabolicAumTotal As Double
    Public mSharesTotal As Double
    Option Explicit
Function ClearLastCalculated(frm As Form)
Dim ctl As Control

For Each ctl In frm.Controls
    If ctl.Tag = "clearCalculated" Then
        Select Case ctl.ControlType
             Case acTextBox
                ctl.Value = Null
             Case acOptionGroup, acComboBox, acListBox
                ctl.Value = Null
             End Select
    End If
Next
End Function
    

'button that calculates the AUMs based on the entries by the user
Private Sub buttonCalculateAUM_Click()

    ClearLastCalculated Me
                        
    Dim sharesAvailableVar As Double
    Dim allocatedAUMvar As Double
    Dim metabolicAUMsLeftVar As Double
    Dim sharesLeftMetabolicVar As Double
    sharesAvailableVar = 0
    
    Dim lbCowVar As Double
    Dim lbCalfVar As Double
    Dim pairGrazing As Double
    Dim monthGraze As Double
    
    'These values are taken from the lookup table, in order to keep it updatable by the user in the future
    lbCowVar = DLookup("[Quantity]", "ShareFactor", "[Factor] = 'Lb. Cow'")
    lbCalfVar = DLookup("[Quantity]", "ShareFactor", "[Factor] = 'Lb. Calf'")
    pairGrazing = DLookup("[Quantity]", "ShareFactor", "[Factor] = 'Stocking Rate'")
    monthGraze = DLookup("[Quantity]", "ShareFactor", "[Factor] = 'Basic Season'")
    
    'check to see that the user entered a share amount, and prompt him if not
    If Len(Me.sharesOwned.Value & "") = 0 Then
        Call MsgBox("Please enter the amount of shares available!", , "CALMS")
        Exit Sub
    End If
    
    
    'set global variables to zero
    metabolicAumTotal = 0: mSharesTotal = 0
    
    'check dates first, to see if there is any where that the end date is less than the start date
    If checkDatesFirst("Row 1", Me.txtStartDate1, Me.txtEndDate1) = False Then
        GoTo bailOut
    End If

    If checkDatesFirst("Row 2", Me.txtStartDate2, Me.txtEndDate2) = False Then
        GoTo bailOut
    End If
    
    If checkDatesFirst("Row 3", Me.txtStartDate3, Me.txtEndDate3) = False Then
        GoTo bailOut
    End If
    
    If checkDatesFirst("Row 4", Me.txtStartDate4, Me.txtEndDate4) = False Then
        GoTo bailOut
    End If
    
    If checkDatesFirst("Row 5", Me.txtStartDate5, Me.txtEndDate5) = False Then
        GoTo bailOut
    End If
    
    If checkDatesFirst("Row 6", Me.txtStartDate6, Me.txtEndDate6) = False Then
        GoTo bailOut
    End If
    
    If checkDatesFirst("Row 7", Me.txtStartDate7, Me.txtEndDate7) = False Then
        GoTo bailOut
    End If
    
    If checkDatesFirst("Row 8", Me.txtStartDate8, Me.txtEndDate8) = False Then
        GoTo bailOut
    End If
    
    If checkDatesFirst("Row 9", Me.txtStartDate9, Me.txtEndDate9) = False Then
        GoTo bailOut
    End If
    
    If checkDatesFirst("Row 10", Me.txtStartDate10, Me.txtEndDate10) = False Then
        GoTo bailOut
    End If
    
    If checkDatesFirst("Row 3", Me.txtStartDate2, Me.txtEndDate3) = False Then
        GoTo bailOut
    End If
    
    'call the function that calculates and sets the AUMs for each row. Arguments are passed as controls
    Call setAumFields("Row 1", Me.txtStartDate1, Me.txtEndDate1, Me.txtAvgWeight1, Me.txtQtyOfCattle1, Me.txtCattleFactor1, Me.txtNumDaysRow1, Me.txtMetabolicAUMrow1, Me.txtMsharesRow1, pairGrazing, monthGraze)
    Call setAumFields("Row 2", Me.txtStartDate2, Me.txtEndDate2, Me.txtAvgWeight2, Me.txtQtyOfCattle2, Me.txtCattleFactor2, Me.txtNumDaysRow2, Me.txtMetabolicAUMrow2, Me.txtMsharesRow2, pairGrazing, monthGraze)
    Call setAumFields("Row 3", Me.txtStartDate3, Me.txtEndDate3, Me.txtAvgWeight3, Me.txtQtyOfCattle3, Me.txtCattleFactor3, Me.txtNumDaysRow3, Me.txtMetabolicAUMrow3, Me.txtMsharesRow3, pairGrazing, monthGraze)
    Call setAumFields("Row 4", Me.txtStartDate4, Me.txtEndDate4, Me.txtAvgWeight4, Me.txtQtyOfCattle4, Me.txtCattleFactor4, Me.txtNumDaysRow4, Me.txtMetabolicAUMrow4, Me.txtMsharesRow4, pairGrazing, monthGraze)
    Call setAumFields("Row 5", Me.txtStartDate5, Me.txtEndDate5, Me.txtAvgWeight5, Me.txtQtyOfCattle5, Me.txtCattleFactor5, Me.txtNumDaysRow5, Me.txtMetabolicAUMrow5, Me.txtMsharesRow5, pairGrazing, monthGraze)
    Call setAumFields("Row 6", Me.txtStartDate6, Me.txtEndDate6, Me.txtAvgWeight6, Me.txtQtyOfCattle6, Me.txtCattleFactor6, Me.txtNumDaysRow6, Me.txtMetabolicAUMrow6, Me.txtMsharesRow6, pairGrazing, monthGraze)
    Call setAumFields("Row 7", Me.txtStartDate7, Me.txtEndDate7, Me.txtAvgWeight7, Me.txtQtyOfCattle7, Me.txtCattleFactor7, Me.txtNumDaysRow7, Me.txtMetabolicAUMrow7, Me.txtMsharesRow7, pairGrazing, monthGraze)
    Call setAumFields("Row 8", Me.txtStartDate8, Me.txtEndDate8, Me.txtAvgWeight8, Me.txtQtyOfCattle8, Me.txtCattleFactor8, Me.txtNumDaysRow8, Me.txtMetabolicAUMrow8, Me.txtMsharesRow8, pairGrazing, monthGraze)
    Call setAumFields("Row 9", Me.txtStartDate9, Me.txtEndDate9, Me.txtAvgWeight9, Me.txtQtyOfCattle9, Me.txtCattleFactor9, Me.txtNumDaysRow9, Me.txtMetabolicAUMrow9, Me.txtMsharesRow9, pairGrazing, monthGraze)
    Call setAumFields("Row 10", Me.txtStartDate10, Me.txtEndDate10, Me.txtAvgWeight10, Me.txtQtyOfCattle10, Me.txtCattleFactor10, Me.txtNumDaysRow10, Me.txtMetabolicAUMrow10, Me.txtMsharesRow10, pairGrazing, monthGraze)

    
    
    sharesAvailableVar = Me.sharesOwned.Value
    
    'these are two textboxes that get their values from the global variables, which change only from the function that is called above 10 times, one for each row
    Me.txtMsharesTotal.Value = mSharesTotal
    Me.txtMetabolicAUMtotal.Value = metabolicAumTotal
    
    
    'Note: 1 Share eaqual 4.2 pairs grazing for 5 months.
    'Calculate the values of these variables
    allocatedAUMvar = sharesAvailableVar * pairGrazing * monthGraze
    metabolicAUMsLeftVar = allocatedAUMvar - metabolicAumTotal
    sharesLeftMetabolicVar = metabolicAUMsLeftVar / (pairGrazing * monthGraze)
    
    
    Me.allocatedAUM.Value = allocatedAUMvar
    Me.metabolicUsed.Value = metabolicAumTotal
    Me.metabolicLeft.Value = metabolicAUMsLeftVar
    Me.sharesLeftMetabolic.Value = sharesLeftMetabolicVar
    Me.percentMetabolic.Value = metabolicAUMsLeftVar / allocatedAUMvar
         
    Call MsgBox("Based on the data entered, the calculations are completed! Check results.", , "CALMS")
    metabolicAumTotal = 0: mSharesTotal = 0
    
bailOut:
End Sub

'This code is triggered with the Optimize button, which calculates the max number of each cattle category based on shares available
Private Sub buttonM_Click()

 metabolicAumTotal = 0: mSharesTotal = 0
    'This are values that will be taken from the lookup table
    Dim lbCowVar As Double
    Dim lbCalfVar As Double
    Dim pairGrazing As Double
    Dim monthGraze As Double
    
    'initialize these values from the ShareFactor table, which can be updated from the user
    lbCowVar = DLookup("[Quantity]", "ShareFactor", "[Factor] = 'Lb. Cow'")
    lbCalfVar = DLookup("[Quantity]", "ShareFactor", "[Factor] = 'Lb. Calf'")
    pairGrazing = DLookup("[Quantity]", "ShareFactor", "[Factor] = 'Stocking Rate'")
    monthGraze = DLookup("[Quantity]", "ShareFactor", "[Factor] = 'Basic Season'")
  
    Dim mSharesRow1Var As Double: Dim mSharesRow2Var As Double: Dim mSharesRow3Var As Double: Dim mSharesRow4Var As Double: Dim mSharesRow5Var As Double:
    Dim mSharesRow6Var As Double: Dim mSharesRow7Var As Double: Dim mSharesRow8Var As Double: Dim mSharesRow9Var As Double: Dim mSharesRow10Var As Double:
    Dim sharesAvailableVar As Double
    Dim qtyRow1 As Double: Dim qtyRow2 As Double: Dim qtyRow3 As Double: Dim qtyRow4 As Double: Dim qtyRow5 As Double
    Dim qtyRow6 As Double: Dim qtyRow7 As Double: Dim qtyRow8 As Double: Dim qtyRow9 As Double: Dim qtyRow10 As Double
    Dim shareTotalOf10Rows As Double

    'initialize variables to zero
    qtyRow1 = 0: qtyRow2 = 0: qtyRow3 = 0: qtyRow4 = 0: qtyRow5 = 0: qtyRow6 = 0: qtyRow7 = 0: qtyRow8 = 0: qtyRow9 = 0: qtyRow10 = 0:
    sharesAvailableVar = 0
   
    ClearLastCalculated Me
    
    'check to see that the user entered an amount for shares available. Prompt the user if not
    If Len(Me.sharesOwned.Value & "") = 0 Then
        Call MsgBox("Please enter the amount of shares available!", , "CALMS")
        Exit Sub
    End If
    
        'check dates first, to see if there is any where that the end date is less than the start date
    If checkDatesFirst("Row 1", Me.txtStartDate1, Me.txtEndDate1) = False Then
        GoTo bailOut
    End If

    If checkDatesFirst("Row 2", Me.txtStartDate2, Me.txtEndDate2) = False Then
        GoTo bailOut
    End If
    
    If checkDatesFirst("Row 3", Me.txtStartDate3, Me.txtEndDate3) = False Then
        GoTo bailOut
    End If
    
    If checkDatesFirst("Row 4", Me.txtStartDate4, Me.txtEndDate4) = False Then
        GoTo bailOut
    End If
    
    If checkDatesFirst("Row 5", Me.txtStartDate5, Me.txtEndDate5) = False Then
        GoTo bailOut
    End If
    
    If checkDatesFirst("Row 6", Me.txtStartDate6, Me.txtEndDate6) = False Then
        GoTo bailOut
    End If
    
    If checkDatesFirst("Row 7", Me.txtStartDate7, Me.txtEndDate7) = False Then
        GoTo bailOut
    End If
    
    If checkDatesFirst("Row 8", Me.txtStartDate8, Me.txtEndDate8) = False Then
        GoTo bailOut
    End If
    
    If checkDatesFirst("Row 9", Me.txtStartDate9, Me.txtEndDate9) = False Then
        GoTo bailOut
    End If
    
    If checkDatesFirst("Row 10", Me.txtStartDate10, Me.txtEndDate10) = False Then
        GoTo bailOut
    End If
    
    If checkDatesFirst("Row 3", Me.txtStartDate2, Me.txtEndDate3) = False Then
        GoTo bailOut
    End If
    
        
        mSharesRow1Var = calcSharesFor1Cattle("Row1", Me.txtCattleFactor1, Me.txtAvgWeight1, Me.txtStartDate1, Me.txtEndDate1, Me.txtNumDaysRow1, pairGrazing, monthGraze)
        mSharesRow2Var = calcSharesFor1Cattle("Row2", Me.txtCattleFactor2, Me.txtAvgWeight2, Me.txtStartDate2, Me.txtEndDate2, Me.txtNumDaysRow2, pairGrazing, monthGraze)
        mSharesRow3Var = calcSharesFor1Cattle("Row3", Me.txtCattleFactor3, Me.txtAvgWeight3, Me.txtStartDate3, Me.txtEndDate3, Me.txtNumDaysRow3, pairGrazing, monthGraze)
        mSharesRow4Var = calcSharesFor1Cattle("Row4", Me.txtCattleFactor4, Me.txtAvgWeight4, Me.txtStartDate4, Me.txtEndDate4, Me.txtNumDaysRow4, pairGrazing, monthGraze)
        mSharesRow5Var = calcSharesFor1Cattle("Row5", Me.txtCattleFactor5, Me.txtAvgWeight5, Me.txtStartDate5, Me.txtEndDate5, Me.txtNumDaysRow5, pairGrazing, monthGraze)
        mSharesRow6Var = calcSharesFor1Cattle("Row6", Me.txtCattleFactor6, Me.txtAvgWeight6, Me.txtStartDate6, Me.txtEndDate6, Me.txtNumDaysRow6, pairGrazing, monthGraze)
        mSharesRow7Var = calcSharesFor1Cattle("Row7", Me.txtCattleFactor7, Me.txtAvgWeight7, Me.txtStartDate7, Me.txtEndDate7, Me.txtNumDaysRow7, pairGrazing, monthGraze)
        mSharesRow8Var = calcSharesFor1Cattle("Row8", Me.txtCattleFactor8, Me.txtAvgWeight8, Me.txtStartDate8, Me.txtEndDate8, Me.txtNumDaysRow8, pairGrazing, monthGraze)
        mSharesRow9Var = calcSharesFor1Cattle("Row9", Me.txtCattleFactor9, Me.txtAvgWeight9, Me.txtStartDate9, Me.txtEndDate9, Me.txtNumDaysRow9, pairGrazing, monthGraze)
        mSharesRow10Var = calcSharesFor1Cattle("Row10", Me.txtCattleFactor10, Me.txtAvgWeight10, Me.txtStartDate10, Me.txtEndDate10, Me.txtNumDaysRow10, pairGrazing, monthGraze)
        
        sharesAvailableVar = Me.sharesOwned.Value
        Me.allocatedAUM.Value = sharesAvailableVar * pairGrazing * monthGraze
        shareTotalOf10Rows = mSharesRow1Var + mSharesRow2Var + mSharesRow3Var + mSharesRow4Var + mSharesRow5Var + mSharesRow6Var + mSharesRow7Var + mSharesRow8Var + mSharesRow9Var + mSharesRow10Var
              
        If (shareTotalOf10Rows) > 0 Then
              
        While (sharesAvailableVar >= shareTotalOf10Rows * 0.49) 'this while loop runs as long as shares available is bigger than
            qtyRow1 = qtyRow1 + 1
            qtyRow2 = qtyRow2 + 1
            qtyRow3 = qtyRow3 + 1
            qtyRow4 = qtyRow4 + 1
            qtyRow5 = qtyRow5 + 1
            qtyRow6 = qtyRow6 + 1
            qtyRow7 = qtyRow7 + 1
            qtyRow8 = qtyRow8 + 1
            qtyRow9 = qtyRow9 + 1
            qtyRow10 = qtyRow10 + 1
            sharesAvailableVar = sharesAvailableVar - shareTotalOf10Rows
        Wend
        
        'function that checks if the row has the right data
        Call ifRowHasData(mSharesRow1Var, qtyRow1, Me.txtQtyOfCattle1)
        Call ifRowHasData(mSharesRow2Var, qtyRow2, Me.txtQtyOfCattle2)
        Call ifRowHasData(mSharesRow3Var, qtyRow3, Me.txtQtyOfCattle3)
        Call ifRowHasData(mSharesRow4Var, qtyRow4, Me.txtQtyOfCattle4)
        Call ifRowHasData(mSharesRow5Var, qtyRow5, Me.txtQtyOfCattle5)
        Call ifRowHasData(mSharesRow6Var, qtyRow6, Me.txtQtyOfCattle6)
        Call ifRowHasData(mSharesRow7Var, qtyRow7, Me.txtQtyOfCattle7)
        Call ifRowHasData(mSharesRow8Var, qtyRow8, Me.txtQtyOfCattle8)
        Call ifRowHasData(mSharesRow9Var, qtyRow9, Me.txtQtyOfCattle9)
        Call ifRowHasData(mSharesRow10Var, qtyRow10, Me.txtQtyOfCattle10)
        
        Call setAumFields("Row 1", Me.txtStartDate1, Me.txtEndDate1, Me.txtAvgWeight1, Me.txtQtyOfCattle1, Me.txtCattleFactor1, Me.txtNumDaysRow1, Me.txtMetabolicAUMrow1, Me.txtMsharesRow1, pairGrazing, monthGraze)
        Call setAumFields("Row 2", Me.txtStartDate2, Me.txtEndDate2, Me.txtAvgWeight2, Me.txtQtyOfCattle2, Me.txtCattleFactor2, Me.txtNumDaysRow2, Me.txtMetabolicAUMrow2, Me.txtMsharesRow2, pairGrazing, monthGraze)
        Call setAumFields("Row 3", Me.txtStartDate3, Me.txtEndDate3, Me.txtAvgWeight3, Me.txtQtyOfCattle3, Me.txtCattleFactor3, Me.txtNumDaysRow3, Me.txtMetabolicAUMrow3, Me.txtMsharesRow3, pairGrazing, monthGraze)
        Call setAumFields("Row 4", Me.txtStartDate4, Me.txtEndDate4, Me.txtAvgWeight4, Me.txtQtyOfCattle4, Me.txtCattleFactor4, Me.txtNumDaysRow4, Me.txtMetabolicAUMrow4, Me.txtMsharesRow4, pairGrazing, monthGraze)
        Call setAumFields("Row 5", Me.txtStartDate5, Me.txtEndDate5, Me.txtAvgWeight5, Me.txtQtyOfCattle5, Me.txtCattleFactor5, Me.txtNumDaysRow5, Me.txtMetabolicAUMrow5, Me.txtMsharesRow5, pairGrazing, monthGraze)
        Call setAumFields("Row 6", Me.txtStartDate6, Me.txtEndDate6, Me.txtAvgWeight6, Me.txtQtyOfCattle6, Me.txtCattleFactor6, Me.txtNumDaysRow6, Me.txtMetabolicAUMrow6, Me.txtMsharesRow6, pairGrazing, monthGraze)
        Call setAumFields("Row 7", Me.txtStartDate7, Me.txtEndDate7, Me.txtAvgWeight7, Me.txtQtyOfCattle7, Me.txtCattleFactor7, Me.txtNumDaysRow7, Me.txtMetabolicAUMrow7, Me.txtMsharesRow7, pairGrazing, monthGraze)
        Call setAumFields("Row 8", Me.txtStartDate8, Me.txtEndDate8, Me.txtAvgWeight8, Me.txtQtyOfCattle8, Me.txtCattleFactor8, Me.txtNumDaysRow8, Me.txtMetabolicAUMrow8, Me.txtMsharesRow8, pairGrazing, monthGraze)
        Call setAumFields("Row 9", Me.txtStartDate9, Me.txtEndDate9, Me.txtAvgWeight9, Me.txtQtyOfCattle9, Me.txtCattleFactor9, Me.txtNumDaysRow9, Me.txtMetabolicAUMrow9, Me.txtMsharesRow9, pairGrazing, monthGraze)
        Call setAumFields("Row 10", Me.txtStartDate10, Me.txtEndDate10, Me.txtAvgWeight10, Me.txtQtyOfCattle10, Me.txtCattleFactor10, Me.txtNumDaysRow10, Me.txtMetabolicAUMrow10, Me.txtMsharesRow10, pairGrazing, monthGraze)
        
        Me.txtMsharesTotal.Value = mSharesTotal
        Me.txtMetabolicAUMtotal.Value = metabolicAumTotal
        Me.metabolicUsed.Value = Null: Me.metabolicLeft.Value = Null: Me.sharesLeftMetabolic.Value = Null: Me.percentMetabolic.Value = Null
        
        metabolicAumTotal = 0: mSharesTotal = 0
        Else
            Exit Sub
        End If
bailOut:
End Sub
'function that checks if the row is filled with data appropriatele and then writes the quantity of cattle optimized to the appropriate textbox
Function ifRowHasData(mSharesCalculated As Double, cattleQtyCalculated As Double, txtCattleQty As Control) As Boolean
    If (mSharesCalculated > 0) Then
        txtCattleQty.Value = cattleQtyCalculated
        ifRowHasData = True
    Else
        txtCattleQty.Value = Null
        ifRowHasData = False
    End If
End Function


'function that clears all the controls in the form
Function ClearAll(frm As Form)
Dim ctl As Control

For Each ctl In frm.Controls
   Select Case ctl.ControlType
      Case acTextBox
           ctl.Value = Null
      Case acOptionGroup, acComboBox, acListBox
          ctl.Value = Null
   End Select
Next
End Function

'function that checks if the end date is less than the start date
Function checkDates(startDateArg As Date, endDateArg As Date) As Integer
    If (endDateArg <= startDateArg) Then
    checkDates = 0
    
    ElseIf (endDateArg > startDateArg) Then
    checkDates = 1

    End If
End Function

'after the user selects a shareholder, then the shares available is updated
Private Sub comboShareholderName_AfterUpdate()
    Dim db As DAO.Database
    Dim rsShareholder As Recordset
    Dim sqlSharesAvailable As String
    'sqlSharesAvailable = "SELECT [SH_Shares_Used], [SH_Total_Shares], [SH_Carryover], [SH_Transfers_Balance]  from Shareholder WHERE SH_ID = (" & Me.comboShareholderName.Column(0) & ");"
    sqlSharesAvailable = "SELECT *  from Shareholder WHERE SH_ID = (" & Me.comboShareholderName.Column(0) & ");"

    
    Set db = CurrentDb
    Set rsShareholder = db.OpenRecordset(sqlSharesAvailable)
    
    'Me.sharesOwned.Value = (rsShareholder!SH_Total_Shares + rsShareholder!SH_Carryover + rsShareholder!SH_Transfers_Balance - rsShareholder!SH_Shares_Used)
    Me.sharesOwned.Value = rsShareholder!SH_Total_Shares + rsShareholder!SH_Transfers_Balance - rsShareholder!SH_Shares_Used
 
     
    rsShareholder.Close
    Set rsShareholder = Nothing
    db.Close
    Set db = Nothing
End Sub

'this subsections update the value of the cattle factor based on the value of the combo box selection
Public Sub comboCattle1_Change()
    Me.txtCattleFactor1.Value = Me.comboCattle1.Column(2)
End Sub

Private Sub comboCattle2_Change()
    Me.txtCattleFactor2.Value = Me.comboCattle2.Column(2)
End Sub

Private Sub comboCattle3_Change()
    Me.txtCattleFactor3.Value = Me.comboCattle3.Column(2)
End Sub

Private Sub comboCattle4_Change()
     Me.txtCattleFactor4.Value = Me.comboCattle4.Column(2)
End Sub

Private Sub comboCattle5_Change()
     Me.txtCattleFactor5.Value = Me.comboCattle5.Column(2)
End Sub

Private Sub comboCattle6_Change()
     Me.txtCattleFactor6.Value = Me.comboCattle6.Column(2)
End Sub

Private Sub comboCattle7_Change()
     Me.txtCattleFactor7.Value = Me.comboCattle7.Column(2)
End Sub

Private Sub comboCattle8_Change()
     Me.txtCattleFactor8.Value = Me.comboCattle8.Column(2)
End Sub

Private Sub comboCattle9_Change()
     Me.txtCattleFactor9.Value = Me.comboCattle9.Column(2)
End Sub

Private Sub comboCattle10_Change()
    Me.txtCattleFactor10.Value = Me.comboCattle10.Column(2)
End Sub

Function checkDatesFirst(rowX As String, sDate As Control, eDate As Control) As Boolean

    If Not (IsNull(sDate.Value) Or IsNull(eDate.Value)) Then
        
            Dim dateInRowVar As Date
            Dim dateOutRowVar As Date
            
            dateInRowVar = sDate.Value
            dateOutRowVar = eDate.Value
    
            If checkDates(dateInRowVar, dateOutRowVar) = 1 Then
                checkDatesFirst = True
            Else
                checkDatesFirst = False
                Call MsgBox("The " & rowX & " end date is less than or equal to the start date. Please enter correct dates", , "CALMS")
                Exit Function
            End If
    ElseIf (IsNull(sDate.Value) Or IsNull(eDate.Value)) Then
         checkDatesFirst = True
    End If
    
End Function


Sub setAumFields(rowX As String, sDate As Control, eDate As Control, avgWeight As Control, qtyCattle As Control, cattleFactor As Control, dateDifference As Control, metabolicAumRowX As Control, mSharesRowX As Control, pairGraze As Double, monthsGraze As Double)
        
        If Not (IsNull(sDate.Value) Or IsNull(eDate.Value) Or IsNull(avgWeight.Value) Or IsNull(qtyCattle.Value) Or IsNull(cattleFactor.Value)) Then
        
            Dim dateInRowVar As Date
            Dim dateOutRowVar As Date
            Dim avgWeightRowVar As Double
            Dim cattleQtyRowVar As Double
            Dim cattleFactorRowVar As Double
        
            Dim seasonFactorRowVar As Double
            Dim metabolicAUMrowVar As Double
            Dim mSharesRowVar As Double
        
            'metabolicAumTotal = 0: mSharesTotal = 0
        
        
            dateInRowVar = sDate.Value
            dateOutRowVar = eDate.Value
            avgWeightRowVar = avgWeight.Value
            cattleQtyRowVar = qtyCattle.Value
            cattleFactorRowVar = cattleFactor.Value
    
            If checkDates(dateInRowVar, dateOutRowVar) = 1 Then
        
                seasonFactorRowVar = seasonalFactorTotal(dateInRowVar, dateOutRowVar)
                metabolicAUMrowVar = metabolicFunctAUM(avgWeightRowVar, cattleQtyRowVar, seasonFactorRowVar)
                
                'shares used  = metabolic AUMs/ (4.2 pairs * grazing for 5 months)
                mSharesRowVar = metabolicAUMrowVar / (pairGraze * monthsGraze)
                
                dateDifference.Value = DateDiff("d", dateInRowVar, dateOutRowVar) + 1
        
                'set the value of the fields
                metabolicAumRowX.Value = metabolicAUMrowVar
                mSharesRowX.Value = mSharesRowVar
                
                'global variables that store the totals
                metabolicAumTotal = metabolicAumTotal + metabolicAUMrowVar
                mSharesTotal = mSharesTotal + mSharesRowVar
        
                'setAumFields = True
                
            Else:
                'setAumFields = False
                Call MsgBox("The " & rowX & " end date is less than or equal to the start date. Please enter correct dates", , "CALMS")
                GoTo bailOut
            End If
        
        End If
bailOut:
End Sub


Private Sub txtEndDate1_AfterUpdate()
    If (Me.txtEndDate3.Visible = False) Then
        Me.txtEndDate2.Value = Me.txtEndDate1.Value
    End If
End Sub

Private Sub txtStartDate1_AfterUpdate()
    If (Me.txtStartDate3.Visible = False) Then
        Me.txtStartDate2.Value = Me.txtStartDate1.Value
    End If
End Sub

'function that calculates the metabolic AUM, when given the average weight, cattle quantity, and the season factor
Function metabolicFunctAUM(averageWeightArg As Double, cattleQtyArg As Double, seasonFactorArg As Double) As Double
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
    'MsgBox ("The metabolic AUMs are: " & metabolicFunctAUM)
    
End Function


Function oldFunctAUM(cattleQtyArg As Double, cattleFactorArg As Double, seasonFactorArg As Double) As Double
    oldFunctAUM = cattleQtyArg * cattleFactorArg * seasonFactorArg
End Function


'function that calculates the seasonal factor given the start date and end date
Function seasonalFactorTotal(startingDateArg As Date, endingDateArg As Date) As Double
     
     'Declare variables that will hold how many days fall under each month
    Dim intCount As Integer
    Dim JanCount As Integer: Dim FebCount As Integer: Dim MarCount As Integer: Dim AprCount As Integer: Dim MayCount As Integer: Dim JunCount As Integer: Dim JulCount As Integer
    Dim AugCount As Integer: Dim SeptCount As Integer: Dim OctCount As Integer: Dim NovCount As Integer: Dim DecCount As Integer
    
    Dim inDateVar As Date
    Dim outDateVar As Date
    
    Dim numDays As Integer
    Dim getMonthVar As Integer
    
    'initialize variables to zero
    intCount = 0
    JanCount = 0: FebCount = 0: MarCount = 0: AprCount = 0: MayCount = 0: JunCount = 0: JulCount = 0: AugCount = 0: SeptCount = 0: OctCount = 0: NovCount = 0: DecCount = 0

    inDateVar = startingDateArg
    outDateVar = endingDateArg
    numDays = 0
    getMonthVar = 0
        
    'Use dateDiff function to find out the difference in days inclusive, between the two dates entered by the user
    numDays = DateDiff("d", inDateVar, outDateVar) + 1
    
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
    
'    Spring = DLookup("[SEASON_Factor]", "SeasonFactor", "[SEASON_ID] = 1")
'    Summer = DLookup("[SEASON_Factor]", "SeasonFactor", "[SEASON_ID] = 2")
'    Fall = DLookup("[SEASON_Factor]", "SeasonFactor", "[SEASON_ID] = 3")
'    Winter = DLookup("[SEASON_Factor]", "SeasonFactor", "[SEASON_ID] = 4")
    
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
    'MsgBox ("The seasonal Factor is: " & seasonalFactorTotal)
    
End Function


'button to clear all controls in the form
Private Sub buttonClearAll_Click()
    ClearAll Me
End Sub

'this function calculates the shares needed for one cattle based on what the user has entered for that row
Function calcSharesFor1Cattle(rowX As String, cattleFactor As Control, avgWeight As Control, sDate As Control, eDate As Control, dateDifference As Control, pairGraze As Double, monthsGraze As Double) As Double
        
        If Not (IsNull(sDate.Value) Or IsNull(eDate.Value) Or IsNull(avgWeight.Value) Or IsNull(cattleFactor.Value)) Then
        
        
            Dim dateInRowVar As Date
            Dim dateOutRowVar As Date
            Dim avgWeightRowVar As Double
            Dim cattleQtyRowVar As Double
            Dim cattleFactorRowVar As Double
        
            Dim seasonFactorRowVar As Double
            Dim metabolicAUMrowVar As Double
            Dim mSharesRowVar As Double
        
            dateInRowVar = sDate.Value
            dateOutRowVar = eDate.Value
            avgWeightRowVar = avgWeight.Value
            cattleQtyRowVar = 1
            cattleFactorRowVar = cattleFactor.Value
    
            If checkDates(dateInRowVar, dateOutRowVar) = 1 Then
        
                seasonFactorRowVar = seasonalFactorTotal(dateInRowVar, dateOutRowVar)
                metabolicAUMrowVar = metabolicFunctAUM(avgWeightRowVar, cattleQtyRowVar, seasonFactorRowVar)
                
                'shares used  = metabolic AUMs/ (4.2 pairs * grazing for 5 months)
                mSharesRowVar = metabolicAUMrowVar / (pairGraze * monthsGraze)
                
                dateDifference.Value = DateDiff("d", dateInRowVar, dateOutRowVar) + 1
        
                calcSharesFor1Cattle = mSharesRowVar
                
            Else:
                calcSharesFor1Cattle = 0
                Call MsgBox("The " & rowX & " end date is less than or equal to the start date. Please enter correct dates", , "CALMS")
                Exit Function
            End If
        
        End If
        
End Function

