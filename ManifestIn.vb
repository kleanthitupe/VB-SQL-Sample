Option Compare Database

'This is the code for the button that adds a new field after the user has added appropriate info
Private Sub buttonAddNewField_Click()
    If Not (IsNull(Me.txtFieldName.Value) Or IsNull(Me.txtFieldRating.Value) Or IsNull(Me.txtFieldSize.Value)) Then

        Dim db As Database
        Dim rs As DAO.Recordset
        'Open the recordset, which is the table Field in this case
        Set db = CurrentDb
        Set rs = db.OpenRecordset("Field")
        
        'Declare all local variables
        Dim fieldNameVar As String
        Dim fieldSizeVar As Double
        Dim deededAcresVar As Double
        Dim forestryAcresVar As Double
        Dim leaseAcresVar As Double
        Dim stockingRateVar As Double
        Dim fieldRatingVar As Double
        
        
        'Initialize all local variables taking the value from the appropriate text field
        fieldNameVar = Me.txtFieldName.Value
        fieldSizeVar = Me.txtFieldSize.Value
        If Not IsNull(Me.txtDeededAcres.Value) Then
            deededAcresVar = Me.txtDeededAcres.Value
        End If
        If Not IsNull(Me.txtForestryAcres.Value) Then
            forestryAcresVar = Me.txtForestryAcres.Value
        End If
        If Not IsNull(Me.txtLeaseAcres.Value) Then
            leaseAcresVar = Me.txtLeaseAcres.Value
        End If
        
        
        'Call calculateRating
        
        'Add new records into the table Field, with the values that were taken from the textboxes in the Add_Field form
        rs.AddNew
        rs("F_Name").Value = fieldNameVar
        rs("F_Size").Value = fieldSizeVar
        rs("F_DeededAcres").Value = deededAcresVar
        rs("F_ForestryAcres").Value = forestryAcresVar
        rs("F_LeaseAcres").Value = leaseAcresVar
        'rs("F_StockingRate").Value = stockingRateVar
        rs("F_Rating").Value = Me.txtFieldRating.Value
        rs.Update
        
        Call MsgBox("Field: " & fieldNameVar & " was Added Successfully!", , "CALMS")
        'set all the textboxes to Null
        Me.txtFieldName.Value = Null
        Me.txtFieldSize.Value = Null
        Me.txtDeededAcres.Value = Null
        Me.txtForestryAcres.Value = Null
        Me.txtLeaseAcres.Value = Null
        Me.txtFieldRating.Value = Null
        
        rs.Close
        db.Close
        Set rs = Nothing
        Set db = Nothing
    Else
        Call MsgBox("Please enter all the necessary input to add a field.", , "CALMS")
    End If
        
End Sub

'the code for the button that makes visible only the textboxes and buttons that are necessary for adding a Field
Private Sub buttonMenuToAddField_Click()
    Dim lngGreen As Long, lngBlue As Long
    lngGreen = RGB(84, 130, 53)
    lngBlue = RGB(46, 117, 182)
    
    Me.buttonMenuToAddField.BackColor = lngBlue
    Me.buttonEditField.BackColor = lngGreen
    
    
    Me.radioActiveField.Visible = False
        
    Me.txtFieldName.Visible = True
    Me.comboEditField.Visible = False
    Me.txtFieldName.Value = Null
    Me.txtFieldSize.Value = Null
    Me.txtDeededAcres.Value = Null
    Me.txtForestryAcres.Value = Null
    Me.txtLeaseAcres.Value = Null
    Me.txtFieldRating.Value = Null
    Me.buttonUpdateRating.Visible = True
    Me.buttonAddNewField.Visible = True
    Me.buttonUpdateRecord.Visible = False
    
End Sub

'the code for the button that makes visible only the textboxes and buttons that are necessary for editing a Field
Private Sub buttonEditField_Click()
    Dim lngGreen As Long, lngBlue As Long
    lngGreen = RGB(84, 130, 53)
    lngBlue = RGB(46, 117, 182)
    
    Me.buttonMenuToAddField.BackColor = lngGreen
    Me.buttonEditField.BackColor = lngBlue
    
    Me.radioActiveField.Visible = True
    Me.txtFieldName.Value = Null
    Me.txtFieldSize.Value = Null
    Me.txtDeededAcres.Value = Null
    Me.txtForestryAcres.Value = Null
    Me.txtLeaseAcres.Value = Null
    Me.txtFieldRating.Value = Null
    Me.comboEditField.Visible = True
    Me.buttonUpdateRating.Visible = True
    Me.buttonAddNewField.Visible = False
    Me.buttonUpdateRecord.Visible = True

End Sub

'This is the code for the button that pulls information from the Field table and enables the user to edit info
Private Sub buttonUpdateRecord_Click()
    If Not (IsNull(Me.comboEditField.Column(0)) Or IsNull(Me.txtFieldName.Value) Or IsNull(Me.txtFieldRating.Value)) Then
        Dim db As DAO.Database
        Set db = CurrentDb
        Dim sql As String
        Dim prompt As String
        
        Dim fieldIdVar As Integer
        Dim fieldNameVar As String
        Dim fieldSizeVar As Double
        Dim deededAcresVar As Double
        Dim forestryAcresVar As Double
        Dim leaseAcresVar As Double
        Dim stockingRateVar As Double
        Dim fieldRatingVar As Double
        
        fieldIdVar = Me.comboEditField.Column(0)
        fieldNameVar = Me.txtFieldName.Value
        fieldSizeVar = Me.txtFieldSize.Value
        deededAcresVar = Me.txtDeededAcres.Value
        forestryAcresVar = Me.txtForestryAcres.Value
        leaseAcresVar = Me.txtLeaseAcres.Value
        
        'Call calculateRating
        fieldRatingVar = Me.txtFieldRating.Value
        'fieldRatingVar = fieldSizeVar * 12 / stockingRateVar
        
        prompt = "Are you sure you want to go ahead with the update?"
        
        If (MsgBox(prompt, 1, "CALMS") = 1) Then
            
            sql = "UPDATE Field " _
                & "SET F_Size = " & fieldSizeVar & ", " _
                & "F_Name = '" & fieldNameVar & "', " _
                & "F_DeededAcres = " & deededAcresVar & ", " _
                & "F_ForestryAcres = " & forestryAcresVar & ", " _
                & "F_LeaseAcres = " & leaseAcresVar & ", " _
                & "F_Active = " & Me.radioActiveField.Value & ", " _
                & "F_Rating = " & Me.txtFieldRating.Value & " " _
                & "WHERE F_ID = " & fieldIdVar & " ;"
                
            'Me.txtFieldRating.Value = fieldRatingVar
            DoCmd.SetWarnings False
            DoCmd.RunSQL sql
            Call MsgBox("Field: " & fieldNameVar & " was successfully updated!", , "CALMS")
            DoCmd.SetWarnings True
        End If
        
        db.Close
        Set db = Nothing
    Else
        Call MsgBox("Please select all the necessary input.", , "CALMS")
    End If
End Sub

Private Sub comboEditField_Change()
    'open database and recordset
    Dim db As DAO.Database
    Dim recordSt As Recordset
    Dim sql As String
    Set db = CurrentDb
    
    'makeing sql statements and setting the reccordset
    sql = "SELECT * FROM Field WHERE [F_ID] = (" & Me.comboEditField.Column(0) & ");"
    Set recordSt = db.OpenRecordset(sql)
    
    'filling the textboxs with the data from the Field table
    Me.txtFieldName.Value = recordSt!F_Name
    Me.txtFieldSize.Value = recordSt!F_Size
    Me.txtDeededAcres.Value = recordSt!F_DeededAcres
    Me.txtForestryAcres.Value = recordSt!F_ForestryAcres
    Me.txtLeaseAcres.Value = recordSt!F_LeaseAcres
    Me.txtFieldRating.Value = recordSt!F_Rating
    Me.radioActiveField.Value = recordSt!F_Active
    
    
    recordSt.Close
    db.Close
    Set recordSt = Nothing
    Set db = Nothing
End Sub

Private Sub txtDeededAcres_KeyPress(KeyAscii As Integer)
    LimitNumeric KeyAscii
End Sub

Private Sub txtFieldName_KeyPress(KeyAscii As Integer)
    LimitAlphanumeric KeyAscii
End Sub


Private Sub txtFieldRating_KeyPress(KeyAscii As Integer)
    LimitNumeric KeyAscii
End Sub

Private Sub txtFieldSize_AfterUpdate()
    If Me.txtFieldSize.Value = Null Or Me.txtFieldSize.Value <= 0 Then
        Call MsgBox("Need to enter appropriate format", , "CALMS")
    End If
End Sub

Private Sub buttonUpdateRating_Click()
    Call calculateRating
End Sub

Sub calculateRating()
    Dim deededSize As Double
    Dim leasedSize As Double
    Dim forestrySize As Double
    'initialize variables to zero
    deededSize = 0: leasedSize = 0: forestrySize = 0
    
    If Not IsNull(Me.txtDeededAcres.Value) Then
        deededSize = Me.txtDeededAcres.Value
    End If
    If Not IsNull(Me.txtLeaseAcres.Value) Then
        leasedSize = Me.txtLeaseAcres.Value
    End If
    If Not IsNull(Me.txtForestryAcres.Value) Then
        forestrySize = Me.txtForestryAcres.Value
    End If
    
    'these are constants that are set by Waldron
    Dim deedLandAUM As Double
    Dim leaseLandAUM As Double
    Dim forestryLandAUM As Double
    
    deedLandAUM = 2         'acres/AUM
    leaseLandAUM = 2.77     'acres/AUM
    forestryLandAUM = 4.48  'acres/AUM
    
    
    Dim fieldRatingVar As Double
    fieldRatingVar = deededSize / deedLandAUM + leasedSize / leaseLandAUM + forestrySize / forestryLandAUM
    
    Me.txtFieldRating.Value = Round(fieldRatingVar, 2)


End Sub



Private Sub txtFieldSize_KeyPress(KeyAscii As Integer)
    LimitNumeric KeyAscii
End Sub

Private Sub txtForestryAcres_KeyPress(KeyAscii As Integer)
    LimitNumeric KeyAscii
End Sub

Private Sub txtLeaseAcres_KeyPress(KeyAscii As Integer)
    LimitNumeric KeyAscii
End Sub
