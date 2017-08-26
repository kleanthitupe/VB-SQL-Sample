Option Compare Database


Private Sub buttonDelete_Click()
    If Not (IsNull(Me.comboShareholder.Value) Or IsNull(Me.comboManifestNumber.Value) Or IsNull(Me.txtDateInLast.Value) Or Me.txtDateOutLast.Value) Then

        prompt = "Are you sure you want to delete this manifest and its contents?"
        
        If (MsgBox(prompt, 1, "CALMS") = 1) Then
        
            Dim sharesGiveBack As Double
            sharesGiveBack = Me.txtTotalSharesLast.Value
            
            Me.txtShareholerDeleted.Value = Me.comboShareholder.Column(1)
            Me.txtManNoDeleted.Value = Me.comboManifestNumber.Column(1)
            Me.txtDateInDeleted.Value = Me.txtDateInLast.Value
            Me.txtDateOutDeleted.Value = Me.txtDateOutLast.Value
            Me.txtNumDaysDeleted.Value = Me.txtNumDaysLast.Value
            Me.txtFieldDeleted.Value = Me.txtFieldLast.Value
            Me.txtTotalAUMsDeleted.Value = Me.txtTotalAumLast.Value
            Me.txtTotalSharesDeleted.Value = Me.txtTotalSharesLast.Value
            Me.txtCommentsDeleted.Value = Me.txtCommentsLast.Value
            
            Me.txtQtyBHeifDeleted.Value = Me.txtQtyBredHeifLast.Value
            Me.txtBHeifDeleted.Value = Me.txtBredHeifWeightLast.Value
                
            Me.txtQtyBullsDeleted.Value = Me.txtQtyBullsLast.Value
            Me.txtBullsDeleted.Value = Me.txtBullsWeightLast.Value
                
            Me.txtQtyCalvesDeleted.Value = Me.txtQtyCalvesLast.Value
            Me.txtCalvesDeleted.Value = Me.txtCalvesWeightLast.Value
                
            Me.txtQtyCowsDeleted.Value = Me.txtQtyCowsLast.Value
            Me.txtCowsDeleted.Value = Me.txtCowsWeightLast.Value
                
            Me.txtQtyYearlingsDeleted.Value = Me.txtQtyYearlingsLast.Value
            Me.txtYearlingsDeleted.Value = Me.txtYearlingsWeightLast.Value
            
            Call deleteLocations
            Call deleteBovines
            Call deleteManifest
            
            
            Dim db As DAO.Database
            Dim rsShareholder As DAO.Recordset
            Dim sqlSharesAvailable As String
            sqlSharesAvailable = "SELECT [SH_Shares_Used], [SH_Total_Shares] FROM Shareholder WHERE SH_ID = (" & Me.comboShareholder.Column(0) & ");"
            
        
            Set db = CurrentDb
            Set rsShareholder = db.OpenRecordset(sqlSharesAvailable)
        
            rsShareholder.Edit
            rsShareholder!SH_Shares_Used = rsShareholder!SH_Shares_Used - sharesGiveBack
            rsShareholder.Update
        
            rsShareholder.Close
            Set rsShareholder = Nothing
            db.Close
            Set db = Nothing
            
                
            'MessageBox.Show("The manifest and its contents were deleted succesfully.", "Edit Manifest.")
            Call MsgBox("The manifest and its contents were deleted succesfully.", , "CALMS")
            Me.Refresh
            Me.Requery
            
            Call ClearGroupWithTag(Me.Form, "deleteManifestData")
            Me.comboShareholder.Value = Null
            
            
            
        
        End If
    Else
        Call MsgBox("Please choose a valid manifest to delete.", , "CALMS")
        Exit Sub
    End If

End Sub

Sub deleteLocations()

On Error GoTo ErrorHandler

Dim sql As String
Dim rsLocation As DAO.Recordset

sql = "SELECT * FROM DeleteManifestIn WHERE Bovine.Manifest_ID_In = " & Me.comboManifestNumber.Column(0) & ";"

Set rsLocation = CurrentDb.OpenRecordset(sql)

With rsLocation

    While (Not .BOF And Not .EOF)

        .MoveLast
        .MoveFirst

        If .Updatable Then

            .Delete

        End If
    Wend

    .Close
End With


ExitSub:
    Set rsLocation = Nothing
    
    Exit Sub
ErrorHandler:
    Resume ExitSub
End Sub

Sub deleteBovines()
        
On Error GoTo ErrorHandler

Dim sql As String
Dim rs As DAO.Recordset

sql = "SELECT * FROM Bovine WHERE Bovine.Manifest_ID_In = " & Me.comboManifestNumber.Column(0) & ";"

Set rs = CurrentDb.OpenRecordset(sql)

With rs

    While (Not .BOF And Not .EOF)
    
        .MoveLast
        .MoveFirst
        
        If .Updatable Then
        
            .Delete
           
        End If
    Wend
    
    .Close
End With

ExitSub:
    Set rs = Nothing
  
    Exit Sub
ErrorHandler:
    Resume ExitSub

End Sub

Sub deleteManifest()
        
On Error GoTo ErrorHandler

Dim sql As String
Dim rs As DAO.Recordset

sql = "SELECT * FROM Manifest WHERE MANIFEST_ID = " & Me.comboManifestNumber.Column(0) & ";"


Set rs = CurrentDb.OpenRecordset(sql)

With rs

    While (Not .BOF And Not .EOF)
    
        .MoveLast
        .MoveFirst
        
        If .Updatable Then
        
            .Delete
           
        End If
    Wend
    
    .Close
End With

ExitSub:
    Set rs = Nothing
    
    Exit Sub
ErrorHandler:
    Resume ExitSub

End Sub


Private Sub comboManifestNumber_Change()
    Dim ManifestID As Long
    Dim sql As String
    Dim sql1 As String
    
    ManifestID = Me.comboManifestNumber.Column(0)
    sql1 = "SELECT * FROM ManifestInEdit WHERE MANIFEST_ID = " & ManifestID & ";"
    sql = "SELECT * FROM Manifest WHERE MANIFEST_ID = " & ManifestID & ";"

    
    Dim rsManifest As DAO.Recordset
    Set rsManifest = CurrentDb.OpenRecordset(sql)
    
    Dim rsManifestEdit As DAO.Recordset
    Set rsManifestEdit = CurrentDb.OpenRecordset(sql1)
        
        Me.txtDateInLast.Value = rsManifest("MANIFEST_Date").Value
        Me.txtDateOutLast.Value = rsManifest("MANIFEST_Predicted_Date_Out").Value
        Me.txtNumDaysLast.Value = rsManifest("MANIFEST_Predicted_Date_Out").Value - rsManifest("MANIFEST_Date").Value
        Me.txtTotalAumLast.Value = rsManifest("MANIFEST_AUMs_Predicted").Value
        Me.txtTotalSharesLast.Value = rsManifest("MANIFEST_Shares_Predicted").Value
        Me.txtCommentsLast.Value = rsManifest("MANIFEST_Information").Value
        If Not rsManifestEdit.EOF Then
            Me.txtFieldLast.Value = rsManifestEdit("F_Name").Value
        End If
        
    Dim bredHeifersID As Integer
    Dim bullsID As Integer
    Dim cowsID As Integer
    Dim calvesID As Integer
    Dim yearlingsID As Integer
    
    'values taken from the lookup table: Category
    bredHeifersID = DLookup("[CAT_ID]", "Category", "[CAT_Category] = 'Bred Heifers'")
    bullsID = DLookup("[CAT_ID]", "Category", "[CAT_Category] = 'Bulls'")
    cowsID = DLookup("[CAT_ID]", "Category", "[CAT_Category] = 'Cows'")
    calvesID = DLookup("[CAT_ID]", "Category", "[CAT_Category] = 'Calves'")
    yearlingsID = DLookup("[CAT_ID]", "Category", "[CAT_Category] = 'Yearlings'")

    Me.txtQtyBredHeifLast.Value = DCount("BOV_ID", "ManifestInEdit", "CAT_ID = " & bredHeifersID & " And MANIFEST_ID = " & ManifestID & " ")
    Me.txtQtyBullsLast.Value = DCount("BOV_ID", "ManifestInEdit", "CAT_ID = " & bullsID & " And MANIFEST_ID = " & ManifestID & " ")
    Me.txtQtyCowsLast.Value = DCount("BOV_ID", "ManifestInEdit", "CAT_ID = " & cowsID & " And MANIFEST_ID = " & ManifestID & " ")
    Me.txtQtyCalvesLast.Value = DCount("BOV_ID", "ManifestInEdit", "CAT_ID = " & calvesID & " And MANIFEST_ID = " & ManifestID & " ")
    Me.txtQtyYearlingsLast.Value = DCount("BOV_ID", "ManifestInEdit", "CAT_ID = " & yearlingsID & " And MANIFEST_ID = " & ManifestID & " ")
    
    Me.txtBredHeifWeightLast.Value = DLookup("BOV_Weight", "ManifestInEdit", "CAT_ID = " & bredHeifersID & " And MANIFEST_ID = " & ManifestID & " ")
    Me.txtBullsWeightLast.Value = DLookup("BOV_Weight", "ManifestInEdit", "CAT_ID = " & bullsID & " And MANIFEST_ID = " & ManifestID & " ")
    Me.txtCowsWeightLast.Value = DLookup("BOV_Weight", "ManifestInEdit", "CAT_ID = " & cowsID & " And MANIFEST_ID = " & ManifestID & " ")
    Me.txtCalvesWeightLast.Value = DLookup("BOV_Weight", "ManifestInEdit", "CAT_ID = " & calvesID & " And MANIFEST_ID = " & ManifestID & " ")
    Me.txtYearlingsWeightLast.Value = DLookup("BOV_Weight", "ManifestInEdit", "CAT_ID = " & yearlingsID & " And MANIFEST_ID = " & ManifestID & " ")
    
    rsManifest.Close
    rsManifestEdit.Close
    
    Set rsManifestEdit = Nothing
    Set rsManifest = Nothing
End Sub

Private Sub comboShareholder_Change()
    Call ClearGroupWithTag(Me.Form, "deleteManifestData")
    Dim unchecked As String
    unchecked = "No"
    
    Dim sql As String
    sql = "SELECT DISTINCT Manifest.MANIFEST_ID, Manifest.MANIFEST_Number, [Manifest.MANIFEST_In/Out], Shareholder.SH_ID, Shareholder.SH_Contact_Name_First, Shareholder.Sh_Contact_Name_Last FROM Shareholder INNER JOIN Manifest ON Shareholder.SH_ID = Manifest.SH_ID WHERE Shareholder.SH_ID = " & Me.comboShareholder.Column(0) & " AND [Manifest.MANIFEST_In/Out] = No;"
    
    Me.comboManifestNumber.RowSource = sql
    Me.comboManifestNumber.Requery
    
End Sub


