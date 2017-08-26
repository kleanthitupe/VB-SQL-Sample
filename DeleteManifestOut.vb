Option Compare Database


Private Sub buttonDelete_Click()
    If Not (IsNull(Me.comboShareholder.Value) Or IsNull(Me.comboManifestNumber.Value) Or IsNull(Me.txtDateInLast.Value)) Then

        prompt = "Are you sure you want to delete this manifest?"
        
        If (MsgBox(prompt, 1, "CALMS") = 1) Then
        
            Dim sharesGiveBack As Double
            sharesGiveBack = Me.txtTotalSharesLast.Value
            
            Me.txtShareholerDeleted.Value = Me.comboShareholder.Column(1)
            Me.txtManNoDeleted.Value = Me.comboManifestNumber.Column(1)
            Me.txtDateInDeleted.Value = Me.txtDateInLast.Value
            Me.txtFieldDeleted.Value = Me.txtFieldLast.Value
            Me.txtTotalAUMsDeleted.Value = Me.txtTotalAumLast.Value
            Me.txtTotalSharesDeleted.Value = Me.txtTotalSharesLast.Value
            Me.txtCommentsDeleted.Value = Me.txtCommentsLast.Value
            
            Me.txtQtyBHeifDeleted.Value = Me.txtQtyBredHeifLast.Value
                
            Me.txtQtyBullsDeleted.Value = Me.txtQtyBullsLast.Value
             
            Me.txtQtyCalvesDeleted.Value = Me.txtQtyCalvesLast.Value
                
            Me.txtQtyCowsDeleted.Value = Me.txtQtyCowsLast.Value
                
            Me.txtQtyYearlingsDeleted.Value = Me.txtQtyYearlingsLast.Value
            
            Call locationsDateOutSetNull
            Call deleteManifest
            
            
            Dim db As DAO.Database
            Dim rsShareholder As DAO.Recordset
            Dim sqlSharesAvailable As String
            sqlSharesAvailable = "SELECT [SH_Shares_Used_MOut], [SH_Total_Shares], [SH_Num_Avail_Shares] FROM Shareholder WHERE SH_ID = (" & Me.comboShareholder.Column(0) & ");"
            
        
            Set db = CurrentDb
            Set rsShareholder = db.OpenRecordset(sqlSharesAvailable)
        
            rsShareholder.Edit
            rsShareholder!SH_Shares_Used_MOut = rsShareholder!SH_Shares_Used_MOut - sharesGiveBack
            rsShareholder!SH_Num_Avail_Shares = rsShareholder!SH_Num_Avail_Shares + sharesGiveBack
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

Sub locationsDateOutSetNull()

Dim sql As String
sql = "UPDATE Manifest INNER JOIN (Bovine INNER JOIN Location ON Bovine.BOV_ID = Location.BOV_ID) ON Manifest.MANIFEST_ID = Bovine.Manifest_ID_Out SET Location.[LOC_Date_OUT] = Null, Location.LOC_AUMs_USED = 0, Location.LOC_Manifest_Out = False WHERE MANIFEST_ID = " & Me.comboManifestNumber.Column(0) & " AND Location.LOC_Date_OUT = #" & Me.txtDateInLast.Value & "#;"

DoCmd.SetWarnings False
DoCmd.RunSQL (sql)
DoCmd.SetWarnings True

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
    Dim fieldName As String

    ManifestID = Me.comboManifestNumber.Column(0)

    sql = "SELECT * FROM Manifest WHERE MANIFEST_ID = " & ManifestID & ";"

    Dim rsManifest As DAO.Recordset
    Set rsManifest = CurrentDb.OpenRecordset(sql)
    Me.txtDateInLast.Value = rsManifest("MANIFEST_Date").Value
    Me.txtCommentsLast.Value = rsManifest("MANIFEST_Information").Value
    Me.txtTotalAumLast.Value = rsManifest("MANIFEST_AUMs_Man_Out").Value
    Me.txtTotalSharesLast.Value = rsManifest("MANIFEST_Shares_Man_Out").Value
          
    Dim sql1 As String
    sql1 = "SELECT * FROM ManifestOutEdit WHERE MANIFEST_ID = " & ManifestID & " AND Location.LOC_Date_OUT = #" & Me.txtDateInLast.Value & "#;"
    
    Dim rsManifestEdit As DAO.Recordset
    Set rsManifestEdit = CurrentDb.OpenRecordset(sql1)

    If Not rsManifestEdit.EOF Then
            Me.txtFieldLast.Value = rsManifestEdit("F_Name").Value
    End If
    
    Dim bredHeifersID As Integer
    Dim bullsID As Integer
    Dim cowsID As Integer
    Dim calvesID As Integer
    Dim yearlingsID As Integer

    bredHeifersID = DLookup("[CAT_ID]", "Category", "[CAT_Category] = 'Bred Heifers'")
    bullsID = DLookup("[CAT_ID]", "Category", "[CAT_Category] = 'Bulls'")
    cowsID = DLookup("[CAT_ID]", "Category", "[CAT_Category] = 'Cows'")
    calvesID = DLookup("[CAT_ID]", "Category", "[CAT_Category] = 'Calves'")
    yearlingsID = DLookup("[CAT_ID]", "Category", "[CAT_Category] = 'Yearlings'")

    Me.txtQtyBredHeifLast.Value = DCount("LOC_ID", "ManifestOutEdit", "CAT_ID = " & bredHeifersID & " And [MANIFEST_ID] = " & ManifestID & "  And LOC_Date_OUT = #" & Me.txtDateInLast.Value & "# ")

    Me.txtQtyBullsLast.Value = DCount("LOC_ID", "ManifestOutEdit", "CAT_ID = " & bullsID & " And MANIFEST_ID = " & ManifestID & "  And LOC_Date_OUT = #" & Me.txtDateInLast.Value & "# ")
    Me.txtQtyCowsLast.Value = DCount("BOV_ID", "ManifestOutEdit", "CAT_ID = " & cowsID & " And MANIFEST_ID = " & ManifestID & "  And LOC_Date_OUT = #" & Me.txtDateInLast.Value & "# ")
    Me.txtQtyCalvesLast.Value = DCount("BOV_ID", "ManifestOutEdit", "CAT_ID = " & calvesID & " And MANIFEST_ID = " & ManifestID & "  And LOC_Date_OUT = #" & Me.txtDateInLast.Value & "# ")
    Me.txtQtyYearlingsLast.Value = DCount("BOV_ID", "ManifestOutEdit", "CAT_ID = " & yearlingsID & " And MANIFEST_ID = " & ManifestID & "  And LOC_Date_OUT = #" & Me.txtDateInLast.Value & "# ")
    Dim aums As Double
    'aums = DSum("LOC_AUMs_USED", "ManifestOutEdit", "MANIFEST_ID = " & ManifestID & " ")
    'MsgBox (aums)
    

    rsManifestEdit.Close
    Set rsManifestEdit = Nothing
    rsManifest.Close
    Set rsManifest = Nothing

End Sub

Private Sub comboShareholder_Change()
    Call ClearGroupWithTag(Me.Form, "deleteManifestData")
    Dim unchecked As String
    unchecked = "No"
    
    Dim sql As String
    sql = "SELECT DISTINCT Manifest.MANIFEST_ID, Manifest.MANIFEST_Number, [Manifest.MANIFEST_In/Out], Shareholder.SH_ID, Shareholder.SH_Contact_Name_First, Shareholder.Sh_Contact_Name_Last FROM Shareholder INNER JOIN Manifest ON Shareholder.SH_ID = Manifest.SH_ID WHERE Shareholder.SH_ID = " & Me.comboShareholder.Column(0) & " AND [Manifest.MANIFEST_In/Out] = Yes;"
    
    Me.comboManifestNumber.RowSource = sql
    Me.comboManifestNumber.Requery
    
End Sub

