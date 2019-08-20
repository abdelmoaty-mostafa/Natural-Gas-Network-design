Private Sub UIButtonControl1_Click()

Dim sure As Integer
 sure = MsgBox("LOADING A NEW DATA : ARE YOU SURE ?", vbYesNo, "LOAD NEW Data")

If sure = 6 Then

'------------------------------------------------------------------------------
'GET THE FIRST CUSTOMER DATA
'create new gx dialog
Dim pGXDialog As IGxDialog
Set pGXDialog = New GxDialog
' set properties
pGXDialog.ButtonCaption = "LOAD DATA"
pGXDialog.StartingLocation = "E:\Gis"
pGXDialog.Title = "Custom Add Data"
pGXDialog.AllowMultiSelect = False
'create filter
Dim pLFilter As IGxObjectFilter
Set pLFilter = New GxFilterPersonalGeodatabases
Set pGXDialog.ObjectFilter = pLFilter
Dim pTablefiles As IEnumGxObject
Dim pick_file As Boolean
If Not pGXDialog.DoModalOpen(0, pTablefiles) Then
Exit Sub
End If
Dim path As String
path = pTablefiles.Next.FullName
'-----------------------------------------
Dim pFact As IWorkspaceFactory
Set pFact = New AccessWorkspaceFactory

Dim pWorkspace As IWorkspace
Set pWorkspace = pFact.OpenFromFile(path, 0)
 
 Dim pFWorkspace As IFeatureWorkspace
 Set pFWorkspace = pWorkspace
 
 Dim pTable As ITable
 Set pTable = pFWorkspace.OpenTable("NEW_CUSTOMERS_DATA")

'--------------------------------------------------------------------------------
' get block no from sector layer

Dim pMxDoc As IMxDocument
Set pMxDoc = ThisDocument
Dim pMap As IMap
Set pMap = pMxDoc.FocusMap
Dim palllayer As IEnumLayer
Set palllayer = pMap.Layers
Dim pLayer As ILayer
Set pLayer = palllayer.Next
 
'--------------------------------------------------------

Do Until pLayer Is Nothing
If pLayer.name = "SECTORS" Then
  Dim pFeatureLayer As IFeatureLayer
  Set pFeatureLayer = pLayer
  Dim pFeatureClass As IFeatureClass
  Set pFeatureClass = pFeatureLayer.FeatureClass
    Dim pQfilt As IQueryFilter
      Set pQfilt = New QueryFilter
      Dim SEC_no As Integer
        SEC_no = InputBox("Please Enter the sector no : ")
      clause_SEC = "SECTOR_NUM =" & SEC_no
      pQfilt.whereClause = clause_SEC
      
      Dim pFeatureCursor As IFeatureCursor
      Set pFeatureCursor = pFeatureClass.Search(pQfilt, False)  'Do the search
      Dim pFeature As IFeature
      Set pFeature = pFeatureCursor.NextFeature  'Get the first feature
  Dim blk_no As Integer
  blk_no = pFeature.Value(9)
  Dim sec_num As String
  sec_num = pFeature.Value(8)
'MsgBox blk_no
End If
Set pLayer = palllayer.Next
Loop
 
'-------------------------------------------------------------------
'calculate the building loads according to building-no in the block
  
  Dim i As Integer
For i = 0 To (blk_no - 1)
  
  '---------------------------------------------
  ' get building_no in block i from blocks
  palllayer.Reset
  Set pLayer = palllayer.Next

Do Until pLayer Is Nothing
If pLayer.name = "BLOCKS" Then
   Set pFeatureLayer = pLayer
   Set pFeatureClass = pFeatureLayer.FeatureClass
   Set pQfilt = New QueryFilter
   Dim clause_bld As String
   clause_bld = i + 1
   pQfilt.whereClause = "BLOCK_NUM =" & clause_bld
   Set pFeatureCursor = pFeatureClass.Search(pQfilt, False)
   Set pFeature = pFeatureCursor.NextFeature
   Dim bld_no As Integer
  bld_no = pFeature.Value(10)
'MsgBox bld_no
End If
Set pLayer = palllayer.Next
Loop
 
 '--------------------------------------------------------------
 ' get all rows in block i from customer table
 
  Dim blk_clause As String
  Dim blk_num_clause As Integer
  blk_num_clause = i + 1
  blk_clause = "block=" & blk_num_clause
  Dim pQueryFilter As IQueryFilter
  Set pQueryFilter = New QueryFilter
  pQueryFilter.whereClause = blk_clause
  Dim pCursor As ICursor
  Set pCursor = pTable.Search(pQueryFilter, False)
  
  Dim prow As IRow
  Dim status As String
  Dim activity As String
  Dim comm_load As Double
  
  Dim AC As Integer
  Dim NA As Integer
  Dim UC As Integer
  Dim RG As Integer
  Dim VAC As Integer
  Dim NG As Integer
  Dim comm As Integer
  Dim NAPP As Integer
  Dim load As Double
  'Dim load_AC As Double
  'Dim load_Napp As Double
  Dim j As Integer
For j = 1 To bld_no
  AC = 0
  NA = 0
  UC = 0
  RG = 0
  VAC = 0
  NG = 0
  comm = 0
  NAPP = 0
  load = 0
 
  
  Set prow = pCursor.NextRow
  Do Until prow Is Nothing
     If prow.Value(pTable.FindField("building")) = j Then
  
  '----------------------------------------------
  'get status values for row
  status = prow.Value(pTable.FindField("status"))
  activity = prow.Value(pTable.FindField("activity"))
  comm_load = prow.Value(pTable.FindField("comm_load"))
 'MsgBox status
  '-------------------------------------------------------------------------
  'get result load
  If status = "Ac" Then
  AC = AC + 1
    If activity = "com" Then
  load = load + comm_load
  comm = comm + 1
  AC = AC - 1
  End If
  ElseIf status = "Na" Then
  NA = NA + 1
    If activity = "com" Then
  comm = comm + 1
  'NA = NA - 1
  End If
  ElseIf status = "Uc" Then
  UC = UC + 1
    If activity = "com" Then
  comm = comm + 1
  'UC = UC - 1
  End If
  ElseIf status = "Rg" Then
  RG = RG + 1
    If activity = "com" Then
  comm = comm + 1
  'RG = RG - 1
  End If
  ElseIf status = "Vac" Then
  VAC = VAC + 1
    If activity = "com" Then
  comm = comm + 1
  'VAC = VAC - 1
  End If
  ElseIf status = "Ng" Then
  NG = NG + 1
    If activity = "com" Then
  comm = comm + 1
   'NG = NG - 1
  End If
  ElseIf status = "Napp" Then
  NAPP = NAPP + 1
  
    If activity = "com" Then
  load = load + comm_load
  comm = comm + 1
  'NAPP = NAPP - 1
  End If
  End If
     End If
  Set prow = pCursor.NextRow
  
  Loop
   
 'MsgBox AC & "-" & NA & "-" & UC & "-" & RG & "-" & VAC & "-" & NG
 '----------------------------------------------------------
 'Edit loads in landuse
 palllayer.Reset
 Dim pLayer_LU As ILayer
Set pLayer_LU = palllayer.Next
 Dim sec_num_lu As Integer
 sec_num_lu = Val(sec_num)
 
 
Do Until pLayer_LU Is Nothing
  If pLayer_LU.name = "LANDUSE" Then
     Dim pfeaturelayer_LU As IFeatureLayer
     Set pfeaturelayer_LU = pLayer_LU
     Dim pFeatureClass_LU As IFeatureClass
     Set pFeatureClass_LU = pfeaturelayer_LU.FeatureClass
     Dim pQFilt_LU As IQueryFilter
     Set pQFilt_LU = New QueryFilter
     Dim clause_LU As String
     clause_LU = "SECTOR_NUM =" & sec_num
     pQFilt_LU.whereClause = clause_LU
     Dim pFeatureCursor_LU As IFeatureCursor
     Set pFeatureCursor_LU = pFeatureClass_LU.Update(pQFilt_LU, False)
     Dim pFeature_LU As IFeature
     Set pFeature_LU = pFeatureCursor_LU.NextFeature
         Do Until pFeature_LU Is Nothing
            'MsgBox i & j
             
            If ((pFeature_LU.Value(9) = (i + 1)) And (pFeature_LU.Value(10) = j)) Then
               
               pFeature_LU.Value(11) = AC
               pFeature_LU.Value(12) = NA
               pFeature_LU.Value(13) = UC
               pFeature_LU.Value(14) = RG
               pFeature_LU.Value(15) = VAC
               pFeature_LU.Value(16) = NG
               pFeature_LU.Value(17) = comm
               pFeature_LU.Value(18) = NAPP
               pFeature_LU.Value(19) = load
               pFeatureCursor_LU.UpdateFeature pFeature_LU
            '--------------------------------------------------------------
            ' Edit load in service layer
             palllayer.Reset
             Dim pLayer_SE As ILayer
             Set pLayer_SE = palllayer.Next
             Dim Land_id_for_SE As String
             Land_id_for_SE = pFeature_LU.Value(0)
             Do Until pLayer_SE Is Nothing
                    If pLayer_SE.name = "SERVICE" Then
                        Dim pfeaturelayer_SE As IFeatureLayer
                        Set pfeaturelayer_SE = pLayer_SE
                        Dim pFeatureClass_SE As IFeatureClass
                        Set pFeatureClass_SE = pfeaturelayer_SE.FeatureClass
                        Dim pQFilt_SE As IQueryFilter
                        Set pQFilt_SE = New QueryFilter
                        Dim clause_SE As String
                        clause_SE = "LANDUSE_ID =" & Land_id_for_SE
                        pQFilt_SE.whereClause = clause_SE
                        Dim pFeatureCursor_SE As IFeatureCursor
                        Set pFeatureCursor_SE = pFeatureClass_SE.Update(pQFilt_SE, False)
                        Dim pFeature_SE As IFeature
                        Set pFeature_SE = pFeatureCursor_SE.NextFeature
                        Do Until pFeature_SE Is Nothing
                                pFeature_SE.Value(10) = AC
                                pFeature_SE.Value(11) = NA
                                pFeature_SE.Value(12) = UC
                                pFeature_SE.Value(13) = RG
                                pFeature_SE.Value(14) = NG
                                pFeature_SE.Value(15) = VAC
                                pFeature_SE.Value(16) = comm
                                pFeature_SE.Value(17) = NAPP
                                pFeature_SE.Value(18) = load
                                pFeatureCursor_SE.UpdateFeature pFeature_SE
                                Set pFeature_SE = pFeatureCursor_SE.NextFeature
                        Loop
                   
                   End If
            Set pLayer_SE = palllayer.Next
            Loop
            '-------------------------------------------------------------
            'Edit load in serv layer
            palllayer.Reset
            Dim pLayer_serv As ILayer
            Set pLayer_serv = palllayer.Next
                'Dim SEC_STR As String
                'SEC_STR = sec_num
                Do Until pLayer_serv Is Nothing
                    If pLayer_serv.name = "SERV" Then
                        Dim pfeaturelayer_serv As IFeatureLayer
                        Set pfeaturelayer_serv = pLayer_serv
                        Dim pFeatureClass_serv As IFeatureClass
                        Set pFeatureClass_serv = pfeaturelayer_serv.FeatureClass
                        Dim pQFilt_serv As IQueryFilter
                        Set pQFilt_serv = New QueryFilter
                        Dim clause_serv As String
                        clause_serv = "SECTOR_NO =" & "'" & sec_num & "'"
                        pQFilt_serv.whereClause = clause_serv
                        Dim pFeatureCursor_serv As IFeatureCursor
                        Set pFeatureCursor_serv = pFeatureClass_serv.Update(pQFilt_serv, False)
                        Dim pFeature_serv As IFeature
                        Set pFeature_serv = pFeatureCursor_serv.NextFeature
                            Do Until pFeature_serv Is Nothing
                                If ((pFeature_serv.Value(8) = (i + 1)) And (pFeature_serv.Value(9) = j)) Then
               
                                    pFeature_serv.Value(10) = AC
                                    pFeature_serv.Value(11) = NA
                                    pFeature_serv.Value(12) = UC
                                    pFeature_serv.Value(13) = RG
                                    pFeature_serv.Value(14) = NG
                                    pFeature_serv.Value(15) = VAC
                                    pFeature_serv.Value(16) = comm
                                    pFeature_serv.Value(17) = NAPP
                                    pFeature_serv.Value(18) = load
                                    pFeatureCursor_serv.UpdateFeature pFeature_serv
                                End If
                                Set pFeature_serv = pFeatureCursor_serv.NextFeature
                            Loop
                    End If
                Set pLayer_serv = palllayer.Next
                Loop
            '--------------------------------------------------------------
            'Edit load in riser layer
            palllayer.Reset
            Dim pLayer_R As ILayer
            Set pLayer_R = palllayer.Next
                Do Until pLayer_R Is Nothing
                    If pLayer_R.name = "RISER" Then
                        Dim pfeaturelayer_R As IFeatureLayer
                        Set pfeaturelayer_R = pLayer_R
                        Dim pFeatureClass_R As IFeatureClass
                        Set pFeatureClass_R = pfeaturelayer_R.FeatureClass
                        Dim pQFilt_R As IQueryFilter
                        Set pQFilt_R = New QueryFilter
                        Dim clause_R As String
                        clause_R = "SECTOR_NO =" & "'" & sec_num & "'"
                        pQFilt_R.whereClause = clause_R
                        Dim pFeatureCursor_R As IFeatureCursor
                        Set pFeatureCursor_R = pFeatureClass_R.Update(pQFilt_R, False)
                        Dim pFeature_R As IFeature
                        Set pFeature_R = pFeatureCursor_R.NextFeature
                            Do Until pFeature_R Is Nothing
                                If ((pFeature_R.Value(7) = (i + 1)) And (pFeature_R.Value(8) = j)) Then
                                    pFeature_R.Value(10) = AC
                                    pFeatureCursor_R.UpdateFeature pFeature_R
                                End If
                                Set pFeature_R = pFeatureCursor_R.NextFeature
                            Loop
                    End If
                Set pLayer_R = palllayer.Next
                Loop
            
            '-------------------------------------------------------------
            End If
            Set pFeature_LU = pFeatureCursor_LU.NextFeature
         Loop
   End If
   Set pLayer_LU = palllayer.Next
Loop
'----------------------------------------------------------------------

 Set pCursor = pTable.Search(pQueryFilter, False)
Next j
 
Next i
MsgBox ("Load transfer: complete successfully.")
Else
Dim terminated As Integer
terminated = MsgBox("Load transfer terminated.", vbExclamation)
End If
End Sub
