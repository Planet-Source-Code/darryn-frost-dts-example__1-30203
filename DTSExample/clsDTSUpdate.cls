VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDTSUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents goPackage As DTS.Package
Attribute goPackage.VB_VarHelpID = -1

Private m_Provider As String
Private m_DataSource As String
Private m_Catalog As String
Private m_WorkStation As String
Private m_UserID As String
Private m_Pwd As String
Private m_UseNTSecurity As Boolean
Private m_Cancel As Boolean
Private m_Totaltasks As Integer
Private m_TasksCompleted As Integer

Public Event ErrorOccurred(ByVal pErr As Long, ByVal pSource As String, ByVal pDescription As String)
Public Event PercentDone(ByVal percent As Integer)
Public Event RowsCopied(ByVal RowsCopied As String)
Public Event Currenttask(ByVal pCurrenttask As String)
Public Event CurrentStep(ByVal pCurrentStep As String)

Public Property Let CancelTask(pCancel As Boolean)
  m_Cancel = pCancel
End Property

Private Sub ParseConnectionString(ByVal strConnect As String)
  Dim i As Long
  'Dim pCon As New ADODB.Connection
  
  'm_Provider = ""
  'm_DataSource = ""
  'm_UserID = "admin"
  'm_Pwd = ""
  'm_WorkStation = ""
  'm_Catalog = ""
  
  'pCon.ConnectionString = strConnect
  
  'm_Provider = pCon.Provider
    
  'If m_Provider = "SQLOLEDB.1" Then m_UserID = "sa"
  
  'For i = 0 To pCon.Properties.Count - 1
    
  '  Select Case pCon.Properties(i).Name
      
  '    Case "User ID"
  '      If Not Len(pCon.Properties(i).Value) = 0 Then m_UserID = pCon.Properties(i).Value
  '    Case "Data Source"
  '      m_DataSource = pCon.Properties(i).Value
  '    Case "Initial Catalog"
  '      m_Catalog = pCon.Properties(i).Value
  '    Case "WorkStation ID"
 '       m_WorkStation = pCon.Properties(i).Value
 '     Case "Persist Security Info"
 '       If pCon.Properties(i).Value = "SSP1" And m_Provider = "SQLOLEDB.1" Then
 '         m_UseNTSecurity = True
 '       Else
 '         m_UseNTSecurity = pCon.Properties(i).Value
 '       End If
        
'    End Select
  
'  Next i

 ' Set pCon = Nothing
  
End Sub

Private Sub AddConnection(ByVal lonConnectionID As Long, ByVal strConnectionName As String, _
                         Optional ByVal pFilePath As String, Optional ByVal pReusable As Boolean = False)

  Dim oConnection As DTS.Connection
  
  Set oConnection = goPackage.Connections.New(m_Provider)
        
     With oConnection
        .Name = strConnectionName
        .ID = lonConnectionID
        .Reusable = pReusable
        .ConnectImmediate = False
        .DataSource = m_DataSource
        .UserID = m_UserID
        .Password = m_Pwd
        .ConnectionTimeout = 30
        .Catalog = m_Catalog
        .UseTrustedConnection = m_UseNTSecurity
      End With
      
      goPackage.Connections.Add oConnection
      Set oConnection = Nothing
      
End Sub

Private Sub CreateStep(ByVal pStepName As String, ByVal pStepDescription As String, _
                      ByVal pTaskName As String, Optional ByVal pCloseConnection As Boolean = False, _
                      Optional ByVal pCommitOnSuccess = False, Optional ByVal pRollbackOnFailure = False)
  
  Dim oStep As DTS.Step
  
  Set oStep = goPackage.Steps.New
    oStep.Name = pStepName
    oStep.Description = pStepDescription
    oStep.ExecutionStatus = 1
    oStep.TaskName = pTaskName
    oStep.CommitSuccess = False
    oStep.RollbackFailure = False
    oStep.ScriptLanguage = "VBScript"
    oStep.AddGlobalVariables = True
    oStep.RelativePriority = 3
    oStep.CloseConnection = pCloseConnection
    oStep.ExecuteInMainThread = True
    oStep.IsPackageDSORowset = False
    oStep.JoinTransactionIfPresent = False
    oStep.DisableStep = False
  
  goPackage.Steps.Add oStep
  
  Set oStep = Nothing

End Sub

Private Sub AddExecuteSQLTask(ByVal pTaskName As String, ByVal pTaskDescription As String, _
                              ByVal pSQL As String, ByVal pConnectionID As Long)
  
  Dim oCustomTask As DTS.ExecuteSQLTask
  Dim oTask As DTS.Task
  
  Set oTask = goPackage.Tasks.New("DTSExecuteSQLTask")
        
  Set oCustomTask = oTask.CustomTask
    
    oCustomTask.Name = pTaskName
    oCustomTask.Description = pTaskDescription
    oCustomTask.SQLStatement = pSQL
    oCustomTask.ConnectionID = pConnectionID
    oCustomTask.CommandTimeout = 0
    
    goPackage.Tasks.Add oTask
    
  Set oCustomTask = Nothing
  Set oTask = Nothing

End Sub

Private Sub AddColumnTransformation(ByRef oCustomTask As DTS.DataPumpTask, ByVal pTransformName As String, _
                      ByVal pSourceColName As String, ByVal pDestColName As String, _
                      ByVal pSourceDatatype As Long, _
                      ByVal pDestDataType As Long, ByVal pSourceFieldSize As Long, _
                      ByVal pDestFieldSize As Long, ByVal pSourceFlags As Long, _
                      ByVal pDestFlags As Long, ByVal pNullable As Boolean)
    
    Dim oColumn As DTS.Column
    Dim oTransformation As DTS.Transformation
    
      Set oTransformation = oCustomTask.Transformations.New("DTS.DataPumpTransformCopy.1")
        oTransformation.Name = pTransformName
        oTransformation.TransformFlags = 63
        oTransformation.ForceSourceBlobsBuffered = 0
        oTransformation.ForceBlobsInMemory = False
        oTransformation.InMemoryBlobSize = 1048576
        
          Set oColumn = oTransformation.SourceColumns.New(pSourceColName, 1)
            
            oColumn.Name = pSourceColName
            oColumn.Ordinal = 1
            oColumn.Flags = pSourceFlags
            oColumn.Size = 255
            oColumn.DataType = pSourceDatatype
            oColumn.Precision = 0
            oColumn.NumericScale = 0
            oColumn.Nullable = pNullable
          
          oTransformation.SourceColumns.Add oColumn
          
          Set oColumn = Nothing

          Set oColumn = oTransformation.DestinationColumns.New(pDestColName, 1)
            
            oColumn.Name = pDestColName
            oColumn.Ordinal = 1
            oColumn.Flags = pDestFlags
            oColumn.Size = 2
            oColumn.DataType = pDestDataType
            oColumn.Precision = 0
            oColumn.NumericScale = 0
            oColumn.Nullable = pNullable
          
          oTransformation.DestinationColumns.Add oColumn
          
          Set oColumn = Nothing
    
    oCustomTask.Transformations.Add oTransformation
    Set oTransformation = Nothing
        
End Sub

Private Sub AddPrecedenceConstraint(ByVal pStep As String, ByVal priorStep As String, _
                                      ByVal pConstraintResult As DTSStepExecResult)
  Dim oStep As DTS.Step
  Dim oPrecConstraint As DTS.PrecedenceConstraint
  
  Set oStep = goPackage.Steps(pStep)
  Set oPrecConstraint = oStep.PrecedenceConstraints.New(priorStep)
    oPrecConstraint.StepName = priorStep
    oPrecConstraint.PrecedenceBasis = 1
    oPrecConstraint.Value = pConstraintResult
    oStep.PrecedenceConstraints.Add oPrecConstraint
  Set oPrecConstraint = Nothing

End Sub
                                      
Private Sub CreateCustomTask(ByRef pTask As DTS.Task, ByVal pTaskName As String, _
                                         ByVal pTaskDescription As String, ByVal pSourceConnectionID As Long, _
                                         ByVal pSourceObjectName As String, ByVal pDestConnectionID As Long, _
                                         ByVal pDestObjectName As String, _
                                         Optional ByVal pSourceSQL As String = "", _
                                         Optional ByVal pDestSQL As String = "", _
                                         Optional ByVal pFetchBuffer As Long = 1, _
                                         Optional ByVal pProgressCount As Long = 1000, _
                                         Optional ByVal pFastLoad As Boolean = True)
          
    pTask.CustomTask.Properties("Name") = pTaskName
    pTask.CustomTask.Properties("Description") = pTaskDescription
    pTask.CustomTask.Properties("SourceConnectionID") = pSourceConnectionID
    pTask.CustomTask.Properties("DestinationConnectionID") = pDestConnectionID
    
    If pSourceSQL = "" Then
      pTask.CustomTask.Properties("SourceObjectName") = pSourceObjectName
    Else
      pTask.CustomTask.Properties("SourceSQLStatement") = pSourceSQL
    End If
    
    If pDestSQL = "" Then
      pTask.CustomTask.Properties("DestinationObjectName") = pDestObjectName
    Else
      pTask.CustomTask.Properties("DestinationSQLStatement") = pDestSQL
    End If
    
    pTask.CustomTask.Properties("ProgressRowCount") = pProgressCount
    pTask.CustomTask.Properties("MaximumErrorCount") = 0
    pTask.CustomTask.Properties("FetchBufferSize") = pFetchBuffer
    pTask.CustomTask.Properties("UseFastLoad") = pFastLoad
    pTask.CustomTask.Properties("InsertCommitSize") = 0
    pTask.CustomTask.Properties("ExceptionFileColumnDelimiter") = "|"
    pTask.CustomTask.Properties("ExceptionFileRowDelimiter") = vbCrLf
    pTask.CustomTask.Properties("AllowIdentityInserts") = False
    pTask.CustomTask.Properties("FirstRow") = 0
    pTask.CustomTask.Properties("LastRow") = 0
    pTask.CustomTask.Properties("FastLoadOptions") = 2
          
End Sub


Public Sub TransferVolumeFile(ByVal pSourceMDBPath As String, ByVal pDestpath As String, ByVal ErrlogDir As String)
  
  Set goPackage = New DTS.Package
  
  goPackage.Name = "TransferVolToMDB"
  goPackage.Description = "WinSort Access Volume Insert"
  goPackage.WriteCompletionStatusToNTEventLog = True
  goPackage.LogFileName = ErrlogDir & "\DTSErrlog.txt"
  goPackage.FailOnError = False
  goPackage.PackagePriorityClass = 3
  goPackage.MaxConcurrentSteps = 4
  goPackage.LineageOptions = 0
  goPackage.UseTransaction = True
  goPackage.TransactionIsolationLevel = 4096
  goPackage.AutoCommitTransaction = True
  goPackage.RepositoryMetadataOptions = 0
  goPackage.UseOLEDBServiceComponents = True
  
  Dim oConnection As DTS.Connection
  
  '[FROM] Access Connection
  Set oConnection = goPackage.Connections.New("Microsoft.Jet.OLEDB.4.0")
    oConnection.Name = "Microsoft Access FROM"
    oConnection.ID = 1
    oConnection.Reusable = True
    oConnection.ConnectImmediate = False
    oConnection.DataSource = pSourceMDBPath
    oConnection.UserID = "Admin"
    oConnection.ConnectionTimeout = 60
    oConnection.UseTrustedConnection = False
    oConnection.UseDSL = False
            
  goPackage.Connections.Add oConnection
  Set oConnection = Nothing
 
  '[TO] Access Connection
  Set oConnection = goPackage.Connections.New("Microsoft.Jet.OLEDB.4.0")
    oConnection.Name = "Microsoft Access To"
    oConnection.ID = 2
    oConnection.Reusable = True
    oConnection.ConnectImmediate = False
    oConnection.DataSource = pDestpath
    oConnection.UserID = "Admin"
    oConnection.ConnectionTimeout = 60
    oConnection.UseTrustedConnection = False
    oConnection.UseDSL = False
            
  goPackage.Connections.Add oConnection
  Set oConnection = Nothing
    
  Call CreateStep("STEP01_INSERT_NEW_RECORDS", "INSERT_NEW_RECORDS", "TASK01_INSERT_NEW_RECORDS", False)
  Call CreateStep("STEP02_DELETE_TEMP_RECORDS", "DELETE_TEMP_RECORDS", "TASK02_DELETE_TEMP_RECORDS", False)
  Call CreateStep("STEP03_INSERT_UPDATE_RECORDS", "INSERT_UPDATE_RECORDS", "TASK03_INSERT_UPDATE_RECORDS", False)
  Call CreateStep("STEP04_UPDATE_DEST_TABLE", "UPDATE_DEST_TABLE", "TASK04_UPDATE_DEST_TABLE", True)
 
  Call AddPrecedenceConstraint("STEP02_DELETE_TEMP_RECORDS", "STEP01_INSERT_NEW_RECORDS", DTSStepExecResult_Success)
  Call AddPrecedenceConstraint("STEP03_INSERT_UPDATE_RECORDS", "STEP02_DELETE_TEMP_RECORDS", DTSStepExecResult_Success)
  Call AddPrecedenceConstraint("STEP04_UPDATE_DEST_TABLE", "STEP03_INSERT_UPDATE_RECORDS", DTSStepExecResult_Success)
  
  'Transfer Volume Table Task ****************************************************************************************
    
  Dim oTask As DTS.Task
  Dim strSQL As String
  
  strSQL = "SELECT Pass, OldID, NewID, Pocket, Amount FROM TestTable WHERE FldUpdate = 'New'"
  
  Set oTask = goPackage.Tasks.New("DTSDataPumpTask")
    
  Call CreateCustomTask(oTask, "TASK01_INSERT_NEW_RECORDS", "Task01_Insert", 1, vbNullString, 2, "TestTable", strSQL)
    
    Call AddColumnTransformation(oTask.CustomTask, "TestTableCol_1", "Pass", "Pass", 130, 130, 8, 8, 102, 120, True)
    Call AddColumnTransformation(oTask.CustomTask, "TestTableCol_2", "OldID", "OldID", 130, 130, 4, 4, 102, 120, True)
    Call AddColumnTransformation(oTask.CustomTask, "TestTableCol_3", "NewID", "NewID", 130, 130, 4, 4, 102, 102, True)
    Call AddColumnTransformation(oTask.CustomTask, "TestTableCol_4", "Pocket", "Pocket", 130, 130, 2, 2, 102, 102, True)
    Call AddColumnTransformation(oTask.CustomTask, "TestTableCol_5", "Amount", "Amount", 130, 130, 3, 3, 32, 102, True)
        
  goPackage.Tasks.Add oTask
  Set oTask = Nothing
  
  
  strSQL = "DELETE * FROM TempTestTable"
  Call AddExecuteSQLTask("TASK02_DELETE_TEMP_RECORDS", "Task02_Delete", strSQL, 2)
  
  
  strSQL = "SELECT Pass, OldID, NewID, Pocket, Amount FROM TestTable WHERE FldUpdate = 'Update'"
  
  Set oTask = goPackage.Tasks.New("DTSDataPumpTask")
  
  Call CreateCustomTask(oTask, "TASK03_INSERT_UPDATE_RECORDS", "Task03_Insert", 1, vbNullString, 2, "TempTestTable", strSQL)
    
    Call AddColumnTransformation(oTask.CustomTask, "TestTableCol_1", "Pass", "Pass", 130, 130, 8, 8, 102, 120, True)
    Call AddColumnTransformation(oTask.CustomTask, "TestTableCol_2", "OldID", "OldID", 130, 130, 4, 4, 102, 120, True)
    Call AddColumnTransformation(oTask.CustomTask, "TestTableCol_3", "NewID", "NewID", 130, 130, 4, 4, 102, 102, True)
    Call AddColumnTransformation(oTask.CustomTask, "TestTableCol_4", "Pocket", "Pocket", 130, 130, 2, 2, 102, 102, True)
    Call AddColumnTransformation(oTask.CustomTask, "TestTableCol_5", "Amount", "Amount", 130, 130, 3, 3, 32, 102, True)
        
  goPackage.Tasks.Add oTask
  Set oTask = Nothing
  
  strSQL = "UPDATE TestTable INNER JOIN TempTestTable ON ((TestTable.Pass = TempTestTable.Pass) " _
          & "AND (TestTable.OldID = TempTestTable.OldID)) SET TestTable.NewID = TempTestTable.OldID, " _
          & "TestTable.Amount = TempTesttable.Amount"
          
  Call AddExecuteSQLTask("TASK04_UPDATE_DEST_TABLE", "Task04_Delete", strSQL, 2)
  
  m_Totaltasks = goPackage.Tasks.Count
  m_TasksCompleted = 0
  
End Sub

Public Function ExecutePackage() As Boolean
  Dim errStep As Long
    
  RaiseEvent PercentDone(0)
  
  goPackage.Execute
  
  ExecutePackage = True
  
  For errStep = 1 To goPackage.Steps.Count
    If goPackage.Steps(errStep).ExecutionResult = DTSStepExecResult_Failure Then
      ExecutePackage = False
      m_Cancel = True
      Debug.Print "Step " & goPackage.Steps(errStep).Name & " has failed " _
        & vbCrLf & "Affected tables have been rolled back to previous state."
    End If
  Next errStep
  
  RaiseEvent PercentDone(100)
  
  Set goPackage = Nothing
  
End Function

Private Sub Class_Terminate()
  Set goPackage = Nothing
End Sub

Private Sub goPackage_OnError(ByVal EventSource As String, ByVal ErrorCode As Long, ByVal source As String, ByVal Description As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal IDofInterfaceWithError As String, pbCancel As Boolean)
  RaiseEvent ErrorOccurred(ErrorCode, source, Description)
End Sub

Private Sub goPackage_OnFinish(ByVal EventSource As String)
  m_TasksCompleted = m_TasksCompleted + 1
  RaiseEvent PercentDone((m_TasksCompleted / m_Totaltasks) * 100)
  RaiseEvent Currenttask("")
  DoEvents
End Sub

Private Sub goPackage_OnProgress(ByVal EventSource As String, ByVal ProgressDescription As String, ByVal PercentComplete As Long, ByVal ProgressCountLow As Long, ByVal ProgressCountHigh As Long)
  RaiseEvent RowsCopied(ProgressDescription)
  DoEvents
End Sub

Private Sub goPackage_OnQueryCancel(ByVal EventSource As String, pbCancel As Boolean)
  If m_Cancel = True Then
    pbCancel = True
  Else
    pbCancel = False
  End If
End Sub

Private Sub goPackage_OnStart(ByVal EventSource As String)
  RaiseEvent Currenttask(EventSource)
  DoEvents
End Sub


