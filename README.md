Option Compare Database
'Option Explicit

Public gUser As User

Public Function InitializeApplication()
'******************************************************************************
'* Init applicaton
'******************************************************************************
    Call ConfigureApplication
                            '# check current server path. If error-ask user to
                            '# input new path. 3 attempts only possible
    Call Server.CheckStringPathToServer
                            '# check actuality for local tables. If need=>update
    Call UpdateTablesInLocalDb
    
    If Not DirExists(AppSettings.StoragePath & "Images\") Then
        IO.CreateDirTree (AppSettings.StoragePath & "Images\")
    End If
                            '# set user
    Call Authorize
    
    Call SetVisible
    
    MsgBox "App optimized for screen 1920X1080 with 100% Scale", vbOKOnly, AppSettings.Header
    
    DoCmd.OpenForm "frmProjects", acNormal
End Function

Public Sub ConfigureApplication()
      
   With gIdPrefix
        .Department = "idDepartment"
        .Kpi = "idKpi"
        .Project = "IdProject"
        .Sla = "idSla"
        .Service = "idService"
        .Person = "idPerson"
        .ProjectService = "idProjectService"
        .Account = "idAccount"
   End With
   
    AppSettings.Header = "Sla Service Catalogue"
    AppSettings.AdminPassword = "admin"
    AppSettings.strDATE_FORMAT = "mm-dd-yyyy"
    AppSettings.StoragePath = CurrentProject.path & "\Files\"
    AppSettings.ImgFolderPath = CurrentProject.path & "\Files\Images\"
    AppSettings.CommentNoColour = RGB(200, 200, 200)
    AppSettings.CommentYesColour = RGB(100, 100, 255)

    UserSettings.DepartmentPathNameSymbols = CInt(DLookup("ParamValue", "tblSettings", "Parameter='DepartmentPathNameSymbols'"))
    UserSettings.ShowDepartmentAs = CInt(DLookup("ParamValue", "tblSettings", "Parameter='ShowDepartmentAs'"))
    UserSettings.ShowAccountAs = CInt(DLookup("ParamValue", "tblSettings", "Parameter='ShowAccountAs'"))
    
    Call CheckExcelReference
    Call checkOutlookReference
End Sub

Private Sub Authorize()
    Set gUser = RepoUser.GetCurrentUser
    If gUser.ID <= 0 Then
        Dim resp As Integer
        resp = MsgBox("User " & gUser.Login & " not registered in the database. Do you want to register " & gUser.Login & " now?", vbYesNo, AppSettings.Header)
        If resp = vbNo Then
            Exit Sub
        Else
            DoCmd.OpenForm "frmRegister", acNormal, , , , acDialog
        End If
    End If
    
End Sub

Private Sub SetVisible()
    If gUser.Role = "Admin" Then
        gAdminsTabVisible = -1
        gDevTabsVisible = 0
                        '# Hide navigation left bar
        DoCmd.NavigateTo "acNavigationCategoryObjectType"
        DoCmd.RunCommand acCmdWindowHide
        
     ElseIf gUser.Role = "Developer" Then
        gAdminsTabVisible = -1
        gDevTabsVisible = -1
        
     Else
        gAdminsTabVisible = 0
        gDevTabsVisible = 0
                        '# Hide navigation left bar
        DoCmd.NavigateTo "acNavigationCategoryObjectType"
        DoCmd.RunCommand acCmdWindowHide
    End If
                            '# if MyRibbon exist, invalidate it
    If Not (ribbon.MyRibbon Is Nothing) Then
        MyRibbon.Invalidate
    End If
End Sub
Public Sub UpdateTablesInLocalDb()
'******************************************************************************
'*      Compare timestamp for all tblCatalog on server and LocalDb.
'*      If localDb timeStamp<then on Serve => update local
'******************************************************************************

    Dim tblName As String
    Dim srvTableUpdated, lclTableUpdated As Date

    Dim rsServer, rsLocal As DAO.Recordset
    Set rsServer = Server.GetRecordset("Select * FROM tblDatabaseTablesList WHERE HasLocalCopyOnUI=-1", DAO.dbOpenSnapshot)
    Set rsLocal = LocalDb.GetRecordset("Select * FROM tblDatabaseTablesList WHERE HasLocalCopyOnUI=-1", DAO.dbOpenSnapshot)
    
    If rsServer.BOF And rsServer.EOF Then
        Exit Sub
    End If
    
    rsServer.MoveFirst
    Do While Not rsServer.EOF
        tblName = rsServer.Fields("tblName").value
        
        rsLocal.FindFirst "tblName='" & tblName & "'"
        If (rsLocal.NoMatch) _
                        Or (rsLocal!TimeStamp < rsServer!TimeStamp) _
                        Or Not Common.TableExists(tblName) Then
            CopyTableFromServer (tblName)
        End If
        
        rsServer.MoveNext
    Loop

Terminate:
    rsServer.Close
    rsLocal.Close

    Set rsServer = Nothing
    Set rsLocal = Nothing
End Sub

Private Sub CopyTableFromServer(tblName)
'# copy table from server to the localDb

    If Common.TableExists(tblName) Then
        DoCmd.DeleteObject acTable, tblName
    End If
    
    DoCmd.TransferDatabase acImport, "Microsoft Access", Server.GetConnectionString, acTable, _
                            tblName, tblName, 0
                            
    LocalDb.ExecuteSql ("Update tblDatabaseTablesList set [TimeStamp]=Now() WHERE tblName='" & tblName & "'")
End Sub
        
Private Sub CheckExcelReference()
    If refExists("excel") Then
        Exit Sub
    End If
    
    Dim refAddSuccess As Boolean
    Dim vbProj As VBIDE.VBProject
    Set vbProj = VBE.ActiveVBProject

    If Dir("C:\Program Files (x86)\Microsoft Office\Office12\EXCEL.exe") <> "" Then
        vbProj.References.AddFromFile ("C:\Program Files (x86)\Microsoft Office\Office12\EXCEL.exe")
        refAddSuccess = True
    End If
    If Dir("C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.exe") <> "" Then
        vbProj.References.AddFromFile ("C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.exe")
        refAddSuccess = True
    End If
    If Dir("C:\Program Files\Microsoft Office\Office15\EXCEL.exe") <> "" Then
        vbProj.References.AddFromFile ("C:\Program Files\Microsoft Office\Office15\EXCEL.exe")
        refAddSuccess = True
    End If
    If Dir("C:\Program Files\Microsoft Office\root\Office16\EXCEL.exe") <> "" Then
        vbProj.References.AddFromFile ("C:\Program Files\Microsoft Office\root\Office16\EXCEL.exe")
        refAddSuccess = True
    End If
    
    If Not refAddSuccess Then
        MsgBox "Error! App cant find MsExcel on this computer", vbCritical, AppSettings.Header
    End If
End Sub

Private Sub checkOutlookReference()
    If refExists("Outlook") Then
        Exit Sub
    End If
    
    Dim BoolExists As Boolean

    Dim vbProj As VBIDE.VBProject
    Set vbProj = VBE.ActiveVBProject

    If Dir("C:\Program Files (x86)\Microsoft Office\Office14\MSOUTL.OLB") <> "" Then
        vbProj.References.AddFromFile ("C:\Program Files (x86)\Microsoft Office\Office14\MSOUTL.OLB")
        refAddSuccess = True
    End If
    
    If Dir("C:\Program Files (x86)\Microsoft Office\Office15\MSOUTL.OLB") <> "" Then
        vbProj.References.AddFromFile ("C:\Program Files (x86)\Microsoft Office\Office15\MSOUTL.OLB")
        refAddSuccess = True
    End If
    
    If Dir("C:\Program Files\Microsoft Office\Office15\MSOUTL.OLB") <> "" Then
        vbProj.References.AddFromFile ("C:\Program Files\Microsoft Office\Office15\MSOUTL.OLB")
        refAddSuccess = True
    End If
    
    If Dir("C:\Program Files\Microsoft Office\root\Office16\MSOUTL.OLB") <> "" Then
        vbProj.References.AddFromFile ("C:\Program Files\Microsoft Office\root\Office16\MSOUTL.OLB")
        refAddSuccess = True
    End If
    
    If Not refAddSuccess Then
        MsgBox "Error! App cant find MS Outlook on this computer", vbCritical, AppSettings.Header
    End If
End Sub


Private Function refExists(ByVal refName As String) As Boolean
On Error GoTo ErrHandler

    Dim VBAEditor As VBIDE.VBE
    Dim vbProj As VBIDE.VBProject
    Dim chkRef As VBIDE.Reference
    Dim BoolExists As Boolean

    Set VBAEditor = Application.VBE
    Set vbProj = VBE.ActiveVBProject 'ActiveWorkbook.VBProject

    For Each chkRef In vbProj.References
        If chkRef.Name = refName Then
            refExists = True
            GoTo Terminate
        End If
    Next chkRef

Terminate:
    Set vbProj = Nothing
    Set VBAEditor = Nothing
    
Exit Function
ErrHandler:
    MsgBox Err.Source & " " & Err.Description & " in refExists"
End Function
