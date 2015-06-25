'=========================================================================== 
' Checks all accounts to determine what needs to be disabled. 
' If LastLogonTimeStamp is Null and object is older than specified date, it is disabled and moved. 
' If account has been used, but not within duration specified, it is disabled and moved. 
' If account is already disabled it is left where it is. 
' Created 23/7/09 by Grant Brunton 
'=========================================================================== 
 
'=========================================================================== 
' BEGIN USER VARIABLES 
'=========================================================================== 
 
' Flag to enable the disabling and moving of unused accounts 
' 1 - Will Disable and move accounts 
' 0 - Will create ouput log only 
bDisable=1 
 
' Number of days before an account is deemed inactive 
' Accounts that haven't been logged in for this amount of days are selected 
iLogonDays=45
 
' LDAP Location of OUs to search for accounts 
' LDAP location format eg: "OU=ORGANIZATION" 
strSearchOU="OU=DEP,OU=USUARIOS,OU=BRASIL,OU=SITES" 
 
' Search depth to find users 
' Use "OneLevel" for the specified OU only or "Subtree" to search all child OUs as well. 
strSearchDepth="SUBTREE" 

' Location of new OU to move disabled user accounts to 
' eg: "OU=Disable,dc=contoso,dc=com" 
strNewOU="OU=INATIVOS,OU=USUARIOS,OU=BRASIL,OU=SITES" 
 
' Log file path (include trailing \ ) 
' Use either full directory path or relational to script directory 
strLogPath="C:\ScriptBloqueio\logs\" 
 
' Error log file name prefix (tab delimited text file. Name will be appended with date and .err extension) 
strErrorLog="DisabledAccounts_" 
 
' Output log file name prefix (tab delimited text file. Name will be appended with date and .txt extension) 
strOutputLog="DisabledAccountsE_" 
 
'=========================================================================== 
' END USER VARIABLES 
'=========================================================================== 
 
 
'=========================================================================== 
' MAIN CODE BEGINS 
'=========================================================================== 
sDate = Year(Now()) & Right("0" & Month(Now()), 2) & Right("0" & Day(Now()), 2)  
Set oFSO=CreateObject("Scripting.FileSystemObject") 
If Not oFSO.FolderExists(strLogPath) Then CreateFolder(strLogPath) 
Set output=oFSO.CreateTextFile(strLogPath & strOutputLog & sDate & ".txt") 
Set errlog=oFSO.CreateTextFile(strLogPath & strErrorLog & sDate & ".err") 
output.WriteLine "Sam Account Name" &vbTab& "LDAP Path" &vbTab& "Last Logon Date" &vbTab& "Date Created" &vbTab& "Home Directory" 
errlog.WriteLine "Sam Account Name" &vbTab& "LDAP Path" &vbTab& "Problem" &vbTab& "Error" 
 
Set rootDSE = GetObject("LDAP://rootDSE") 
Set objConnection = CreateObject("ADODB.Connection") 
objConnection.Open "Provider=ADsDSOObject;" 
Set ObjCommand = CreateObject("ADODB.Command") 
ObjCommand.ActiveConnection = objConnection 
ObjCommand.Properties("Page Size") = 10 
DSEroot=rootDSE.Get("DefaultNamingContext") 
 
Set objNewOU = GetObject("LDAP://" & strNewOU & "," & DSEroot) 
ObjCommand.CommandText = "<LDAP://" & strSearchOU & "," & DSEroot & ">;(&(objectClass=User)(objectcategory=Person));adspath;" & strSearchDepth  

Set objRecordset = ObjCommand.Execute 
 
On Error Resume Next 
 
While Not objRecordset.EOF 
    LastLogon = Null 
    intLogonTime = Null 
  
    Set objUser=GetObject(objRecordset.fields("adspath")) 
 
    If DateDiff("d",objUser.WhenCreated,Now) > iLogonDays Then 
        Set objLogon=objUser.Get("lastlogontimestamp") 
        If Err.Number <> 0 Then 
            WriteError objUser, "Get LastLogon Failed" 
            DisableAccount objUser, "Never" 
       Else 
           intLogonTime = objLogon.HighPart * (2^32) + objLogon.LowPart 
           intLogonTime = intLogonTime / (60 * 10000000) 
           intLogonTime = intLogonTime / 1440 
           LastLogon=intLogonTime+#1/1/1601# 

           If DateDiff("d",LastLogon,Now) > iLogonDays Then 
               DisableAccount objUser, LastLogon 
	    End If 
        End If 
    End If 
    WriteError objUser, "Unknown Error" 
    objRecordset.MoveNext 
 		
Wend 
'=========================================================================== 
' MAIN CODE ENDS 
'=========================================================================== 
 
 
'=========================================================================== 
' SUBROUTINES 
'=========================================================================== 
Sub CreateFolder( strPath ) 
    If Not oFSO.FolderExists( oFSO.GetParentFolderName(strPath) ) Then Call CreateFolder( oFSO.GetParentFolderName(strPath) ) 
    oFSO.CreateFolder( strPath ) 
End Sub 
 
Sub DisableAccount( objUser, lastLogon ) 
    On Error Resume Next 
    If bDisable <> 0 Then 
        If objUser.accountdisabled=False Then 
           objUser.accountdisabled=True 
           objUser.SetInfo 
           WriteError objUser, "Disable Account Failed" 
           objNewOU.MoveHere objUser.adspath, "CN="&objUser.CN 
           WriteError objUser, "Account Move Failed" 
       Else 
           Err.Raise 1,,"Account already disabled. User not moved." 
           WriteError objUser, "Disable Account Failed" 
       End If 
   End If 
   output.WriteLine objUser.samaccountname &vbTab& objUser.adspath &vbTab& lastLogon &vbTab& objUser.whencreated &vbTab& objUser.homedirectory 
End Sub 
output.Close

Sub WriteError( objUser, strProblem ) 
    If Err.Number <> 0 Then 
        errlog.WriteLine objUser.samaccountname &vbTab& objUser.adspath &vbTab& strProblem &vbTab& Replace(Err.Description,vbCrlf,"") 
        Err.Clear 
    End If 
End Sub 

Set objEmail = CreateObject("CDO.Message")

objEmail.From = "BloqueioContas@<dominio>" 

objEmail.Subject = "Rotina - Contas Desabilitadas por inatividade em 45 dias " 
objEmail.To = "<email>@<dominio>"

objEmail.Textbody = "Rotina - Contas Desabilitadas por inatividade em 45 dias - Vide Anexo!!!"
objEmail.AddAttachment strLogPath & strOutputLog & sDate & ".txt" 

objEmail.Configuration.Fields.Item _
 ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objEmail.Configuration.Fields.Item _
 ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = _
"172.25.0.200"
objEmail.Configuration.Fields.Item _
 ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objEmail.Configuration.Fields.Update
objEmail.Configuration.Fields.Update
objEmail.Send
 
'=========================================================================== 
' END SUBROUTINES 
'===========================================================================