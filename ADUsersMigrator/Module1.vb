Imports System.IO
Module Module1
    Sub Main()
        Dim ADUsersReport As String = "C:\temp\users.txt"
        Dim ADGroupsReport As String = "C:\temp\groups.txt"
        Dim NewDomainName As String = "destination.org"
        Dim ADGroups As List(Of ADGroup) = GetADGroups(ADGroupsReport, NewDomainName)
        Dim ADUsers As List(Of ADUser) = GetADUsers(ADUsersReport, NewDomainName)
        Dim ADOrganizationalUnits As List(Of ADOrganizationalUnit) = GetADOrganizationalUnits(ADUsers, ADGroups, NewDomainName)
        Dim workingdir As String = Path.GetDirectoryName(ADGroupsReport) & "\"
        Using cmds As New StreamWriter(workingdir & "adcommands.txt")
            cmds.AutoFlush = True
            For Each ou As ADOrganizationalUnit In ADOrganizationalUnits
                cmds.WriteLine(ou.NewADOrganizationalUnitCMD)
            Next
            Using rpt As New StreamWriter(workingdir & "migratedusers.csv")
                rpt.AutoFlush = True
                rpt.WriteLine("DisplayName,SamAccountName,UserPrincipalName,AccountPassword")
                For Each u As ADUser In ADUsers
                    cmds.WriteLine(u.NewSetADUserCMDs)
                    rpt.WriteLine(u.ReportLine)
                Next
            End Using
            For Each g As ADGroup In ADGroups
                cmds.WriteLine(g.NewADGroupCMD)
                cmds.WriteLine(g.AddGroupMemeberCMDs)
            Next
        End Using
        Process.Start("notepad.exe", workingdir & "adcommands.txt")
        Process.Start("excel.exe", workingdir & "migratedusers.csv")
    End Sub
    Function GetADGroups(reportpath As String, newdomain As String) As List(Of ADGroup)
        Dim ADGroups As New List(Of ADGroup)
        Using sr As New StreamReader(reportpath)
            While sr.Peek() >= 0
                Dim s As String = sr.ReadLine
                Dim aGroup As New ADGroup(s, newdomain)
                If Not ADGroups.Contains(aGroup) Then
                    ADGroups.Add(aGroup)
                Else
                    Dim i As Integer = ADGroups.IndexOf(aGroup)
                    ADGroups.Item(i).Members.Add(aGroup.Members(0))
                End If
            End While
        End Using
        Return ADGroups
    End Function
    Function GetADUsers(UsersReport As String, newdomain As String) As List(Of ADUser)
        Dim ADUsers As New List(Of ADUser)
        Using sr As New StreamReader(UsersReport)
            Dim reportheader As List(Of String) = sr.ReadLine().Split(vbTab).ToList
            Dim currentuser As New ADUser
            Console.Write("Processing users")
            While sr.Peek() >= 0
                Dim s As String = sr.ReadLine
                Dim u As ADUser
                Do
                    u = New ADUser(s, newdomain, reportheader)
                Loop While currentuser.AccountPassword = u.AccountPassword
                Console.Write(".")
                currentuser = u
                ADUsers.Add(u)
            End While
        End Using
        Return ADUsers
    End Function
    Function GetADOrganizationalUnits(ADUsers As List(Of ADUser), ADGroups As List(Of ADGroup), newdomain As String) As List(Of ADOrganizationalUnit)
        Dim ADOrganizationalUnits As New List(Of ADOrganizationalUnit)
        Dim OUs As New List(Of String)
        For Each u As ADUser In ADUsers
            If Not OUs.Contains(u.Path) Then OUs.Add(u.Path)
        Next
        For Each g As ADGroup In ADGroups
            If Not OUs.Contains(g.Path) Then OUs.Add(g.Path)
        Next
        Dim newdn As String = ",DC=" & newdomain.Split(".")(0) & ",DC=" & newdomain.Split(".")(1)
        For Each s As String In OUs
            Dim oupath As New List(Of String)
            oupath = s.Replace("""", "").Replace(newdn, "").Split(",").ToList
            Do
                Dim ouname As String = oupath(0).Replace("OU=", "")
                oupath.RemoveAt(0)
                Dim newpath As String = ""
                For Each ou As String In oupath
                    newpath &= ou & ","
                Next
                Dim anOU As New ADOrganizationalUnit(ouname, newpath & newdn.TrimStart(CChar(",")))
                If Not ADOrganizationalUnits.Contains(anOU) Then ADOrganizationalUnits.Add(anOU)
            Loop While oupath.Count > 0
        Next
        ADOrganizationalUnits.Sort(Function(x, y) x.PathLengh.CompareTo(y.PathLengh))
        Return ADOrganizationalUnits
    End Function
End Module
Class ADOrganizationalUnit
    Private ReadOnly Name As String
    Private ReadOnly Path As String
    Private ReadOnly qt As String = Chr(34)
    Public Sub New(sName As String, sPath As String)
        Name = qt & sName & qt
        Path = qt & sPath & qt
    End Sub
    ReadOnly Property NewADOrganizationalUnitCMD As String
        Get
            Dim cmd As String = "New-ADOrganizationalUnit -Name " & Name & " -Path " & Path
            Return cmd
        End Get
    End Property
    ReadOnly Property PathLengh As Integer
        Get
            Return Path.Split(",").Count
        End Get
    End Property
    Public Overloads Overrides Function Equals(obj As Object) As Boolean
        If obj Is Nothing OrElse Not Me.GetType() Is obj.GetType() Then
            Return False
        End If
        Dim g As ADOrganizationalUnit = CType(obj, ADOrganizationalUnit)
        Return Me.Name = g.Name And Me.Path = g.Path
    End Function
End Class
Class ADGroup
    Private Name As String
    Public Path As String
    Private ReadOnly GroupCategory As String
    Private ReadOnly GroupScope As String
    Public Members As New List(Of String)
    Private ReadOnly qt As String = Chr(34)
    Public Sub New(GroupLine As String, newdomain As String)
        Dim newdn As String = "DC=" & newdomain.Split(".")(0) & ",DC=" & newdomain.Split(".")(1)
        Dim values As String() = GroupLine.Split(vbTab)
        Members.Add(qt & values(0).Substring(0, values(0).IndexOf("DC=")) & newdn & qt)
        Name = values(1)
        Path = qt & Mid(values(2), values(2).IndexOf("OU=") + 1, values(2).IndexOf("DC=") - values(2).IndexOf("OU=")) & newdn & qt
        GroupCategory = values(3)
        GroupScope = values(4)
    End Sub
    ReadOnly Property NewADGroupCMD As String
        Get
            Dim cmd As String = "New-ADGroup -Name " & qt & Name & qt & " -SamAccountName " & qt & Name & qt & " -GroupCategory " & GroupCategory & " -GroupScope " & GroupScope & " -DisplayName " & qt & Name & qt & " -Path " & Path
            Return cmd
        End Get
    End Property
    ReadOnly Property AddGroupMemeberCMDs As String
        Get
            Dim cmds As String = ""
            For Each m As String In Members
                cmds &= "Add-ADGroupMember -Identity " & qt & Name & qt & " -Members " & m & Environment.NewLine
            Next
            Return cmds.Remove(cmds.LastIndexOf(Environment.NewLine))
        End Get
    End Property
    Public Overloads Overrides Function Equals(obj As Object) As Boolean
        If obj Is Nothing OrElse Not Me.GetType() Is obj.GetType() Then
            Return False
        End If
        Dim g As ADGroup = CType(obj, ADGroup)
        Return Me.Name = g.Name
    End Function
End Class
Class ADUser
    Private ReadOnly SamAccountName As String
    Private ReadOnly DisplayName As String
    Public Path As String
    Private ReadOnly UserPrincipalName As String
    Public AccountPassword As String = RandomPassword(12)
    Private ReadOnly OldEmail As String
    Private ReadOnly qt As String = Chr(34)
    Public Properties As New List(Of ADUserProperty)
    Public Sub New()
    End Sub
    Public Sub New(UserLine As String, newdomain As String, reportheader As List(Of String))
        Dim values As String() = UserLine.Split(vbTab)
        Dim newdn As String = "DC=" & newdomain.Split(".")(0) & ",DC=" & newdomain.Split(".")(1)
        Dim i As Integer = 0
        For Each h As String In reportheader
            Select Case h.Trim(CChar(qt))
                Case "SamAccountName"
                    SamAccountName = values(i)
                    Properties.Add(New ADUserProperty(reportheader(i), SamAccountName))
                Case "DisplayName"
                    DisplayName = values(i)
                    Properties.Add(New ADUserProperty(reportheader(i), DisplayName))
                Case "DistinguishedName"
                    Path = qt & Mid(values(i), values(i).IndexOf("OU=") + 1, values(i).IndexOf("DC=") - values(i).IndexOf("OU=")) & newdn & qt
                    Properties.Add(New ADUserProperty("Path", Path))
                Case "UserPrincipalName"
                    UserPrincipalName = values(i).Substring(0, values(i).IndexOf("@") + 1) & newdomain & qt
                    Properties.Add(New ADUserProperty(reportheader(i), UserPrincipalName))
                Case "EmailAddress"
                    OldEmail = values(i)
                    Properties.Add(New ADUserProperty(reportheader(i), UserPrincipalName))
                Case Else
                    Properties.Add(New ADUserProperty(reportheader(i), values(i)))
            End Select
            i += 1
        Next
        Properties.Add(New ADUserProperty("AccountPassword", "(ConvertTo-SecureString " & "'" & AccountPassword & "'" & " -AsPlainText -Force)"))
        Properties.Add(New ADUserProperty("Enabled", "$true"))
    End Sub
    ReadOnly Property NewSetADUserCMDs As String
        Get
            Dim cmd As String = "New-ADUser"
            For Each p As ADUserProperty In Properties
                If p.IsSet Then cmd &= p.Name & p.Value
            Next
            cmd &= Environment.NewLine & "Set-ADUser " & SamAccountName & " -add @{ProxyAddresses=" & qt & "SMTP:" & UserPrincipalName.Trim(CChar(qt)) & ",smtp:" & OldEmail.Trim(CChar(qt)) & qt & " -split " & qt & "," & qt & "}"
            Return cmd
        End Get
    End Property
    ReadOnly Property ReportLine As String
        Get
            Return DisplayName & "," & SamAccountName & "," & UserPrincipalName & "," & AccountPassword
        End Get
    End Property
    Private Function RandomPassword(passedLength As Integer) As String
        Dim myLowercase As String = "abcdefghijklmnopqrstuvwxyz"
        Dim myUppercase As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        Dim myNumbers As String = "0123456789"
        Dim mySymbols As String = "!@#$%^&*_+"
        Dim myAllChars As String = myLowercase & myUppercase & myNumbers & mySymbols
        Dim myRandom As New System.Random
        Dim myPassword As String = ""
        myPassword = myPassword & myLowercase(myRandom.Next(0, myLowercase.Length)) & myLowercase(myRandom.Next(0, myLowercase.Length))
        myPassword = myPassword & myUppercase(myRandom.Next(0, myUppercase.Length)) & myUppercase(myRandom.Next(0, myUppercase.Length))
        myPassword = myPassword & myNumbers(myRandom.Next(0, myNumbers.Length)) & myNumbers(myRandom.Next(0, myNumbers.Length))
        myPassword = myPassword & mySymbols(myRandom.Next(0, mySymbols.Length)) & mySymbols(myRandom.Next(0, mySymbols.Length))
        For i As Integer = 0 To (passedLength - 9)
            myPassword &= myAllChars(myRandom.Next(0, myAllChars.Length))
        Next
        Dim strInput As String = myPassword
        Dim strOutput As String = ""
        Dim rand As New System.Random
        Dim intPlace As Integer
        While strInput.Length > 0
            intPlace = rand.Next(0, strInput.Length)
            strOutput += strInput.Substring(intPlace, 1)
            strInput = strInput.Remove(intPlace, 1)
        End While
        myPassword = strOutput
        Return myPassword
    End Function
End Class
Class ADUserProperty
    Public Name As String
    Public Value As String
    Private ReadOnly qt As String = Chr(34)
    Public Sub New(inName As String, inValue As String)
        Name = " -" & inName.Trim(CChar(qt)) & " "
        Value = inValue
    End Sub
    ReadOnly Property IsSet As Boolean
        Get
            If Value.Trim.Length > 0 Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property
End Class