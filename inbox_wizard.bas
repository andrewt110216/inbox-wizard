Attribute VB_Name = "inbox_wizard"
Option Explicit
Option Base 1
Public myNameSpace As Outlook.Namespace, myExplorer As Outlook.Explorer, InboxPath As Outlook.MAPIFolder, SharedFolder As String
Public myEmail As Outlook.MailItem, mySender As String, myAttachment As Outlook.Attachment
Public ExcelLastSaved As Date, ExcelLastOpened As Date
Public MgrName() As String, MgrRADFolder() As String, MgrAltsFolder() As String, MgrAction() As String, MgrSubAction() As String, sDividerSub() As String, MgrEmail() As String
Public ClientConversion() As String, FileConditions() As String, FundConversion() As String, MultMgrs() As String
Public MgrCountEval() As Integer, MgrCountSkip() As Integer, MgrCountSave() As Integer, MgrCountMove() As Integer, MgrBooleanUsed() As Boolean
Public myIndex As Integer, Action As String, iIndexRows As Integer
Public TotSelect As Integer, TotEval As Integer, TotUnknown As Integer
Public RADSubfolder As String, SaveName As String, AltsPath As String
Public FullPath As String, sAddToPath As String, sFund As String, sClient As String, sYear As String, sFileName As String, iFind1 As Integer, iFind2 As Integer, iFind3 As Integer, iFind4 As Integer
Public fso As Object, Fileout As Object, strTracker As String, strRecent As String, strMsgBox As String, strFile As String
Public MgrSumEval As Integer, MgrSumSkip As Integer, MgrSumSave As Integer, MgrSumMove As Integer
Public Exists As Boolean, Jan As Boolean
Public i As Integer, j As Integer, n As Integer, iTemp As Integer

Sub InboxManager()
' Started September 13, 2017 by ACT.
' Goal is to make Old Inbox Manager more flexible, easier to add/remove managers to/from,
'     and to make the code more efficient and easier to understand

Set fso = CreateObject("Scripting.FileSystemObject")
Set myNameSpace = Application.GetNamespace("MAPI")
Set myExplorer = Application.ActiveExplorer
TotSelect = myExplorer.Selection.count

ExcelLastSaved = FileDateTime("U:\atracey\Inbox Manager\Data Input Tables.xlsx")
If ExcelLastSaved > ExcelLastOpened Then ExcelTables
Set InboxPath = myNameSpace.Folders(SharedFolder).Folders("Inbox")

ReDim MgrCountEval(1 To iIndexRows), MgrCountSkip(1 To iIndexRows), MgrCountSkip(1 To iIndexRows), MgrCountSave(1 To iIndexRows), MgrCountMove(1 To iIndexRows), MgrBooleanUsed(1 To iIndexRows)
strTracker = "": TotUnknown = 0: TotEval = 0: MgrSumEval = 0: MgrSumSkip = 0: MgrSumSave = 0: MgrSumMove = 0

CompileTracker ("Initiate")
'~~~~LOOP THROUGH EMAILS~~~~
For Each myEmail In myExplorer.Selection
    If myEmail.Class <> olMail Then Action = "Skip"
    mySender = myEmail.SenderEmailAddress
    If myEmail.Attachments.count > 0 Then Set myAttachment = myEmail.Attachments.Item(1)
    CompileTracker ("EmailNumber")
    myIndex = 0
    ' Based on Sender Email Address, determine the manager (myIndex)
    For i = 1 To iIndexRows
        For j = 1 To UBound(MgrEmail, 2)
            If MgrEmail(i, j) = "N/A" Then Exit For
            If mySender = MgrEmail(i, j) Then
                myIndex = i
                MgrBooleanUsed(i) = True
                MgrCountEval(i) = MgrCountEval(i) + 1
                Exit For
            End If
        Next j
        If myIndex = i Then Exit For
    Next i
    ' Complete the Action for the Manager
    If myIndex = 0 Then
        Action = "Unknown"
    Else
        Action = MgrAction(i)
    End If
    If Action = "Function" Then Action = ActionFunction()
    CompileTracker ("Output")
    Select Case Action
        Case "Unknown"
            CompileTracker ("Unknown")
        Case "Skip"
            CompileTracker ("Skip")
        Case "FileOnly"
                myEmail.Move InboxPath.Folders(MgrRADFolder(i))
                CompileTracker ("Moved")
        Case "FileSubfolder"
                RADSubfolder = RADSubfolder_Manager3
                myEmail.Move InboxPath.Folders(MgrRADFolder(i)).Folders(RADSubfolder)
                CompileTracker ("MovedSub")
        Case Else
            myAttachment.SaveAsFile Action
            CompileTracker ("Saved")
            'myEmail.Move InboxPath.Folders(MgrRADFolder(i))
            CompileTracker ("Moved")
    End Select
Next myEmail

'~~~~ SUMMARIZE AND END ~~~~
strMsgBox = "Inbox Manager evaluated the " & TotEval & " selected e-Mail(s)." & vbCr & vbCr & "Inbox Manager skipped " & TotUnknown & _
    " e-mail(s) because the Sender was not recognized." & vbCr & vbCr & "For recognized email addresses, Inbox Manager completed the following actions, by Manager:" & vbCr & vbCr
For i = 1 To iIndexRows
    If MgrBooleanUsed(i) = True Then strMsgBox = strMsgBox & MgrName(i) & ":" & vbCr & vbTab & vbTab & "Evaluated: " & _
        MgrCountEval(i) & ".   Skipped: " & MgrCountSkip(i) & ".   Saved: " & MgrCountSave(i) & ".   Filed: " & MgrCountMove(i) & vbCr
    MgrSumEval = MgrSumEval + MgrCountEval(i)
    MgrSumSkip = MgrSumSkip + MgrCountSkip(i)
    MgrSumSave = MgrSumSave + MgrCountSave(i)
    MgrSumMove = MgrSumMove + MgrCountMove(i)
Next i

strMsgBox = strMsgBox & "TOTAL for Recognized Email Addresses:" & vbCr & vbTab & vbTab & "Evaluated: " & MgrSumEval & ".   Skipped: " & _
            MgrSumSkip & ".   Saved: " & MgrSumSave & ".   Filed: " & MgrSumMove & vbCr
MsgBox strMsgBox, , "Summary of Actions Taken by Inbox Manager"
CompileTracker ("End")
Set fso = CreateObject("Scripting.FileSystemObject")
Set Fileout = fso.CreateTextFile("U:\atracey\Inbox Manager\TeamRAD Manager - Run Results\New Inbox Manager\DATE_" & month(Date) & "-" & Day(Date) & "-" & year(Date) & _
            "_TIME_" & Hour(Time) & "h" & Minute(Time) & "m" & Second(Time) & "s.txt", True, True)
Fileout.Write strTracker + vbCr + strMsgBox
Fileout.Close
End Sub

Sub CompileTracker(ActionType As String)
    Select Case ActionType
        Case "Initiate"
            strRecent = vbCr & "Details of actions taken for execution at: " & Date + Time
        Case "EmailNumber"
            TotEval = TotEval + 1
            strRecent = "  Email #" & TotEval & " of " & TotSelect & ":" & vbCr & "    From: " & mySender & vbCr & "    Subject: " & Left(myEmail.Subject, 70)
        Case "Output"
            strRecent = "       Action Output used was: " & Action
        Case "Unknown"
            strRecent = "         Unknown Sender. E-mail skipped."
            TotUnknown = TotUnknown + 1
        Case "Skip"
            strRecent = "         Skipped e-mail from recognized sender."
            MgrCountSkip(i) = MgrCountSkip(i) + 1
        Case "Moved"
            strRecent = "         Moved e-mail to: " & MgrRADFolder(i)
            MgrCountMove(i) = MgrCountMove(i) + 1
        Case "MovedSub"
            strRecent = "         Moved e-mail to: " & MgrRADFolder(i) & " > " & RADSubfolder
            MgrCountMove(i) = MgrCountMove(i) + 1
        Case "Saved"
            strRecent = "         Saved Attachment."
            MgrCountSave(i) = MgrCountSave(i) + 1
        Case "End"
            strRecent = "End of execution at: " & Date + Time & vbCr
        Case "DNE"
            strRecent = "       Folder Path Does Not Exist: " & FullPath & "\" & sAddToPath
        Case "NewClient"
            strRecent = "Client not found in Excel tables: " & sClient
            MsgBox strRecent & vbCr & vbCr & "Inbox Manager will skip this email and continue. Please add the client for future use.", vbOKOnly, "Inbox Manager Message"
    End Select
    Debug.Print strRecent
    strTracker = vbCr + strTracker + strRecent
End Sub

Sub ExcelTables()
' Populates array variables with manager information stored in Excel File
Dim xlApp As Object
Set xlApp = CreateObject("Excel.Application")
strFile = "U:\atracey\Inbox Manager\Inbox Manager Data Input Tables.xlsx"
ExcelLastOpened = Now()

With xlApp
    .Visible = True
    .Workbooks.Open strFile
    .Worksheets("General Definitions").Activate
        AltsPath = .Range("AltsPath").Value
        SharedFolder = .Range("SharedFolder").Value
        iIndexRows = .Range("iIndexRows").Value
    .Worksheets("Manager Variables").Activate
        ReDim MgrName(1 To iIndexRows), MgrRADFolder(1 To iIndexRows), MgrAltsFolder(1 To iIndexRows), MgrAction(1 To iIndexRows)
        ReDim MgrSubAction(1 To iIndexRows), sDividerSub(1 To iIndexRows), MgrEmail(1 To iIndexRows, 6)
        ReDim MgrBooleanUsed(1 To iIndexRows), MgrCountEval(1 To iIndexRows), MgrCountSkip(1 To iIndexRows)
        ReDim MgrCountSave(1 To iIndexRows), MgrCountMove(1 To iIndexRows)
        For i = 1 To iIndexRows
            .Range("TempArea").Formula = "=match(""MgrName(i)""," & .Range("sManagerVariables").Value & ",0)"
                MgrName(i) = .Range("IndexTable").Cells(i, .Range("TempArea").Value)
            .Range("TempArea").Formula = "=match(""MgrRADFolder(i)""," & .Range("sManagerVariables").Value & ",0)"
                MgrRADFolder(i) = .Range("IndexTable").Cells(i, .Range("TempArea").Value)
            .Range("TempArea").Formula = "=match(""MgrAltsFolder(i)""," & .Range("sManagerVariables").Value & ",0)"
                MgrAltsFolder(i) = .Range("IndexTable").Cells(i, .Range("TempArea").Value)
            .Range("TempArea").Formula = "=match(""MgrAction(i)""," & .Range("sManagerVariables").Value & ",0)"
                MgrAction(i) = .Range("IndexTable").Cells(i, .Range("TempArea").Value)
            .Range("TempArea").Formula = "=match(""MgrSubAction(i)""," & .Range("sManagerVariables").Value & ",0)"
                MgrSubAction(i) = .Range("IndexTable").Cells(i, .Range("TempArea").Value)
            .Range("TempArea").Formula = "=match(""MgrEmail(i,1)""," & .Range("sManagerVariables").Value & ",0)"
                MgrEmail(i, 1) = .Range("IndexTable").Cells(i, .Range("TempArea").Value)
            .Range("TempArea").Formula = "=match(""MgrEmail(i,2)""," & .Range("sManagerVariables").Value & ",0)"
                MgrEmail(i, 2) = .Range("IndexTable").Cells(i, .Range("TempArea").Value)
            .Range("TempArea").Formula = "=match(""MgrEmail(i,3)""," & .Range("sManagerVariables").Value & ",0)"
                MgrEmail(i, 3) = .Range("IndexTable").Cells(i, .Range("TempArea").Value)
            .Range("TempArea").Formula = "=match(""MgrEmail(i,4)""," & .Range("sManagerVariables").Value & ",0)"
                MgrEmail(i, 4) = .Range("IndexTable").Cells(i, .Range("TempArea").Value)
            .Range("TempArea").Formula = "=match(""MgrEmail(i,5)""," & .Range("sManagerVariables").Value & ",0)"
                MgrEmail(i, 5) = .Range("IndexTable").Cells(i, .Range("TempArea").Value)
            .Range("TempArea").Formula = "=match(""MgrEmail(i,6)""," & .Range("sManagerVariables").Value & ",0)"
                MgrEmail(i, 6) = .Range("IndexTable").Cells(i, .Range("TempArea").Value)
        Next i
        .Range("TempArea").ClearContents
    .Worksheets("FileOnly Conditions").Activate
        ReDim FileConditions(1 To iIndexRows, 1 To .Range("iFileRows").Value, 2)
        For i = 1 To iIndexRows
            .Range("TempArea2").Formula = "=match(" & i & "," & .Range("sFileIndexes").Value & ",0)"
            If IsError(.Range("TempArea2")) = False Then
                iTemp = .Range("TempArea2").Value
                For n = 1 To .Range("iFileRows").Value
                    If .Range("FileOnlyTable").Cells(n, iTemp) = "N/A" Then Exit For
                    FileConditions(i, n, 1) = .Range("FileOnlyTable").Cells(n, iTemp)
                    FileConditions(i, n, 2) = .Range("FileOnlyTable").Cells(n, iTemp + 1)
                Next n
            End If
        Next i
        .Range("TempArea2").ClearContents
    .Worksheets("Client Names").Activate
        ReDim ClientConversion(1 To iIndexRows, 1 To .Range("iClientRows").Value, 4)
        For i = 1 To iIndexRows
            .Range("TempArea1").Formula = "=match(" & i & "," & .Range("sClientIndexes").Value & ",0)"
            If IsError(.Range("TempArea1")) = False Then
                iTemp = .Range("TempArea1").Value
                For n = 1 To .Range("iClientRows").Value
                    If .Range("ClientTable").Cells(n, iTemp) = "N/A" Then Exit For
                    ClientConversion(i, n, 1) = .Range("ClientTable").Cells(n, iTemp)
                    ClientConversion(i, n, 2) = .Range("ClientTable").Cells(n, iTemp + 1)
                    ClientConversion(i, n, 3) = .Range("ClientTable").Cells(n, iTemp + 2)
                Next n
            End If
        Next i
        .Range("TempArea1").ClearContents
    .Worksheets("Fund Names").Activate
        ReDim FundConversion(1 To iIndexRows, 1 To .Range("iFundRows").Value, 2)
        For i = 1 To iIndexRows
            .Range("TempArea3").Formula = "=match(" & i & "," & .Range("sFundIndexes").Value & ",0)"
            If IsError(.Range("TempArea3")) = False Then
                iTemp = .Range("TempArea3").Value
                For n = 1 To .Range("iFundRows").Value
                    If .Range("FundTable").Cells(n, iTemp) = "N/A" Then Exit For
                    FundConversion(i, n, 1) = .Range("FundTable").Cells(n, iTemp)
                    FundConversion(i, n, 2) = .Range("FundTable").Cells(n, iTemp + 1)
                Next n
            End If
        Next i
        .Range("TempArea3").ClearContents
    .ActiveWorkbook.Close (False)
    .Visible = False
End With
End Sub

Function ActionFunction()
' Saving attachments relies on the following structure of the folder system ("Alts"):
'   AltsPath \ Manager Folder \ Fund Folder, if applicable (i.e. "Manager1 \ FundA vs. FundB") \ Client Folder \ Year Folder

' EVALUATE FILE ONLY CONDITIONS
Dim myString As String
n = 1
If FileConditions(i, 1, 1) <> "" Then
    Do Until FileConditions(i, n, 1) = ""
        Select Case FileConditions(i, n, 2)
            Case "Subject"
                myString = myEmail.Subject
            Case "Body"
                myString = myEmail.Body
            Case "Attachment"
                myString = myAttachment.DisplayName
            Case Else
                MsgBox ("There is an unusable input in the Excel Tables." & vbCr & vbCr & "See FileConditions(" & i & "," & n & "2).")
                ActionFunction = "Skip"
        End Select
        If InStr(myString, FileConditions(i, n, 1)) > 0 Then
            ActionFunction = "FileOnly"
            Exit Function
        End If
        n = n + 1
    Loop
End If

' EVALUATE SUBACTION TO EXTRACT FUND AND CLIENT NAMES and COMPILE sAddToPath
sAddToPath = "": sFileName = ""
FullPath = AltsPath & "\" & MgrAltsFolder(i)
Select Case MgrSubAction(i)
    Case "Manager1"
        Call Extract_Manager1
    Case "State Street"
        Call Extract_StateStreet
    Case "Manager2"
        Call Extract_Manager2
    Case "N/A"
    Case Else
End Select
If FundConversion(i, 1, 1) <> "" Then
    For n = 1 To UBound(FundConversion, 2)
      If sFund = FundConversion(i, n, 1) Then
        sFund = FundConversion(i, n, 2)
        sAddToPath = sFund & "\"
        Exit For
      End If
    Next n
End If
If ClientConversion(i, 1, 1) <> "" Then
    For n = 1 To UBound(ClientConversion, 2)
        If sClient = ClientConversion(i, n, 1) Then
            sClient = ClientConversion(i, n, 2)
            If ClientConversion(i, n, 3) = "N/A" Then
                sAddToPath = sAddToPath & sClient & "\" & sYear
                sFileName = myAttachment.DisplayName
            Else
                sAddToPath = sAddToPath & sClient & "\" & sYear
                sFileName = sFileName & ClientConversion(i, n, 3)
            End If
            Exit For
        End If
        If n = UBound(ClientConversion, 2) Then
            FullPath = "Skip"
            CompileTracker ("NewClient")
        End If
    Next n
End If

' DOES FOLDER PATH EXIST
If FullPath <> "Skip" And sFileName <> "" And DoesPathExist Then
    ActionFunction = FullPath & "\" & sAddToPath & "\" & sFileName
Else
    ActionFunction = "Skip"
End If
End Function

Sub Extract_StateStreet()
' Extract Fund Name and Client Name and year from attachment
If InStr(myAttachment.DisplayName, "_") > 0 Then
    iFind1 = InStr(1, myAttachment.DisplayName, "_")
    iFind2 = iFind1 + InStr(Right(myAttachment.DisplayName, Len(myAttachment.DisplayName) - iFind1), "_")
    sFund = Left(myAttachment.DisplayName, iFind1 - 1)
    sFileName = Left(myAttachment.DisplayName, iFind2)
    If InStr(sFileName, "2015") > 0 Then sYear = "2015"
    If InStr(sFileName, "2016") > 0 Then sYear = "2016"
    If InStr(sFileName, "2017") > 0 Then sYear = "2017"
    If InStr(sFileName, "2018") > 0 Then sYear = "2018"
    If InStr(sFileName, "2019") > 0 Then sYear = "2019"
    If InStr(sFileName, "2020") > 0 Then sYear = "2020"
    If InStr(sFileName, "2021") > 0 Then sYear = "2021"
    If InStr(sFileName, "2022") > 0 Then sYear = "2022"
    If InStr(sFileName, "2023") > 0 Then sYear = "2023"
    If InStr(sFileName, "2024") > 0 Then sYear = "2024"
    If InStr(sFileName, "2025") > 0 Then sYear = "2025"
    If InStr(sFileName, "2026") > 0 Then sYear = "2026"
    sClient = Right(myAttachment.DisplayName, Len(myAttachment.DisplayName) - iFind2)
    sClient = Left(sClient, Len(sClient) - 4)
Else
    FullPath = "Skip"
    Exit Sub
End If
End Sub

Sub Extract_Manager1()
' Extract Fund Name and Client Name and year from attachment
If InStr(myAttachment.DisplayName, "_") > 0 Then
    iFind1 = InStr(1, myAttachment.DisplayName, "_")
    iFind2 = iFind1 + InStr(Right(myAttachment.DisplayName, Len(myAttachment.DisplayName) - iFind1), "_")
    iFind3 = iFind2 + InStr(Right(myAttachment.DisplayName, Len(myAttachment.DisplayName) - iFind2), "_")
    iFind4 = iFind3 + InStr(Right(myAttachment.DisplayName, Len(myAttachment.DisplayName) - iFind3), "_")
    sFund = Left(myAttachment.DisplayName, iFind1 - 1)
    sFileName = Right(Left(myAttachment.DisplayName, iFind4 - 1), 8)
    Jan = False
    If Left(sFileName, 2) = "01" Then Jan = True
    If (Right(sFileName, 4) = "2015" And Jan = False) Or (Right(sFileName, 4) = "2015" And Jan = True) Then sYear = "2015"
    If (Right(sFileName, 4) = "2016" And Jan = False) Or (Right(sFileName, 4) = "2016" And Jan = True) Then sYear = "2016"
    If (Right(sFileName, 4) = "2017" And Jan = False) Or (Right(sFileName, 4) = "2017" And Jan = True) Then sYear = "2017"
    If (Right(sFileName, 4) = "2018" And Jan = False) Or (Right(sFileName, 4) = "2018" And Jan = True) Then sYear = "2018"
    If (Right(sFileName, 4) = "2019" And Jan = False) Or (Right(sFileName, 4) = "2019" And Jan = True) Then sYear = "2019"
    If (Right(sFileName, 4) = "2020" And Jan = False) Or (Right(sFileName, 4) = "2020" And Jan = True) Then sYear = "2020"
    If (Right(sFileName, 4) = "2021" And Jan = False) Or (Right(sFileName, 4) = "2021" And Jan = True) Then sYear = "2021"
    If (Right(sFileName, 4) = "2022" And Jan = False) Or (Right(sFileName, 4) = "2022" And Jan = True) Then sYear = "2022"
    If (Right(sFileName, 4) = "2023" And Jan = False) Or (Right(sFileName, 4) = "2023" And Jan = True) Then sYear = "2023"
    If (Right(sFileName, 4) = "2024" And Jan = False) Or (Right(sFileName, 4) = "2024" And Jan = True) Then sYear = "2024"
    If (Right(sFileName, 4) = "2025" And Jan = False) Or (Right(sFileName, 4) = "2025" And Jan = True) Then sYear = "2025"
    If sFund = "13LP" Then sYear = Right(Left(myAttachment.DisplayName, 9), 4)
    sClient = Right(myAttachment.DisplayName, Len(myAttachment.DisplayName) - iFind2)
    sClient = Left(sClient, Len(sClient) - 4)
    sClient = Left(sClient, 6) & Right(sClient, 7)
Else
    FullPath = "Skip"
    Exit Sub
End If

End Sub

Sub Extract_Manager2()
' Extract Client Name and year from attachment
If InStr(myAttachment.DisplayName, "_") > 0 Then
    iFind1 = InStr(1, myAttachment.DisplayName, "_")
    sClient = Left(myAttachment.DisplayName, iFind1 - 1)
    sYear = Left(Right(myAttachment.DisplayName, 14), 4)
Else
    FullPath = "Skip"
    Exit Sub
End If
End Sub

Function RADSubfolder_Manager3()
Dim sFund As String
If InStr(myEmail.Subject, "K-1") > 0 Or InStr(myEmail.Subject, "Tax Estimate") > 0 Then
    RADSubfolder_Manager3 = "K-1 All funds"
    Exit Function
End If
' Read Subject and extract Fund
If InStr(1, myEmail.Subject, ",") > 0 Then
    iFind1 = InStr(1, myEmail.Subject, ",")
    sFund = Left(myEmail.Subject, iFind1 - 1)
Else
    RADSubfolder_Manager3 = "Skip"
    Exit Function
End If
' Convert from sFund to Subfolder Name
Select Case sFund
    Case "Manager3 QP-Fund I"
        RADSubfolder_Manager3 = "Manager3 I"
    Case "Manager3 Fund III"
        RADSubfolder_Manager3 = "Manager3 III"
    Case "Manager3 Fund IV"
        RADSubfolder_Manager3 = "Manager3 IV"
    Case "Manager3 Fund V"
        RADSubfolder_Manager3 = "Manager3 V"
    Case "Manager3 Fund VI"
        RADSubfolder_Manager3 = "Manager3 VI"
    Case "Manager3 Fund VII"
        RADSubfolder_Manager3 = "Manager3 VII"
    Case "Manager3 Fund VIII"
        RADSubfolder_Manager3 = "Manager3 VIII"
    Case "Manager3 Fund IX"
        RADSubfolder_Manager3 = "Manager3 IX"
    Case "Manager3 Fund X"
        RADSubfolder_Manager3 = "Manager3 X"
    Case "Manager3 Fund XI"
        RADSubfolder_Manager3 = "Manager3 XI"
    Case "Manager3 Fund XII"
        RADSubfolder_Manager3 = "Manager3 XII"
    Case "Manager3Direct II"
        RADSubfolder_Manager3 = "Manager3 Direct II"
    Case "Manager3 Secondary Opportunity Fund"
        RADSubfolder_Manager3 = "Manager3 Sec"
    Case "Manager3 Secondary Opportunity Fund II"
        RADSubfolder_Manager3 = "Manager3 Sec II"
    Case "Manager3 SBIC Opportunities Fund"
        RADSubfolder_Manager3 = "Manager3 SBIC"
    Case "Manager3 FF Small Buyout Co-Investment Fund"
        RADSubfolder_Manager3 = "Manager3 FF Small Buyout"
    Case "Manager3 FF Small Buyout Co-Investment Fund II"
        RADSubfolder_Manager3 = "Manager3 FF Small Buyout"
    Case Else
        RADSubfolder_Manager3 = ""
End Select
End Function

Function DoesPathExist()
If (fso.FolderExists(FullPath & "\" & sAddToPath)) Then
     Exists = True
Else
     Exists = False
     CompileTracker ("DNE")
End If
DoesPathExist = Exists
End Function


