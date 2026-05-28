Attribute VB_Name = "WORKING_ON_DRAFT"
Option Explicit

Public Const DOWNLOAD_FOLDER As String = "C:\Users\User\Downloads\"
Public Const FROM_ADDRESS As String = "support@techroversolutions.com"

' ---------------- EXCHANGE DRAFTS ----------------
Function GetExchangeDraftsFolder() As Object
    Dim ns As Object, acc As Object
    Set ns = Application.GetNamespace("MAPI")
    
    For Each acc In ns.Accounts
        If LCase(acc.SmtpAddress) = FROM_ADDRESS Then
            Set GetExchangeDraftsFolder = acc.DeliveryStore.GetDefaultFolder(16)
            Exit Function
        End If
    Next acc
    Set GetExchangeDraftsFolder = Nothing
End Function

Sub ForceSendAccount(mail As Object)
    Dim acc As Object
    For Each acc In Application.Session.Accounts
        If LCase(acc.SmtpAddress) = FROM_ADDRESS Then
            Set mail.SendUsingAccount = acc
            Exit Sub
        End If
    Next acc
End Sub

' ---------------- SAFE ATTACH ----------------
Sub SafeAttach(mail As Object, filename As String)
    If Trim(filename) = "" Then Exit Sub
    If Dir(DOWNLOAD_FOLDER & filename) = "" Then Exit Sub
    mail.Attachments.Add DOWNLOAD_FOLDER & filename
End Sub

' ---------------- EXACT FILE NAME FINDER ----------------
Function FindExactFile(prefix As String, hotel As String, d As String) As String
    Dim h As String: h = LCase(hotel)

    ' 1. Sheraton Four Points DFW
    If h Like "sheraton four points dfw*" Then
        If prefix = "Daily" Then FindExactFile = "Daily-Labor-Report-Sheraton-4-Points-" & d & ".pdf": Exit Function
        If prefix = "MTD" Then FindExactFile = "MTD-Labor-Report-Sheraton-4-Points-" & d & ".pdf": Exit Function
        If prefix = "Revenue" Then FindExactFile = "Revenue-Report_Sheraton-Four-Points-" & d & ".pdf": Exit Function
    End If

    ' 2. Spanish Fort Fairfield
    If h Like "spanish fort fairfield*" Then
        If prefix = "Daily" Then FindExactFile = "Daily-Spanish-Fort-Fairfield-Inn-Suites-" & d & ".pdf": Exit Function
        If prefix = "MTD" Then FindExactFile = "MTD-Spanish-Fort-Fairfield-Inn-Suites-" & d & ".pdf": Exit Function
        If prefix = "Revenue" Then FindExactFile = "Revenue-Report_Spanish-Fort-Fairfield-" & d & ".pdf": Exit Function
    End If

    ' 3. Spanish Fort Courtyard
    If h Like "spanish fort courtyard*" Then
        If prefix = "Daily" Then FindExactFile = "Daily-Spanish-Fort-Courtyard-" & d & ".pdf": Exit Function
        If prefix = "MTD" Then FindExactFile = "MTD-Spanish-Fort-Courtyard-" & d & ".pdf": Exit Function
        If prefix = "Revenue" Then FindExactFile = "Revenue-Report_Spanish-Fort-Courtyard-" & d & ".pdf": Exit Function
    End If

    ' 4. Savannah Fairfield
    If h Like "savannah fairfield*" Then
        If prefix = "Daily" Then FindExactFile = "Daily-Savannah-Fairfield-Inn-Suites-" & d & ".pdf": Exit Function
        If prefix = "MTD" Then FindExactFile = "MTD-Savannah-Fairfield-Inn-Suites-" & d & ".pdf": Exit Function
        If prefix = "Revenue" Then FindExactFile = "Revenue-Report_Savannah-Fairfield-" & d & ".pdf": Exit Function
    End If

    ' 5. Pensacola Home2 Suites
    If h Like "pensacola-home*" Then
        If prefix = "Daily" Then FindExactFile = "Daily-PensacolaHome2Suites-" & d & ".pdf": Exit Function
        If prefix = "MTD" Then FindExactFile = "MTD-Pensacola-Home-" & d & ".pdf": Exit Function
        If prefix = "Revenue" Then FindExactFile = "Revenue-Report_Pensacola-Home-2-Suites-" & d & ".pdf": Exit Function
    End If

    ' 6. McAllen Residence Inn
    If h Like "mcallen residence inn*" Then
        If prefix = "Daily" Then FindExactFile = "Daily-McAllen-Residence-Inn-" & d & ".pdf": Exit Function
        If prefix = "MTD" Then FindExactFile = "MTD-McAllen-Residence-Inn-" & d & ".pdf": Exit Function
        If prefix = "Revenue" Then FindExactFile = "Revenue-Report_McAllen-Residence-Inn-" & d & ".pdf": Exit Function
    End If

    ' 7. KBS SQR II Q&C
    If h Like "kbs sqr*" Or h Like "*q&c*" Then
        If prefix = "Daily" Then FindExactFile = "Daily-KBS-SOR-II-Q-C-" & d & ".pdf": Exit Function
        If prefix = "MTD" Then FindExactFile = "MTD-KBS-SOR-II-Q-C-" & d & ".pdf": Exit Function
        If prefix = "Revenue" Then FindExactFile = "Revenue-Report_KBS-SOR-Q-C-" & d & ".pdf": Exit Function
    End If

    ' 8. Hampton Inn Boston
    If h Like "hampton inn boston*" Then
        If prefix = "Daily" Then FindExactFile = "Daily-Hampton-Inn-Boston-Braintree-" & d & ".pdf": Exit Function
        If prefix = "MTD" Then FindExactFile = "MTD-Hampton-Inn-Boston-Braintree-" & d & ".pdf": Exit Function
        If prefix = "Revenue" Then FindExactFile = "Revenue-Report_Hampton-Inn-Boston-Braintree-" & d & ".pdf": Exit Function
    End If

    ' 9. Embassy Suites Overland Park
    If h Like "embassy suites overland*" Then
        If prefix = "Daily" Then FindExactFile = "Daily-Embassy-Suites-Overland-Park-" & d & ".pdf": Exit Function
        If prefix = "MTD" Then FindExactFile = "MTD-Embassy-Suites-Overland-Park-" & d & ".pdf": Exit Function
        If prefix = "Revenue" Then FindExactFile = "Revenue-Report_Embassy-Suites-Overland-Park-NEW-" & d & ".pdf": Exit Function
    End If

    ' 10. Embassy Suites Indianapolis North
    If h Like "embassy suites indianapolis north*" Then
        If prefix = "Daily" Then FindExactFile = "Daily-Embassy-Suites-Indianapolis-North-" & d & ".pdf": Exit Function
        If prefix = "MTD" Then FindExactFile = "MTD-Embassy-Suites-Indianapolis-North-" & d & ".pdf": Exit Function
        If prefix = "Revenue" Then FindExactFile = "Revenue-Report_Embassy-Suites-IndianapolisNorth-" & d & ".pdf": Exit Function
    End If

    ' 11. DoubleTree Westminster
    If h Like "doubletree westminster*" Then
        If prefix = "Daily" Then FindExactFile = "Daily-DoubleTree-Westminster-" & d & ".pdf": Exit Function
        If prefix = "MTD" Then FindExactFile = "MTD-DoubleTree-Westminster-" & d & ".pdf": Exit Function
        If prefix = "Revenue" Then FindExactFile = "Revenue-Report_DoubleTree-Westminster-" & d & ".pdf": Exit Function
    End If

    ' 12. DoubleTree Harrisonburg
    If h Like "doubletree harrisonburg*" Then
        If prefix = "Daily" Then FindExactFile = "Daily-DoubleTree-Harrisonburg-" & d & ".pdf": Exit Function
        If prefix = "MTD" Then FindExactFile = "MTD-DoubleTree-Harrisonburg-" & d & ".pdf": Exit Function
        If prefix = "Revenue" Then FindExactFile = "Revenue-Report_Doubletree-Harrisonburg-" & d & ".pdf": Exit Function
    End If

    ' 13. D'iberville Courtyard
    If h Like "d'iberville courtyard*" Or h Like "diberville courtyard*" Then
        If prefix = "Daily" Then FindExactFile = "Daily-D-iberville-Courtyard-" & d & ".pdf": Exit Function
        If prefix = "MTD" Then FindExactFile = "MTD-D-iberville-Courtyard-" & d & ".pdf": Exit Function
        If prefix = "Revenue" Then FindExactFile = "Revenue-Report_D-iberville-Courtyard-" & d & ".pdf": Exit Function
    End If

    ' 14. Atlanta DoubleTree
    If h Like "atlanta doubletree*" Then
        If prefix = "Daily" Then FindExactFile = "Daily-Atlanta-DoubleTree-" & d & ".pdf": Exit Function
        If prefix = "MTD" Then FindExactFile = "MTD-Atlanta-DoubleTree-" & d & ".pdf": Exit Function
        If prefix = "Revenue" Then FindExactFile = "RevenueReport_AtlantaDoubleTree-" & d & ".pdf": Exit Function
    End If

    ' 15. EH Battery Hotel
    If h Like "eh battery hotel*" Then
        If prefix = "Daily" Then FindExactFile = "EH-Battery-Hotel-Daily-" & d & ".pdf": Exit Function
        If prefix = "MTD" Then FindExactFile = "MTD-EH-Battery-Hotel-" & d & ".pdf": Exit Function
        If prefix = "Revenue" Then FindExactFile = "Revenue-Report_Atlanta_Aloft-" & d & ".pdf": Exit Function
    End If

    FindExactFile = ""
End Function

' ============== BUILD CC ===============
Function BuildCC(ws As Object, r As Long) As String
    Dim c As String
    If ws.Cells(r, 15).Value <> "" Then c = ws.Cells(r, 15).Value
    If ws.Cells(r, 16).Value <> "" Then
        If c = "" Then c = ws.Cells(r, 16).Value Else c = c & ";" & ws.Cells(r, 16).Value
    End If
    BuildCC = c
End Function

' ============== COLLECT ALL GMs ===============
Function CollectAllGMs(ws As Object, r As Long) As String
    Dim i As Long, g As String
    For i = 4 To 13
        If ws.Cells(r, i).Value <> "" Then
            If g = "" Then g = ws.Cells(r, i).Value Else g = g & ";" & ws.Cells(r, i).Value
        End If
    Next i
    CollectAllGMs = g
End Function

' ============== COLLECT FIRST N GMs ===============
Function CollectFirstNGMs(ws As Object, r As Long, n As Long) As String
    Dim i As Long, g As String
    For i = 4 To 3 + n
        If Trim(ws.Cells(r, i).Value) <> "" Then
            If g = "" Then g = ws.Cells(r, i).Value Else g = g & ";" & ws.Cells(r, i).Value
        End If
    Next i
    CollectFirstNGMs = g
End Function

' ============== PM EMAIL ===============
Sub CreatePMEmail(hotel As String, pm1 As String, pm2 As String, cc As String, _
                  daily As String, mtd As String)

    Dim mail As Object
    Set mail = GetExchangeDraftsFolder.items.Add("IPM.Note")
    ForceSendAccount mail
    
    mail.Subject = "Labor Productivity Report - " & hotel
    
    mail.To = pm1 & IIf(pm2 <> "", ";" & pm2, "")
    If cc <> "" Then mail.cc = cc
    
    mail.HTMLBody = _
        "Hello Glenn / Jeff,<br><br>" & _
        "Please find the following reports for <b>" & hotel & "</b>:<br><br>" & _
        "<ul>" & _
        IIf(daily <> "", "<li>Daily Labor Report</li>", "") & _
        IIf(mtd <> "", "<li>MTD Labor Report</li>", "") & _
        "</ul><br>" & _
        "Thank you,<br><br>" & _
        "<a href='https://www.techrover.us/' target='_blank'>" & _
        "<img src='https://i.ibb.co/v4f3SQBM/logooo.png' height='40'><br></a>" & _
        "Support Team<br>Techrover Solutions<br>support@techroversolutions.com<br>"
    
    SafeAttach mail, daily
    SafeAttach mail, mtd
    
    mail.Save
End Sub

' ============== GM EMAIL ===============
Sub CreateGMEmail(hotel As String, gmList As String, cc As String, greeting As String, _
                  daily As String, mtd As String, revenue As String, subjectTxt As String)

    Dim mail As Object
    Set mail = GetExchangeDraftsFolder.items.Add("IPM.Note")
    ForceSendAccount mail
    
    If greeting = "" Then greeting = "Hello Glenn / Jeff"
    mail.Subject = subjectTxt
    mail.To = gmList
    If cc <> "" Then mail.cc = cc
    
    mail.HTMLBody = _
        greeting & ",<br><br>" & _
        "Please find the following reports for <b>" & hotel & "</b>:<br><br>" & _
        "<ul>" & _
        IIf(daily <> "", "<li>Daily Labor Report</li>", "") & _
        IIf(mtd <> "", "<li>MTD Labor Report</li>", "") & _
        IIf(revenue <> "", "<li>Revenue Report</li>", "") & _
        "</ul><br>" & _
        "Thank you,<br><br>" & _
        "<a href='https://www.techrover.us/' target='_blank'>" & _
        "<img src='https://i.ibb.co/v4f3SQBM/logooo.png' height='40'><br></a>" & _
        "Support Team<br>Techrover Solutions<br>support@techroversolutions.com<br>"
    
    SafeAttach mail, daily
    SafeAttach mail, mtd
    SafeAttach mail, revenue
    
    mail.Save
End Sub

' ================= MASTER RUN =================
Sub GenerateAllHotelEmails()

    Dim xl As Object, wb As Object, ws As Object
    Dim i As Long, hotel As String
    Dim pm As String, pm2 As String, cc As String
    Dim greeting As String, d As String
    Dim dailyFile As String, mtdFile As String, revenueFile As String
    
    d = Month(Date) & "-" & Day(Date) & "-" & Year(Date)
    
    Set xl = CreateObject("Excel.Application")
    Set wb = xl.Workbooks.Open(DOWNLOAD_FOLDER & "HotelContacts.csv")
    Set ws = wb.Sheets(1)
    
    i = 2
    Do While ws.Cells(i, 1).Value <> ""
        
        hotel = ws.Cells(i, 1).Value
        pm = ws.Cells(i, 2).Value
        pm2 = ws.Cells(i, 3).Value
        greeting = ws.Cells(i, 14).Value
        cc = BuildCC(ws, i)
        
        dailyFile = FindExactFile("Daily", hotel, d)
        mtdFile = FindExactFile("MTD", hotel, d)
        revenueFile = FindExactFile("Revenue", hotel, d)

        ' ===================== SPECIAL SHERATON LOGIC =====================
        If LCase(hotel) Like "sheraton four points dfw*" Then
        
            ' PM ? Daily + MTD
            CreatePMEmail hotel, pm, pm2, cc, dailyFile, mtdFile
            
            ' GM1..GM6 ? Revenue only
            CreateGMEmail hotel, CollectFirstNGMs(ws, i, 6), cc, greeting, _
                          "", "", revenueFile, _
                          "Revenue Report Four Points by Sheraton DFW Airport North"
            
            ' GM1..GM7 ? Daily + MTD
            CreateGMEmail hotel, CollectFirstNGMs(ws, i, 7), cc, greeting, _
                          dailyFile, mtdFile, "", _
                          "Labor Productivity Report Four Points by Sheraton DFW Airport North"
            
            GoTo NextHotel
        End If
        
        ' ===================== NORMAL HOTELS =====================
        CreatePMEmail hotel, pm, pm2, cc, dailyFile, mtdFile
        
        CreateGMEmail hotel, CollectAllGMs(ws, i), cc, greeting, _
                      dailyFile, mtdFile, revenueFile, _
                      "Labor Productivity & Revenue Reports - " & hotel

NextHotel:
        i = i + 1
    Loop
    
    wb.Close False
    xl.Quit
    
    MsgBox "ALL EMAILS CREATED SUCCESSFULLY IN DRAFTS", vbInformation
End Sub

