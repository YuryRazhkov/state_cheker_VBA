Public Sub vbajson()

    Dim http As Object
    Dim authKey As String
    Dim RespArray() As String
    Dim new_states As String
    Dim unp As Long
    Dim a As Byte
    Dim cur_date As Date
    Dim date_now As String
    Dim it_count As String
    Dim ws As Worksheet
    Dim hstr As Integer, wss As Object
    Dim lgs As Integer
    Dim cur_state As String
    Dim sub_name As String
    Dim start_mac, stop_mac As String
    Dim strArray() As String
    Dim type_sub As String
    
    Close hstr
    Close lgs
         
    Application.ScreenUpdating = True
     
    cur_date = Now
    date_now = Format(Now, "yyyy-mm-dd-hh-mm-ss")
 
    
    Sheets("reestr").Select
    
    
     ActiveSheet.Unprotect
       
      On Error Resume Next
      ActiveSheet.ShowAllData
    
    Range("K7") = "last chek"
    Range("L7") = Sheets("last").Range("C2").Value
    
    Range("K2") = "start"
    start_mac = date_now
    Range("L2") = start_mac
    
    
    Range("A2").Select
    
    
    it_count = 0
    
    hstr = FreeFile
    Open ThisWorkbook.Path & "\history.txt" For Append As hstr

    
    Do While ActiveCell > 0
    unn = ActiveCell
    a = 0

    On Error GoTo Stroka
         Set http = CreateObject("MSXML2.XMLHTTP")
    With http
        .Open "GET", "http://egr.gov.by/api/v2/egr/getBaseInfoByRegNum/" & unn, False
        .setRequestHeader "Content-type", "application/json"
        .setRequestHeader "Accept", "application/json"
        .setRequestHeader "Authorization", "Bearer " & authKey
        .Send
    
    End With
   
         
    If http.Status <> 200 Then
        lgs = FreeFile
        
        Open ThisWorkbook.Path & "\log.txt" For Append As lgs
        Print #lgs, cur_date; "EGR is not responding  " & http.Status
        Close lgs
    
        MsgBox ("EGR is not responding. Try later " & http.Status)
        Exit Sub
    End If
        
   
   
   cur_state = http.responsetext
   cur_state = Split(cur_state, "vnsostk"":""")(1)
   cur_state = Split(cur_state, """,")(0)
   
   
   type_sub = http.responsetext
   type_sub = Split(type_sub, """nkvob"":")(1)
   type_sub = Split(type_sub, ",")(0)
   
  
  
    If type_sub = 1 Then
         On Error GoTo Stroka
         Set http = CreateObject("MSXML2.XMLHTTP")
    With http
        .Open "GET", "http://egr.gov.by/api/v2/egr/getJurNamesByRegNum/" & unn, False
        .setRequestHeader "Content-type", "application/json"
        .setRequestHeader "Accept", "application/json"
        .setRequestHeader "Authorization", "Bearer " & authKey
        .Send
     End With
    
     sub_name = Split(http.responsetext, """vn"":""")(1)
     sub_name = Split(sub_name, ",")(0)
     sub_name = Replace(sub_name, "\", "")
     sub_name = Replace(sub_name, """", "")
     
           
    Else
          On Error GoTo Stroka
         Set http = CreateObject("MSXML2.XMLHTTP")
    With http
        .Open "GET", "http://egr.gov.by/api/v2/egr/getIPFIOByRegNum/" & unn, False
        .setRequestHeader "Content-type", "application/json"
        .setRequestHeader "Accept", "application/json"
        .setRequestHeader "Authorization", "Bearer " & authKey
        .Send
     End With
     
     sub_name = Split(http.responsetext, "vfio"":""")(1)
     sub_name = Split(sub_name, ",")(0)
     sub_name = Replace(sub_name, """", "")
     
    
    End If
        
        
  ActiveCell.Offset(0, 1) = sub_name
  ActiveCell.Offset(0, 2) = Format(Now, "dd.mm.yyyy")
  ActiveCell.Offset(0, 3) = cur_state
  ActiveCell.Offset(0, 4) = "=RC[-1]=VLOOKUP(RC[-4],last!C[-4]:C[-1],4,0)"
  ActiveCell.Offset(0, 5) = "=HYPERLINK(""https://egr.gov.by/egrmobile/information?pan=""&RC[-5],""EGR link"")"
  

        
    ActiveCell.Offset(1, 0).Select
       
    date_now = Format(Now, "yyyy-mm-dd-hh-mm-ss")
    Range("K3") = "stop"
       
    stop_mac = date_now
    Range("L3") = stop_mac
      
    it_count = it_count + 1
    Range("K4") = "processed clients"
    Range("L4") = it_count
    
       
       
         Print #hstr, date_now; ";"; Application.UserName; ";"; unn; ";"; sub_name; ";"; cur_state
       
       
    Loop
     
   Close hstr
   
 
      Range("E2:E3500").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Sheets("last").Visible = True
    Sheets("last").Select
    
    Cells.Select
    Selection.Delete
    Sheets("reestr").Select
    Range("A1:E3500").Select
    Selection.Copy
    Sheets("last").Select
    Range("A1").Select
    ActiveSheet.Paste
    Selection.ColumnWidth = 25
    Range("A1").Select
    Sheets("last").Select
    ActiveWindow.SelectedSheets.Visible = False
    
   
    
    Sheets("reestr").Select
    Range("K5") = "new states"
    Range("L5").Select
    
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-7],FALSE)"
    new_states = ActiveCell.Value
    
        
    Columns("E:E").Select
    date_now = Format(Now, "yyyy-mm-dd-hh-mm-ss")
    
    
    Range("A1").Select
   
     
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, AllowInsertingColumns:=True, AllowInsertingRows _
        :=True, AllowInsertingHyperlinks:=True, AllowDeletingColumns:=True, _
        AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True
     
    ThisWorkbook.Save
     
     Close hstr
        lgs = FreeFile
        Open ThisWorkbook.Path & "\log.txt" For Append As lgs
        Print #lgs, cur_date; Application.UserName; " Processed clients: " & it_count & ". New states: " & new_states & " time: "; start_mac; " - "; stop_mac
        
        Close lgs
    
    MsgBox "Processed clients: " & it_count & ". New states: " & new_states
    Exit Sub
    
Stroka:
    
    Close hstr
    Close lgs
    
    ActiveCell.Select
    
        With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    
    Selection.Font.Bold = True
    
    
    lgs = FreeFile
      Open ThisWorkbook.Path & "\log.txt" For Append As lgs
    Print #lgs, cur_date; Application.UserName; "Processed clients:  "; it_count; " "; Err.Description
    Close lgs
    
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, AllowInsertingColumns:=True, AllowInsertingRows _
        :=True, AllowInsertingHyperlinks:=True, AllowDeletingColumns:=True, _
        AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True
    
    MsgBox Err.Description & "Chek connection and valid data"
    
End Sub



