Attribute VB_Name = "GoodsReceiptPO"
Option Explicit
Const THIS_WB_NAME As String = "GRPO_TOOL.xlsm"
Public date_GR As String
Public only_new As Boolean
Public mokc_GRPO As New OneKeyCls




Sub ma3()
'1.检查运行条件是否满足：a SAP 是否打开，b 服务器是否能够访问
Dim majjl As New C_AJJL
Dim mdate As Date



Dim str1 As String, str2 As String, str3 As String

Dim mfso As New CFSO

str1 = "Goods Receipt Purchase Order - Er'gang Zhao"
str2 = "Z:\42_Logistics\01_Project Data"

Dim i As Integer, j As Integer



If majjl.my_findwindow(str1) = 0 Then
MsgBox "please open SAP migo in the first: " & str1
Exit Sub
End If

If mfso.folderexists(str2) = False Then
MsgBox "Can not open:" & str2
Exit Sub
End If
only_new = True
ma2
ma1

For i = 1 To 1000

mdate = now()

str3 = Format(now(), "hh:mm")
If Right(str3, 1) = "0" Or Right(str3, 1) = "5" Then

DoEvents
only_new = False
ma2
ma1
DoEvents
Application.StatusBar = now() & "DELAY 15min"
End If
Application.StatusBar = Format(now(), "YYYY-MM-DD HH:MM:SS")
delay 57000
If Left(str3, 2) > "18" Then
Exit Sub
End If
Next









End Sub

Sub ma1()
'1.判断SAP 收货页面是否打开，如果没有打开则终止
'2.逐条 选择 TOBE_GR 的进行收货，并将 TOBE_GR 修改为 FINISH
Dim xx As Integer
Dim yy As Integer
Dim mfso As New CFSO

Dim b1 As Boolean

Dim sap_wh As String

Dim kk As Integer

Dim str1  As String, str2 As String, str3 As String, str4 As String, str5 As String


Dim para1 As String, para2 As String, para3 As String, para4 As String, para5 As String, para6 As String
Dim s_cannotpost As String


Dim i As Integer

Dim majjl As New C_AJJL



Dim ws_piclib As Worksheet



'Dim i As Integer
Dim i_last As Integer
Dim fdn As String: Dim fln As String: Dim flfp As String
Dim temp_s As String

Dim windowname As String
Dim windowname2 As String
windowname = "Goods Receipt Purchase Order - Er'gang Zhao"
If majjl.my_findwindow(windowname) = 0 Then
MsgBox "please open SAP migo in the first: " & windowname
Exit Sub
Else
sap_wh = CStr(majjl.my_findwindow(windowname))
End If
Dim ws As Worksheet
Set ws = Workbooks(THIS_WB_NAME).Worksheets("GRPO_DATA")

'给按键精灵赋值图片，check_ok ,post, prpo
Dim usname As String
usname = Environ("Computername")
para1 = get_para_rg(ws.Range("A2:Z2"), "PIC_1 for click *.bmp" & usname, "N")
If mfso.FileExists(para1) = False Or Not (para1 Like "*.bmp") Then
para1 = get_para_rg(ws.Range("A2:Z2"), "PIC_1 for click *.bmp" & usname, "Y")
End If
majjl.Add_Pic para1

para2 = get_para_rg(ws.Range("A2:Z2"), "PIC_2 for click *.bmp" & usname, "N")
If mfso.FileExists(para2) = False Or Not (para2 Like "*.bmp") Then
para2 = get_para_rg(ws.Range("A2:Z2"), "PIC_2 for click *.bmp" & usname, "Y")
End If
majjl.Add_Pic para2

para3 = get_para_rg(ws.Range("A2:Z2"), "PIC_3 for click *.bmp" & usname, "N")
If mfso.FileExists(para3) = False Or Not (para3 Like "*.bmp") Then
para3 = get_para_rg(ws.Range("A2:Z2"), "PIC_3 for click *.bmp" & usname, "Y")
End If
majjl.Add_Pic para3


'Can not post ,点击 Display logs 的 “关闭”
para4 = get_para_rg(ws.Range("A2:Z2"), "PIC_4 for click *.bmp" & usname, "N")
If mfso.FileExists(para4) = False Or Not (para4 Like "*.bmp") Then
para4 = get_para_rg(ws.Range("A2:Z2"), "PIC_4 for click *.bmp" & usname, "Y")
End If
majjl.Add_Pic para4
'Can not post ,点击 Display logs 的 “关闭”

' 点F12 后，Restart 窗口 点NO
para5 = get_para_rg(ws.Range("A2:Z2"), "PIC_5 for click *.bmp" & usname, "N")
If mfso.FileExists(para5) = False Or Not (para5 Like "*.bmp") Then
para5 = get_para_rg(ws.Range("A2:Z2"), "PIC_5 for click *.bmp" & usname, "Y")
End If
majjl.Add_Pic para5
' 点F12 后，Restart 窗口 点NO

'Z:\24_Temp\PA_Logs\TOOLS\GoodsReceiptPurchaseOder\AI\PurchaseOrder_red.bmp


para6 = get_para_rg(ws.Range("A2:Z2"), "PIC_6 for click *.bmp" & usname, "N")
If mfso.FileExists(para6) = False Or Not (para6 Like "*.bmp") Then
para6 = get_para_rg(ws.Range("A2:Z2"), "PIC_6 for click *.bmp" & usname, "Y")
End If
majjl.Add_Pic para6




'给按键精灵赋值图片，check_ok ,post, prpo



i_last = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row
For i = 4 To i_last

s_cannotpost = ""


    str1 = ws.Range("A" & i)
    If str1 = "TOBE_GR" Then
    str2 = ws.Range("B" & i)
    str3 = ws.Range("E" & i)
    ws.Range("A" & i) = "DOING"
    str4 = P_SPLIT(str2, "-", 0)
    str5 = P_SPLIT(str2, "-", -1)
    
    
    
    
'提示信息
Application.StatusBar = "(" & i & "/" & i_last & ")" & str1 & str2 & str3 & str4 & str5
If i - 5 > 1 Then
ActiveWindow.ScrollRow = i - 5
End If
 
'提示信息
If InStr(str3, "HOLD") > 0 Then
ws.Range("A" & i) = "HOLD"
ElseIf InStr(str3, "RETURN") > 0 Then
ws.Range("A" & i) = "RETURN"
Else



    'MEMO_LClick_PIC windowname, m_PicCol.Item("PurchaseOrder.bmp").m_StdPicture, 200, 15
   If majjl.L_CLICK_PIC(windowname, para1, 300, 20) = False Then
   If majjl.L_CLICK_PIC(windowname, para6, 300, 20) = False Then
   
   
   MsgBox "can not find " & para6
   Exit Sub
   End If
   
   
   End If
   
    
    SendKeys "{end}": SendKeys "{backspace}{backspace}{backspace}{backspace}{backspace}{backspace}{backspace}{backspace}{backspace}{backspace}{backspace}":
    SendKeys str4: SendKeys "{TAB}": delay 100: SendKeys "{end}": SendKeys "{backspace}{backspace}{backspace}": delay 100
    SendKeys str5: SendKeys "{TAB}": delay 100: SendKeys "{TAB}": delay 100: SendKeys "{TAB}": delay 100: SendKeys "{TAB}": delay 100: SendKeys "{TAB}": delay 100: SendKeys "{TAB}": delay 100: SendKeys "{TAB}": delay 100: SendKeys "{TAB}": delay 100: SendKeys "{TAB}"
    'SendKeys str3: SendKeys "{TAB}": delay 100: SendKeys "{TAB}": delay 100: SendKeys "{TAB}":
    'SendKeys str3: SendKeys "{ENTER}"
    
    If my_sendkeys(str3) = False Then
    MsgBox "can not paste!"
    Exit Sub
    End If
    
    SendKeys "{TAB}": delay 100: SendKeys "{TAB}": delay 100: SendKeys "{TAB}"
    
    If my_sendkeys(str3) = False Then
    MsgBox "can not paste!"
    Exit Sub
    End If
    
    
    SendKeys "{ENTER}"
    
    'MEMO_LClick_PIC windowname, m_PicCol.Item("check ok.bmp").m_StdPicture, 10, 10
         delay 4000
         
         
     
      windowname2 = "Goods Receipt Purchase Order " & str4 & " - Er'gang Zhao"
     
           For kk = 1 To 10
           If majjl.my_findwindow(windowname2) > 0 Then
           Exit For
           Else
           delay 1000
           End If
          Next
          
          
          
     
     
    ' windowname = "Goods Receipt Purchase Order 3000046800 - Er'gang Zhao"
   
    
    If majjl.my_findwindow(windowname2) = 0 Then
        '=============================
        '如果不存在指定MO-item ，则 重登 migo，并 置 TOBE_GR=>DOUBLE_CHECK
        
        
        
        
        If majjl.my_findwindow(sap_wh) = 0 Then Exit Sub
        'ws.Range("A" & i) = "DoubleCheck1"
        
        ws.Range("A" & i) = "ALREADY FINISH"
        SendKeys "{f12}"
        delay 3000
        
        ' MsgBox "EXITSUB"
       ' Exit Sub
       
        
        
        '=============================
        
        
           
        'If MsgBox("Continue NEXT?(Y/N)", vbYesNo) = vbNo Then
        'Exit Sub
        'End If
        
      
        
    Else
         
         '如果点击失败，则 重登Migo，并置
         xx = 0
         yy = 0
      ' b1 = MEMO_LClick_PIC(windowname2, m_PicCol.Item("LOGIN_OK.bmp").m_StdPicture, 10, 10)
          majjl.L_CLICK_PIC windowname2, para2, 15, 15
      
       
        
         delay 100
        'b1 = MEMO_LClick_PIC(windowname2, m_PicCol.Item("Post.bmp").m_StdPicture, 10, 15)
        majjl.L_CLICK_PIC windowname2, para3, 10, 15
        
        delay 3000
       
       '动态等待，窗口消失  windowname2
         For kk = 1 To 10
           If majjl.my_findwindow(windowname) > 0 Then
           Exit For
           Else
           delay 1000
           End If
          Next
        If majjl.my_findwindow(windowname) = 0 Then
 
        
        'Application.DisplayAlerts = False
        
        'Workbooks(THIS_WB_NAME).Save
        'Exit Sub
        '1.关闭 Display logs
       delay 1000
       
       '  majjl.L_CLICK_PIC "Display logs", para4, 10, 10
        '2.发送F12
       delay 1000
       delay 5000
        If majjl.my_findwindow(sap_wh) = 0 Then Exit Sub
        delay 1000
        
        If majjl.my_findwindow("Display logs") = 0 Then
        MsgBox " Display logs  =0 exit "
        Exit Sub
        End If
        
        
        SendKeys "{f12}"
        delay 3000
       '
        
        'If majjl.my_findwindow("Display logs") > 0 Then
      '  MsgBox "ex 0"
      '  Exit Sub
       ' End If
        
        
       '   If majjl.my_findwindow(sap_wh) = 0 Then Exit Sub
          
        'If majjl.my_findwindow("Restart") = 0 Then Exit Sub
        
       
        delay 1000
        SendKeys "{f12}"
        
         'majjl.L_CLICK_PIC windowname2, para6, 10, 10
         delay 3000
       
       If majjl.my_findwindow("Restart") = 0 Then
       MsgBox "EX1"
       Exit Sub
       End If
       
        
        
        '3.Restart 点No
         'majjl.L_CLICK_PIC "Restart", para5, 10, 10
       majjl.L_CLICK_WIN "Restart", 337, 167
       
       delay 3000
       'If majjl.my_findwindow("Restart") > 0 Then Exit Sub
        
        
        ' majjl.L_CLICK_PIC windowname2, para5, 10, 10
         
        delay 1000
       
    
    's_cannotpost = "CAN NOT POST"
           ws.Range("A" & i) = "CAN NOT POST"
        
        
        
        
        End If
       '动态等待，窗口消失  windowname2
       
     
         'myact
         majjl.my_actwindow "GRPO_TOOL.xlsm"
        majjl.delay 1000
        
         
        ' delay 500
         
    
          If ws.Range("A" & i) = "CAN NOT POST" Then
          Else
          If ws.Range("A" & i) <> "DOING" Then
          Exit Sub
          End If
           ws.Range("A" & i) = "FINISH"
           ws.Range("H" & i) = now()
          End If
          
         
           'Workbooks(THIS_WB_NAME).Save
           
                
       ' If MsgBox("Continue NEXT?(Y/N)", vbYesNo) = vbNo Then
       ' Exit Sub
       ' End If
      End If
      
        End If
        
 
  
End If
   
Next
Application.DisplayAlerts = False
Workbooks(THIS_WB_NAME).Save


Application.ScreenUpdating = True
Application.StatusBar = THIS_WB_NAME & "SAVED!"


End Sub
Function myact()
'delay 5000
On Error Resume Next
delay 300
AppActivate "GRPO_TOOL.xlsm"
'delay 2000

Application.WindowState = xlMaximized

'delay 2000
'AppActivate "GRPO_TOOL.xlsm"
'delay 2000

End Function

Sub ma2()
'1.访问服务器 Z:\24_Temp\PA_Logs\TOOLS\GoodsReceiptPurchaseOder\input\42_logistics
'2.获取其中全部的 receiving report*.xlsm
'3.将这些文件录入工作表  receiving report_CN.xxxxx 并比较其新旧
'4.如果新，或者是有更新，则将其只读 禁用宏的方式打开，并复制 开始日期（最近一周，或者三天）的 DATA也中收货记录 到 GRPO_DATA 页中。


Dim str1  As String, str2 As String, str3 As String, str4 As String
Dim para1 As String, para2 As String, para3 As String, para4 As String
Dim i1 As Integer, i2 As Integer, i3 As Integer, i4 As Integer, i5 As Integer, i6 As Integer
                    

Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet
Dim wb1 As Workbook, wb2 As Workbook

Dim b_c As Boolean

If init_para = False Then Exit Sub
Set ws1 = get_ws(Workbooks(THIS_WB_NAME), "receiving report_CN.xxxxx")
Set ws2 = get_ws(Workbooks(THIS_WB_NAME), "GRPO_DATA")
Dim date1 As Date
Dim date2 As Date




Dim mokc As New OneKeyCls
Dim mfso As New CFSO

'============================
 to_mokc mokc, "FD", "$", get_para_rg(Workbooks(THIS_WB_NAME).Worksheets("PARA").Range("A1:Z1"), "FDN_ROOT_GRPO", "N")
 to_mokc mokc, "FD_TPF", "$", "*"
 to_mokc mokc, "FD_DEL", "$", "OLD"
 to_mokc mokc, "FL_TPF", "$", "receiving report*.xlsm"

If Len(date_GR) = 0 Then
para1 = InputBox("Input the date YYYY-MM-DD for GoodsReceipt.", "", Format(now(), "YYYY-MM-DD"))
date_GR = para1
Else
para1 = date_GR
End If



If Len(para1) < 8 Then Exit Sub

date1 = CDate(para1)

If date1 < now() - 10 Or date1 > now() Then
MsgBox "Must be Recent 10 days"
Exit Sub
End If



str1 = PF_Z(get_Single_File(mokc, mfso))
Do While Len(str1) > 0
b_c = False

If mokc_GRPO.Item(ws1.Name).Item("BODY").Item(str1) Is Nothing Then
'MsgBox ""
i2 = ws1.UsedRange.Rows(ws1.UsedRange.Rows.Count).Row + 1
ws1.Cells(i2, 2) = str1
ws1.Cells(i2, 4) = mfso.Datelastmodify(str1)
ws1.Cells(i2, 3) = mfso.Datelastmodify(str1)
'b_c = True


Else

str2 = mokc_GRPO.Item(ws1.Name).Item("BODY").Item(str1).Item(1).Item(1).Item("DATE_LAST_MODIFY").Key
str3 = mfso.Datelastmodify(str1)
i2 = mokc_GRPO.Item(ws1.Name).Item("BODY").Item(str1).Item(1).Item(1).Key

If str2 <> str3 Then
'b_c = True


ws1.Cells(i2, 4) = str3
ws1.Cells(i2, 3) = str3


End If


If CDate(ws1.Cells(i2, 3)) >= date1 Then
b_c = True
End If




End If




If b_c Then


Application.ScreenUpdating = True
Application.StatusBar = str1
DoEvents
Application.ScreenUpdating = False


     
     '只读禁宏方法打开电子表格
     If wb_open_ONLY_READ(wb1, str1, only_new) Then
     
     'MsgBox ""
     
     If ws_exist(wb1, "DATA") = True Then
     Set ws3 = get_ws(wb1, "DATA")
     
     i3 = ws3.UsedRange.Rows(ws3.UsedRange.Rows.Count).Row
     
     For i4 = 1 To i3
     
     i5 = i3 - i4 + 1
   '  MsgBox ""
     
     date2 = my_cdate(ws3.Cells(i5, 5))
 
     
     'MsgBox ""
     If Len(ws3.Cells(i5, 5)) = 0 Then
     
     
     ElseIf date2 = date1 Then
     
     '复制
    ' MsgBox ""
     
     str4 = ws3.Cells(i5, 1)
     
         If mokc_GRPO.Item(ws2.Name).Item("BODY").Item(str4) Is Nothing Then
         'MsgBox ""
         i6 = ws2.UsedRange.Rows(ws2.UsedRange.Rows.Count).Row + 1
         
         ws3.Range("A" & i5 & ":F" & i5).Copy ws2.Range("B" & i6)
         ws2.Range("A" & i6) = "TOBE_GR"
         
         Else
        ' MsgBox ""
         If ws3.Range("E" & i5) <> ws2.Range("F" & mokc_GRPO.Item(ws2.Name).Item("BODY").Item(str4).Item(1).Item(1).Key) Then
         i6 = ws2.UsedRange.Rows(ws2.UsedRange.Rows.Count).Row + 1
         ws3.Range("A" & i5 & ":F" & i5).Copy ws2.Range("B" & i6)
         ws2.Range("A" & i6) = "TOBE_GR"
         End If
         'MsgBox ""
         
         End If
     
     
     ElseIf date2 + 10 < date1 Then
     Exit For
     
     Else
   
     
     End If
     
     
     
     
     
     Next
     
     
     
     End If
     
     
     
     
     wb1.Close 0
     End If
     
     
     
     
End If




DoEvents

str1 = PF_Z(get_Single_File(mokc, mfso))
Loop

Application.DisplayAlerts = False
Workbooks(THIS_WB_NAME).Save

Application.ScreenUpdating = True
Application.StatusBar = THIS_WB_NAME & "SAVED!"


End Sub
Private Function wb_open_ONLY_READ(wb As Workbook, flfp As String, Optional b_n As Boolean = True) As Boolean
On Error GoTo Error1
'只读禁用宏方式打开工作簿
'如果外面有了还一样则不做任何操作，如果外面没有，复制到外面，如果外面有单是不一样，备份外面的，并复制

Dim str1 As String, str2 As String, str3 As String
str1 = P_SPLIT(flfp, "\", -1)
str2 = "Z:\24_Temp\PA_Logs\TOOLS\GoodsReceiptPurchaseOder\TEMP\"

'str2 = "Z:\24_Temp\PA_Logs\TOOLS\GoodsReceiptPurchaseOder\TEMP\" & Format(now(), "YYYY-MM-DD") & "\"
Dim mfso1 As New CFSO
If mfso1.FileExists(str2 & str1) Then
    If mfso1.Datelastmodify(str2 & str1) = mfso1.Datelastmodify(flfp) Then
    If b_n = False Then
    wb_open_ONLY_READ = False
    Exit Function
    End If
    Else
    str3 = str2 & Format(mfso1.Datelastmodify(str2 & str1), "YYYY-MM-DD") & "\"
    mfso1.CreateFolder str3
    mfso1.copy_file str2 & str1, str3 & str1
    mfso1.copy_file flfp, str2 & str1
    End If
Else
mfso1.copy_file flfp, str2 & str1
End If


'

mfso1.CreateFolder str2


wb_open_ONLY_READ = False

Dim i As Integer
Dim fln, flp As String
fln = Right(flfp, Len(flfp) - InStrRev(flfp, "\"))
flp = Left(flfp, Len(flfp) - Len(fln))
Dim temp_b As Boolean
temp_b = False
For i = 1 To Workbooks.Count
If Workbooks(i).Name = fln Then
temp_b = True
Set wb = Workbooks(i)
Exit For
End If
Next
If temp_b = False Then
If Dir(flp & fln) <> "" Then

On Error GoTo Error1:

Application.AutomationSecurity = msoAutomationSecurityForceDisable


'Set wb = Workbooks.Open(flp & fln, False, True)
Set wb = Workbooks.Open(str2 & str1, False, True)

Application.AutomationSecurity = msoAutomationSecurityLow


temp_b = True
End If
End If
wb_open_ONLY_READ = temp_b
Exit Function
Error1:
    MsgBox "open_wb function:" + Err.Description
    Err.Clear
    Exit Function
    

End Function


Private Function init_para() As Boolean
On Error GoTo Errhand
init_para = False
'1.初始化参数,建立应该有的工作表
Dim str1  As String, str2 As String, str3 As String, str4 As String
Dim para1 As String, para2 As String, para3 As String, para4 As String
Dim i1 As Integer, i2 As Integer, i3 As Integer, i4 As Integer
Dim wb As Workbook
Dim ws As Worksheet
For i1 = 1 To mokc_GRPO.Count
mokc_GRPO.Remove 1
Next
'============================================================================================
Set wb = ThisWorkbook
If InStr(wb.Name, Left(THIS_WB_NAME, Len(THIS_WB_NAME) - 5)) = 0 Then
init_para = False
MsgBox "workbook name must be:" & THIS_WB_NAME
Exit Function
End If
'============================================================================================
para1 = "receiving report_CN.xxxxx"
Set ws = get_ws(wb, para1)
ws.Range("A3") = "STATUS"
ws.Range("B3") = "FLFP"
ws.Range("C3") = "DATE_LAST_MODIFY"
ws.Range("D3") = "DATE"
If mokc_GRPO.Item(ws.Name) Is Nothing Then
mokc_read_ws mokc_GRPO, ws, 2, 2, 3
End If

'============================================================================================
para1 = "GRPO_DATA"
'MaterialID  ProjectNo   MaterialDesc    BINLocation DateofRecd  SecDate of Rec'd
Set ws = get_ws(wb, para1)
i1 = 1: i2 = 3
ws.Cells(i2, i1) = "STATUS": i1 = i1 + 1
ws.Cells(i2, i1) = "MaterialID": i1 = i1 + 1
ws.Cells(i2, i1) = "ProjectNo": i1 = i1 + 1
ws.Cells(i2, i1) = "MaterialDesc": i1 = i1 + 1
ws.Cells(i2, i1) = "BINLocation": i1 = i1 + 1
ws.Cells(i2, i1) = "DateofRecd": i1 = i1 + 1
ws.Cells(i2, i1) = "SecDate of Rec'd": i1 = i1 + 1
If mokc_GRPO.Item(ws.Name) Is Nothing Then
mokc_read_ws mokc_GRPO, ws, 2, 2, 3
End If
'============================================================================================
para1 = "PARA"
Set ws = get_ws(wb, para1)
ws.Range("A1") = "Z:\42_Logistics\01_Project Data"
add_comm "FDN_ROOT_GRPO", ws, 1, 1, False
init_para = True
Exit Function
Errhand:
MsgBox Err.Description
End Function

Private Function my_cdate(rg As Range) As Date
If rg = "Date of Rec'd" Then
my_cdate = CDate("2000-01-01")
Exit Function
ElseIf rg = "DateofRecd" Then
my_cdate = CDate("2000-01-01")
Exit Function
Else
  my_cdate = CDate(rg)
 'my_cdate = CDate("2000-01-01")

End If
End Function

Private Function my_sendkeys(Text As String) As Boolean
On Error GoTo ErrorHand
Dim MSForms_DataObject As Object
Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
MSForms_DataObject.SetText Text
MSForms_DataObject.PutInClipboard
Set MSForms_DataObject = Nothing
SendKeys "^v"
my_sendkeys = True
Exit Function
ErrorHand:
my_sendkeys = False
End Function


'VBA宏使用了绑定到文本复制到剪贴板。
'作者Justin Kay，8/15/2014
'  ""
 
 
'End Sub
Private Function delay(milliseconds As Long)
Dim delay_time As Double, start, now
start = Timer
delay_time = milliseconds / 1000
Do While now - start < delay_time
    now = Timer
    If now < start Then now = now + 86400
    DoEvents
Loop
End Function

Sub testxxx()
Dim majjl As New C_AJJL
delay 2000
Dim str1 As String
str1 = "Z:\24_Temp\PA_Logs\TOOLS\GoodsReceiptPurchaseOder\AI\Zhaoergang\CLOSE_DISPLAY_LOGS.bmp"

majjl.Add_Pic str1
majjl.L_CLICK_PIC "Display logs", str1, 10, 10


End Sub


Sub get_stock_data()
'1.访问服务器 Z:\24_Temp\PA_Logs\TOOLS\GoodsReceiptPurchaseOder\input\42_logistics
'2.获取其中全部的 receiving report*.xlsm
'3.将这些文件录入工作表  receiving report_CN.xxxxx 并比较其新旧
'4.如果新，或者是有更新，则将其只读 禁用宏的方式打开，并复制 开始日期（最近一周，或者三天）的 DATA也中收货记录 到 GRPO_DATA 页中。

Dim s_datelast As String

Dim str1  As String, str2 As String, str3 As String, str4 As String
Dim str5 As String, str6 As String
Dim c As Range

Dim para1 As String, para2 As String, para3 As String, para4 As String
Dim i1 As Long, i2 As Long, i3 As Long, i4 As Long, i5 As Long, i6 As Long
Dim i1_last As Long
Dim i2_last As Long
Dim i3_last As Long

Dim dele_row As Boolean



Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet
Dim wb1 As Workbook, wb2 As Workbook

Dim b_c As Boolean

If init_para = False Then Exit Sub
Set ws1 = get_ws(Workbooks(THIS_WB_NAME), "receiving report_CN.xxxxx")
Set ws2 = get_ws(Workbooks(THIS_WB_NAME), "STOCK_DATA")
Dim date1 As Date
Dim date2 As Date

'表头
ws2.Range("A3") = "SN": ws2.Range("B3") = "PO": ws2.Range("C3") = "ProjectNo": ws2.Range("D3") = "suppliers": ws2.Range("E3") = "POgroup": ws2.Range("F3") = "SupplierCode": ws2.Range("G3") = "Item"
ws2.Range("H3") = "MaterialID": ws2.Range("I3") = "MaterialDesc": ws2.Range("J3") = "RequirementNo": ws2.Range("K3") = "POQty": ws2.Range("L3") = "RecdQty": ws2.Range("M3") = "OpenQuant": ws2.Range("N3") = "UOM"
ws2.Range("O3") = "Binlocation": ws2.Range("P3") = "Duedate": ws2.Range("Q3") = "Entering": ws2.Range("R3") = "Tag"
ws2.Range("S3") = "FROM"
'表头




Dim mokc As New OneKeyCls
Dim mfso As New CFSO

'============================
 to_mokc mokc, "FD", "$", get_para_rg(Workbooks(THIS_WB_NAME).Worksheets("PARA").Range("A1:Z1"), "FDN_ROOT_GRPO", "N")
 to_mokc mokc, "FD_TPF", "$", "*"
 to_mokc mokc, "FD_DEL", "$", "OLD$00_Archive$Old$old"
 to_mokc mokc, "FL_TPF", "$", "Scan tool*.xlsm"





str1 = PF_Z(get_Single_File(mokc, mfso))
Do While Len(str1) > 0
b_c = False

If mokc_GRPO.Item(ws1.Name).Item("BODY").Item(str1) Is Nothing Then
'MsgBox ""
i2 = ws1.UsedRange.Rows(ws1.UsedRange.Rows.Count).Row + 1
ws1.Cells(i2, 2) = str1
ws1.Cells(i2, 4) = mfso.Datelastmodify(str1)
ws1.Cells(i2, 3) = mfso.Datelastmodify(str1)
'b_c = True


Else

str2 = mokc_GRPO.Item(ws1.Name).Item("BODY").Item(str1).Item(1).Item(1).Item("DATE_LAST_MODIFY").Key
str3 = mfso.Datelastmodify(str1)
i2 = mokc_GRPO.Item(ws1.Name).Item("BODY").Item(str1).Item(1).Item(1).Key

If str2 <> str3 Then
'b_c = True


ws1.Cells(i2, 4) = str3
ws1.Cells(i2, 3) = str3


End If


If CDate(ws1.Cells(i2, 3)) >= date1 Then
b_c = True
End If




End If




If b_c Then


Application.ScreenUpdating = True
Application.StatusBar = str1
s_datelast = mfso.get_flndatesize(str1)
DoEvents
Application.ScreenUpdating = False


     
     '只读禁宏方法打开电子表格
     only_new = False
     
     If wb_open_ONLY_READ(wb1, str1, only_new) Then
     
     'MsgBox ""
     
     If ws_exist(wb1, "DATA") = True Then
     Set ws3 = get_ws(wb1, "DATA")
     
     i3 = ws3.UsedRange.Rows(ws3.UsedRange.Rows.Count).Row
     'For i4 = 2 To i3
     'i5 = i3 - i4 + 1
     'str4 = ws3.Range("H" & i4)
     'str5 = Trim(ws3.Range("O" & i4))
     'str6 = CStr(ws3.Range("H" & i4).Interior.Color)
     '查找删除
     'With ws2.Columns(8)
    'Set c = .Find(str4, LookIn:=xlValues)
    'If Not c Is Nothing Then
    '    ws2.Rows(c.Row).Delete
    'End If
    ' End With
     '查找删除
     'If str6 = "16777215" And Len(str5) > 0 Then
     '白色,且有库位，添加
      'i6 = ws2.UsedRange.Rows(ws2.UsedRange.Rows.Count).Row + 1
      ' ws3.Range("A" & i4 & ":R" & i4).Copy ws2.Range("A" & i6)
    ' Else
     '其他
     'End If
     'Next
     
     '1.总表中删除 同名表单
     
     del_inws ws2, "S", "4", wb1.Name
     
     
    
  
            
       del_inws ws3, "O", "2", ""
       
    del_inws_rgb ws3, "O", "2", rgb(0, 255, 0)
    
      
        ws3.AutoFilterMode = False
         
         
         
         
         i3_last = ws3.UsedRange.Rows(ws3.UsedRange.Rows.Count).Row
         
         
         i2_last = ws2.UsedRange.Rows(ws2.UsedRange.Rows.Count).Row
         
         If i3_last > 2 Then
         
         ws3.Range("A2:R" & i3_last).Copy ws2.Range("A" & i2_last + 1)
         i3_last = ws2.UsedRange.Rows(ws2.UsedRange.Rows.Count).Row
         ws2.Range("S" & i2_last + 1 & ":S" & i3_last) = wb1.Name
         ws2.Range("T" & i2_last + 1 & ":T" & i3_last) = s_datelast
         Else
         
         End If
         
       

     '2.ws3中删除绿色，和无库位的行，余下的行复制入 总表末尾
     
     
     
     
     End If
     
     
     
     
     wb1.Close 0
     End If
     
     
     
     
End If




DoEvents

str1 = PF_Z(get_Single_File(mokc, mfso))
Loop

Application.DisplayAlerts = False
Workbooks(THIS_WB_NAME).Save

Application.ScreenUpdating = True
Application.StatusBar = THIS_WB_NAME & "SAVED!"


End Sub

Private Function del_inws(ws2 As Worksheet, s_col As String, s_row As String, s_v As String)
    Dim i2_last As Long
    Dim c As Range
    Dim i_start As Long
    
    i2_last = ws2.UsedRange.Rows(ws2.UsedRange.Rows.Count).Row
    If i2_last <= CLng(s_row) Then Exit Function
    
    '查询一下 s_v 没有则结束，有则记录为起始行
 
    With ws2.Columns(s_col)

    Set c = .Find(s_v, LookIn:=xlValues)
    If Not c Is Nothing Then

    i_start = c.Row

    Else

    Exit Function
    
    End If
    End With
    
    
    '查询一下 s_v 没有则结束，有则记录为起始行
    
    
    
    ws2.AutoFilterMode = False
    
    ws2.Range(s_col & s_row & ":" & s_col & i2_last).AutoFilter
    ws2.Range(s_col & s_row & ":" & s_col & i2_last).AutoFilter Field:=1, Criteria1:=s_v
    'ws2.Rows(s_row & ":" & i2_last).Delete Shift:=xlUp
     ws2.Rows(i_start & ":" & i2_last).Delete Shift:=xlUp
     ws2.AutoFilterMode = False
     
     DoEvents
     '部分xls删除一次并能删干净
     
       i2_last = ws2.UsedRange.Rows(ws2.UsedRange.Rows.Count).Row
    If i2_last <= CLng(s_row) Then Exit Function
    
    '查询一下 s_v 没有则结束，有则记录为起始行
 
    With ws2.Columns(s_col)

    Set c = .Find(s_v, LookIn:=xlValues)
    If Not c Is Nothing Then

    i_start = c.Row

    Else

    Exit Function
    
    End If
    End With
    
    
    '查询一下 s_v 没有则结束，有则记录为起始行
    
    
    
    ws2.AutoFilterMode = False
    
    ws2.Range(s_col & s_row & ":" & s_col & i2_last).AutoFilter
    ws2.Range(s_col & s_row & ":" & s_col & i2_last).AutoFilter Field:=1, Criteria1:=s_v
    'ws2.Rows(s_row & ":" & i2_last).Delete Shift:=xlUp
     ws2.Rows(i_start & ":" & i2_last).Delete Shift:=xlUp
     ws2.AutoFilterMode = False
     
End Function

Private Function del_inws_rgb(ws2 As Worksheet, s_col As String, s_row As String, rgb As Long)


    Dim i2_last As Long
    i2_last = ws2.UsedRange.Rows(ws2.UsedRange.Rows.Count).Row
    
    
    ws2.AutoFilterMode = False
    
   
    If i2_last < s_row Then Exit Function
     ws2.Range(s_col & s_row & ":" & s_col & i2_last).AutoFilter
    ws2.Range(s_col & s_row & ":" & s_col & i2_last).AutoFilter Field:=1, Criteria1:=rgb, Operator:=xlFilterCellColor
    ws2.Rows(s_row & ":" & i2_last).Delete Shift:=xlUp
     ws2.AutoFilterMode = False



        

End Function

Sub get_print_lable()
'1.访问服务器 Z:\24_Temp\PA_Logs\TOOLS\GoodsReceiptPurchaseOder\input\42_logistics
'2.获取其中全部的 receiving report*.xlsm
'3.将这些文件录入工作表  receiving report_CN.xxxxx 并比较其新旧
'4.如果新，或者是有更新，则将其只读 禁用宏的方式打开，并复制 开始日期（最近一周，或者三天）的 DATA也中收货记录 到 GRPO_DATA 页中。


Dim str1  As String, str2 As String, str3 As String, str4 As String
Dim str5 As String, str6 As String
Dim c As Range

Dim para1 As String, para2 As String, para3 As String, para4 As String
Dim i1 As Long, i2 As Long, i3 As Long, i4 As Long, i5 As Long, i6 As Long
Dim i1_last As Long
Dim i2_last As Long
Dim i3_last As Long
Dim i3_start As Long

Dim dele_row As Boolean



Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet
Dim wb1 As Workbook, wb2 As Workbook

Dim b_c As Boolean

If init_para = False Then Exit Sub
'Set ws1 = get_ws(Workbooks(THIS_WB_NAME), "receiving report_CN.xxxxx")
Set ws2 = get_ws(Workbooks(THIS_WB_NAME), "PRINT_LABLE")
Dim date1 As Date
Dim date2 As Date

date1 = CDate(InputBox("Input date start YYYY-MM-DD", "", Format(now(), "YYYY-MM-DD")))




'表头
'SN  MaterialID                          Type    PickingListNumber   Express tracking number Comments


ws2.Range("A3") = "SN": ws2.Range("B3") = "MaterialID": ws2.Range("C3") = "Descriptions": ws2.Range("D3") = "ProjectNo": ws2.Range("E3") = "POQty": ws2.Range("F3") = "DelQty": ws2.Range("G3") = "UOM"
ws2.Range("H3") = "UOM Location": ws2.Range("I3") = "CaseNo": ws2.Range("J3") = "StationNo": ws2.Range("K3") = "POQty": ws2.Range("L3") = "Type": ws2.Range("M3") = "PickingListNumber": ws2.Range("N3") = "Comments"
'ws2.Range("O3") = "Binlocation": ws2.Range("P3") = "Duedate": ws2.Range("Q3") = "Entering": ws2.Range("R3") = "Tag"
ws2.Range("S3") = "FROM"

'表头




Dim mokc As New OneKeyCls
Dim mfso As New CFSO

'============================
 to_mokc mokc, "FD", "$", get_para_rg(Workbooks(THIS_WB_NAME).Worksheets("PARA").Range("A1:Z1"), "FDN_ROOT_GRPO", "N")
 to_mokc mokc, "FD_TPF", "$", "*"
 to_mokc mokc, "FD_DEL", "$", "OLD$00_Archive$Old$old"
 to_mokc mokc, "FL_TPF", "$", "Scan tool*.xlsm"





str1 = PF_Z(get_Single_File(mokc, mfso))
Do While Len(str1) > 0



If mfso.Datelastmodify(str1) > date1 Then
b_c = True
Else
b_c = False
End If


If b_c Then


Application.ScreenUpdating = True
Application.StatusBar = str1
DoEvents
Application.ScreenUpdating = False


     
     '只读禁宏方法打开电子表格
     only_new = True
     
     If wb_open_ONLY_READ(wb1, str1, only_new) Then
     
     'MsgBox ""
     
     If ws_exist(wb1, "print record") = True Then
     Set ws3 = get_ws(wb1, "print record")
     
     i3 = ws3.UsedRange.Rows(ws3.UsedRange.Rows.Count).Row
     'For i4 = 2 To i3
     'i5 = i3 - i4 + 1
     'str4 = ws3.Range("H" & i4)
     'str5 = Trim(ws3.Range("O" & i4))
     'str6 = CStr(ws3.Range("H" & i4).Interior.Color)
     '查找删除
     'With ws2.Columns(8)
    'Set c = .Find(str4, LookIn:=xlValues)
    'If Not c Is Nothing Then
    '    ws2.Rows(c.Row).Delete
    'End If
    ' End With
     '查找删除
     'If str6 = "16777215" And Len(str5) > 0 Then
     '白色,且有库位，添加
      'i6 = ws2.UsedRange.Rows(ws2.UsedRange.Rows.Count).Row + 1
      ' ws3.Range("A" & i4 & ":R" & i4).Copy ws2.Range("A" & i6)
    ' Else
     '其他
     'End If
     'Next
     
     '1.总表中删除 同名表单
     
     del_inws ws2, "S", "4", wb1.Name
     
     
    
     
      ' del_inws ws3, "O", "2", ""
       
    'del_inws_rgb ws3, "O", "2", rgb(0, 255, 0)
    
      
        ws3.AutoFilterMode = False
         
         
         
         i3_last = ws3.UsedRange.Rows(ws3.UsedRange.Rows.Count).Row
         i2_last = ws2.UsedRange.Rows(ws2.UsedRange.Rows.Count).Row
         
         i3_start = i3_last
         
         Do While ws3.Range("K" & i3_start) > date1
         i3_start = i3_start - 1
         If i3_start = 2 Then
         Exit Do
         End If
         Loop
         
         If i3_last = i3_start Then
         i3_last = 2
         End If
         
         i3_start = i3_start + 1
         
         
       
         
         Do While Len(Trim(ws3.Range("B" & i3_last))) = 0
         i3_last = i3_last - 1
         If i3_last = 4 Then
         Exit Do
         End If
         Loop
         
         
         
         If i3_last > 2 Then
         
         ws3.Range("A" & i3_start & ":R" & i3_last).Copy ws2.Range("A" & i2_last + 1)
         i3_last = ws2.UsedRange.Rows(ws2.UsedRange.Rows.Count).Row
         ws2.Range("S" & i2_last + 1 & ":S" & i3_last) = wb1.Name
         
         Else
         
         End If
         
       

     '2.ws3中删除绿色，和无库位的行，余下的行复制入 总表末尾
     
     
     
     
     End If
     
     
     
     
     wb1.Close 0
     End If
     
     
     
     
End If




DoEvents

str1 = PF_Z(get_Single_File(mokc, mfso))
Loop

Application.DisplayAlerts = False
Workbooks(THIS_WB_NAME).Save

Application.ScreenUpdating = True
Application.StatusBar = THIS_WB_NAME & "SAVED!"


End Sub
