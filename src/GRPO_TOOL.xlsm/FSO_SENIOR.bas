Attribute VB_Name = "FSO_SENIOR"
Option Explicit

'使用了 CFSO,OneKeyCls

Function Record_file_in_folder(mokc As OneKeyCls, Optional FDN_ROOT As String = "", Optional FLN_include As String = "") As String
Dim b_continue As Boolean
b_continue = True
Dim rec_b As Boolean

Dim i As Long

'mokc.item("PARAP").item("FDN_root")
'mokc.item("PARAP").item("FDN_current")
'mokc.item("PARA").item("FLN_include")

'合法性检测

If mokc.Item("PARA") Is Nothing And FDN_ROOT <> "" And FLN_include <> "" Then
mokc.Add "PARA", "PARA"
mokc.Item("PARA").Add FDN_ROOT, "FDN_root"
mokc.Item("PARA").Add FDN_ROOT, "FDN_current"
mokc.Item("PARA").Add "FLN_include", "FLN_include"
mokc.Item("PARA").Item("FLN_include").Add FLN_include
End If


If mokc.Item("PARA").Item("FDN_root") Is Nothing Then b_continue = False
If mokc.Item("PARA").Item("FDN_current") Is Nothing Then b_continue = False
If mokc.Item("PARA").Item("FLN_include") Is Nothing Then


mokc.Item("PARA").Add "FLN_include", "FLN_include"
mokc.Item("PARA").Item("FLN_include").Add ".tif"
mokc.Item("PARA").Item("FLN_include").Add ".xls"

End If

If b_continue = False Then
Record_file_in_folder = "ERROR:Record_file_in_folder"
MsgBox Record_file_in_folder
Exit Function
End If
'合法性检测


If mokc.Item("FILE") Is Nothing Then mokc.Add "FILE", "FILE"
FDN_ROOT = mokc.Item("PARA").Item("FDN_root").Key



Dim Fso As Object
Set Fso = CreateObject("Scripting.FileSystemObject")
Dim fd As Object
Set fd = Fso.getfolder(mokc.Item("PARA").Item("FDN_current").Key)
Dim fl As Object
Dim sfd As Object
Dim cur_i As Integer
cur_i = 1
Dim temp_s As String
Dim i_curr As Long



For Each fl In fd.Files


rec_b = True
temp_s = fl.path
If Len(temp_s) > Len(FDN_ROOT) Then
temp_s = Right(temp_s, Len(temp_s) - Len(FDN_ROOT))
End If

If rec_b Then

For i = 1 To mokc.Item("PARA").Item("FLN_include").Count
rec_b = False

If InStr(temp_s, mokc.Item("PARA").Item("FLN_include").Item(i).Key) > 0 Then
rec_b = True
Exit For
End If
Next
End If




If rec_b Then
i_curr = mokc.Item("FILE").Count
i_curr = i_curr + 1
mokc.Item("FILE").Add CStr(i_curr), CStr(i_curr)

mokc.Item("FILE").Item(CStr(i_curr)).Add fl.Name, "FLN"
mokc.Item("FILE").Item(CStr(i_curr)).Add CStr(fl.Size), "SIZE"
mokc.Item("FILE").Item(CStr(i_curr)).Add mokc.Item("PARA").Item("FDN_current").Key, "FDN"
mokc.Item("FILE").Item(CStr(i_curr)).Add Format(fl.DateLastModified, "YYYY-MM-DD HH:MM:SS"), "DATE"

End If



Next fl


'是否迭代
If mokc.Item("PARA").Item("ONLY_ROOT") Is Nothing Then
For Each sfd In fd.SubFolders
mokc.Item("PARA").Item("FDN_current").Key = sfd.path
Record_file_in_folder mokc
Next sfd
End If



End Function


Function mokc_read_ws(mokc As OneKeyCls, ws As Worksheet, Optional key_i1 As Integer = 1, Optional key_i2 As Integer = 0, Optional star_row As Integer = 1)
'将电子表格中的内容读入mokc，电子表格的 名称和，A1，B1...单元格的内容是关键字
If key_i2 = 0 Then key_i2 = key_i1


Dim wsn As String
wsn = ws.Name
If Not (mokc.Item(wsn) Is Nothing) Then
mokc.Remove wsn
End If
mokc.Add wsn, wsn
Dim i As Long
Dim i_last As Long
Dim j As Long

i_last = ws.UsedRange.Columns.Count
Dim temp_s1 As String
Dim temp_s2 As String
Dim temp_s3 As String
Dim temp_s4 As String
Dim temp_s5 As String

mokc.Item(wsn).Add "HEAD", "HEAD"
mokc.Item(wsn).Add "BODY", "BODY"
mokc.Item(wsn).Add "KEY", "KEY"
mokc.Item(wsn).Item("KEY").Add CStr(key_i1), "KEY1"
mokc.Item(wsn).Item("KEY").Add CStr(key_i2), "KEY2"



For i = 1 To i_last
temp_s1 = Trim(ws.Cells(star_row, i))
If Len(temp_s1) > 0 Then
If mokc.Item(wsn).Item("HEAD").Item(temp_s1) Is Nothing Then
mokc.Item(wsn).Item("HEAD").Add temp_s1, temp_s1
mokc.Item(wsn).Item("HEAD").Item(temp_s1).Add CStr(i), CStr(i)
End If
End If
Next

i_last = ws.UsedRange.Rows.Count

For i = star_row + 1 To i_last

    temp_s1 = Trim(ws.Cells(i, key_i1))
    temp_s2 = Trim(ws.Cells(i, key_i2))
    If Len(temp_s2) = 0 Then temp_s2 = temp_s1
    
    If Len(temp_s1) > 0 And Len(temp_s2) > 0 Then
    If mokc.Item(wsn).Item("BODY").Item(temp_s1) Is Nothing Then mokc.Item(wsn).Item("BODY").Add temp_s1, temp_s1
    If mokc.Item(wsn).Item("BODY").Item(temp_s1).Item(temp_s2) Is Nothing Then mokc.Item(wsn).Item("BODY").Item(temp_s1).Add temp_s2, temp_s2
    temp_s3 = CStr(i)
    If mokc.Item(wsn).Item("BODY").Item(temp_s1).Item(temp_s2).Item(temp_s3) Is Nothing Then mokc.Item(wsn).Item("BODY").Item(temp_s1).Item(temp_s2).Add temp_s3, temp_s3
    For j = 1 To mokc.Item(wsn).Item("HEAD").Count
    temp_s4 = mokc.Item(wsn).Item("HEAD").Item(j).Key
    temp_s5 = ws.Cells(i, CInt(mokc.Item(wsn).Item("HEAD").Item(j).Item(1).Key))
    mokc.Item(wsn).Item("BODY").Item(temp_s1).Item(temp_s2).Item(temp_s3).Add temp_s5, temp_s4
    Next
    




End If





Next






End Function


Function mokc_get_wsQ(mokc As OneKeyCls, ws As Worksheet, key1 As String) As String
Dim wsn As String

mokc_get_wsQ = ""
If ws.Name <> mokc.Item(1).Key Then
mokc_get_wsQ = "ERROR"
MsgBox "ERROR wsn" & mokc.Item(1).Key & "  " & ws.Name
Exit Function
Else
wsn = ws.Name
End If




'如果判断 有误则终止
Dim str1 As String, str2 As String
Dim firstAddress As String
     Dim temp_s1 As String, temp_s2 As String, temp_s3 As String, temp_s4 As String, temp_s5 As String
     Dim l_crow As Long
 Dim i As Long
 Dim key_i1 As Integer
 
 Dim j As Long
 
'  Dim wsn As String
  
Dim key_i2 As Integer

 
      


Dim mokc_temp As New OneKeyCls

Dim c As Range
Dim i_last As Long
'i_last = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).row

With ws.Columns(CInt(mokc.Item(1).Item("HEAD").Item("KEY").Item(1).Key))

    Set c = .Find(key1, LookIn:=xlValues)
    If Not c Is Nothing Then
        firstAddress = c.Address
        Do
            'c.Value = 5
            
         
            i = c.Row
          
          If c.Value <> key1 Then
    
          
          Else
          
          
          
            
              key_i1 = CInt(mokc.Item(1).Item("HEAD").Item("KEY").Item(1).Key)
              If mokc.Item(1).Item("HEAD").Item("KEY").Item(2) Is Nothing Then
              key_i2 = key_i1
              Else
              key_i2 = CInt(mokc.Item(1).Item("HEAD").Item("KEY").Item(2).Key)
              End If
              
          
          
              temp_s1 = Trim(ws.Cells(i, key_i1))
                  
    temp_s2 = Trim(ws.Cells(i, key_i2))
    If Len(temp_s2) = 0 Then temp_s2 = temp_s1
    
    If Len(temp_s1) > 0 And Len(temp_s2) > 0 Then
    If mokc.Item(wsn).Item("BODY").Item(temp_s1) Is Nothing Then mokc.Item(wsn).Item("BODY").Add temp_s1, temp_s1
    If mokc.Item(wsn).Item("BODY").Item(temp_s1).Item(temp_s2) Is Nothing Then mokc.Item(wsn).Item("BODY").Item(temp_s1).Add temp_s2, temp_s2
    temp_s3 = CStr(i)
    If mokc.Item(wsn).Item("BODY").Item(temp_s1).Item(temp_s2).Item(temp_s3) Is Nothing Then mokc.Item(wsn).Item("BODY").Item(temp_s1).Item(temp_s2).Add temp_s3, temp_s3
    For j = 1 To mokc.Item(wsn).Item("HEAD").Count
    temp_s4 = mokc.Item(wsn).Item("HEAD").Item(j).Key
    temp_s5 = ws.Cells(i, CInt(mokc.Item(wsn).Item("HEAD").Item(j).Item(1).Key))
    mokc.Item(wsn).Item("BODY").Item(temp_s1).Item(temp_s2).Item(temp_s3).Add temp_s5, temp_s4
    Next
    
    End If
    
    
          
          End If
          
       
       
       
            
            
            
            Set c = .FindNext(c)
            If i >= c.Row Then
            Exit Do
            End If
            
        Loop While Not c Is Nothing
    End If
End With




 
End Function


Function mokc_read_wsQ(mokc As OneKeyCls, ws As Worksheet, Optional key_i1 As Integer = 1, Optional key_i2 As Integer = 0, Optional star_row As Integer = 1)
'将电子表格中的内容读入mokc，电子表格的 名称和，A1，B1...单元格的内容是关键字
'具体内容先不读取，只是建立框架

If key_i2 = 0 Then key_i2 = key_i1


Dim wsn As String
wsn = ws.Name
If Not (mokc.Item(wsn) Is Nothing) Then
mokc.Remove wsn
End If
mokc.Add wsn, wsn
Dim i As Long
Dim i_last As Long
Dim j As Long

i_last = ws.UsedRange.Columns.Count
Dim temp_s1 As String
Dim temp_s2 As String
Dim temp_s3 As String
Dim temp_s4 As String
Dim temp_s5 As String

mokc.Item(wsn).Add "HEAD", "HEAD"
mokc.Item(wsn).Add "BODY", "BODY"
mokc.Item(wsn).Add "KEY", "KEY"
mokc.Item(wsn).Item("KEY").Add CStr(key_i1), "KEY1"
mokc.Item(wsn).Item("KEY").Add CStr(key_i2), "KEY2"



For i = 1 To i_last
temp_s1 = Trim(ws.Cells(star_row, i))
If Len(temp_s1) > 0 Then
If mokc.Item(wsn).Item("HEAD").Item(temp_s1) Is Nothing Then
mokc.Item(wsn).Item("HEAD").Add temp_s1, temp_s1
mokc.Item(wsn).Item("HEAD").Item(temp_s1).Add CStr(i), CStr(i)
End If
End If
Next

'i_last = ws.UsedRange.Rows.Count

'For i = star_row + 1 To i_last

 '   temp_s1 = Trim(ws.Cells(i, key_i1))
 '   temp_s2 = Trim(ws.Cells(i, key_i2))
 '   If Len(temp_s2) = 0 Then temp_s2 = temp_s1
    
 '   If Len(temp_s1) > 0 And Len(temp_s2) > 0 Then
 '   If mokc.Item(wsn).Item("BODY").Item(temp_s1) Is Nothing Then mokc.Item(wsn).Item("BODY").Add temp_s1, temp_s1
 '   If mokc.Item(wsn).Item("BODY").Item(temp_s1).Item(temp_s2) Is Nothing Then mokc.Item(wsn).Item("BODY").Item(temp_s1).Add temp_s2, temp_s2
 '   temp_s3 = CStr(i)
 '   If mokc.Item(wsn).Item("BODY").Item(temp_s1).Item(temp_s2).Item(temp_s3) Is Nothing Then mokc.Item(wsn).Item("BODY").Item(temp_s1).Item(temp_s2).Add temp_s3, temp_s3
    'For j = 1 To mokc.Item(wsn).Item("HEAD").Count
    'temp_s4 = mokc.Item(wsn).Item("HEAD").Item(j).key
    'temp_s5 = ws.Cells(i, CInt(mokc.Item(wsn).Item("HEAD").Item(j).Item(1).key))
    'mokc.Item(wsn).Item("BODY").Item(temp_s1).Item(temp_s2).Item(temp_s3).Add temp_s5, temp_s4
    'Next
  '  End If
'Next






End Function


Function mokc_read_ws_A(mokc As OneKeyCls, ws As Worksheet, Optional i As Long = 0) As Boolean
'读取工作表中的一行到 mokc中
Dim temp_s1 As String, temp_s2 As String, temp_s3 As String, temp_s4 As String, temp_s5 As String

Dim key_i1 As Integer, key_i2 As Integer

Dim wsn As String
wsn = ws.Name
Dim j As Integer


If i = 0 Then i = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row
If mokc.Item(ws.Name).Item("BODY") Is Nothing Or mokc.Item(ws.Name).Item("HEAD") Is Nothing Then
mokc_read_ws_A = False
MsgBox "Error! mokc_read_ws_A " & Chr(10) & ws.Name
End If

    
    temp_s1 = Trim(ws.Cells(i, CInt(mokc.Item(ws.Name).Item("KEY").Item("KEY1").Key)))
    temp_s2 = Trim(ws.Cells(i, CInt(mokc.Item(ws.Name).Item("KEY").Item("KEY2").Key)))
    If Len(temp_s2) = 0 Then temp_s2 = temp_s1
    
    If Len(temp_s1) > 0 And Len(temp_s2) > 0 Then
    If mokc.Item(wsn).Item("BODY").Item(temp_s1) Is Nothing Then mokc.Item(wsn).Item("BODY").Add temp_s1, temp_s1
    If mokc.Item(wsn).Item("BODY").Item(temp_s1).Item(temp_s2) Is Nothing Then mokc.Item(wsn).Item("BODY").Item(temp_s1).Add temp_s2, temp_s2
    temp_s3 = CStr(i)
    If mokc.Item(wsn).Item("BODY").Item(temp_s1).Item(temp_s2).Item(temp_s3) Is Nothing Then mokc.Item(wsn).Item("BODY").Item(temp_s1).Item(temp_s2).Add temp_s3, temp_s3
    For j = 1 To mokc.Item(wsn).Item("HEAD").Count
    temp_s4 = mokc.Item(wsn).Item("HEAD").Item(j).Key
    temp_s5 = ws.Cells(i, CInt(mokc.Item(wsn).Item("HEAD").Item(j).Item(1).Key))
    mokc.Item(wsn).Item("BODY").Item(temp_s1).Item(temp_s2).Item(temp_s3).Add temp_s5, temp_s4
    Next
    End If
    


    
End Function


Private Function init_1(mokc As OneKeyCls) As Boolean
Dim str1 As String
If mokc.Item("FD") Is Nothing Then
str1 = InputBox("input folder path", "GET_SINGLE_FILE", "")
to_mokc mokc, "FD", "$", str1

If mokc.Item("FL_TPF") Is Nothing Then

str1 = InputBox("input FL_TPF", "GET_SINGLE_FILE", "*.doc")
to_mokc mokc, "FL_TPF", "$", str1

End If

If mokc.Item("FD_TPF") Is Nothing Then

str1 = InputBox("input FL_TPF", "GET_SINGLE_FILE", "*")
to_mokc mokc, "FD_TPF", "$", str1

End If

End If
End Function


Function get_Single_File(mokc As OneKeyCls, mfso As CFSO) As String

If mokc.Item("FD") Is Nothing Then
init_1 mokc
End If




'逐个的弹出符合要求的文件
Dim fd1 As String, fd2 As String

Dim v_fd1 As Variant
Dim v_fl1 As Variant
Dim str1 As String, str2 As String


Dim i As Integer
Dim j As Integer
Dim b_c As Boolean

If mokc.Item("FD_DEL") Is Nothing Then
mokc.Add "FD_DEL", "FD_DEL"
End If

If mokc.Item("FD_TPF") Is Nothing Then
mokc.Add "FD_TPF", "FD_TPF"
End If

If mokc.Item("FL") Is Nothing Then
mokc.Add "FL", "FL"
End If



If mokc.Item("FL").Count > 0 Then
get_Single_File = mokc.Item("FL").Item(1).Key
mokc.Item("FL").Remove 1
Exit Function
End If

If mokc.Item("FD").Count = 0 Then
get_Single_File = ""
Exit Function
End If
'把 第一个FD拿出来，找子 FD，判断是否合格，合格则加入"FD",不合格则扔掉，找文件 ，合格则 用，不合格则扔掉，删除这个“FD”


fd1 = mokc.Item("FD").Item(1).Key
mokc.Item("FD").Remove 1
v_fd1 = mfso.GetFiles(fd1, False, "fo", False)
For j = LBound(v_fd1) To UBound(v_fd1)
fd2 = v_fd1(j)
If mokc.Item("FD_DEL").Item(fd2) Is Nothing Then
b_c = False
For i = 1 To mokc.Item("FD_TPF").Count
If fd2 Like mokc.Item("FD_TPF").Item(i).Key Then
b_c = True
Exit For
End If
Next
If b_c = True Then
mokc.Item("FD").Add fd1 & "\" & fd2, fd1 & "\" & fd2
End If
End If
Next
v_fl1 = mfso.GetFiles(fd1, False, "f", False)
For j = LBound(v_fl1) To UBound(v_fl1)
str1 = v_fl1(j)
b_c = False
For i = 1 To mokc.Item("FL_TPF").Count
If str1 Like mokc.Item("FL_TPF").Item(i).Key Then
b_c = True
Exit For
End If
Next
If b_c = True Then
mokc.Item("FL").Add fd1 & "\" & str1, fd1 & "\" & str1
End If
Next



get_Single_File = get_Single_File(mokc, mfso)


End Function

Sub macrotess()
Dim mfso As New CFSO
Dim mokc As New OneKeyCls
Dim str1 As String
'添加根目录
 to_mokc mokc, "FD", "$", "Z:\31_PTS\01_Projects\ASY\CN.305899_.BBAC_M254 Engine Assy Line\08_Engineering\088_Concept_Design\09_Station_Data$Z:\31_PTS\01_Projects\ASY\CN.305899_.BBAC_M254 Engine Assy Line\08_Engineering\088_Concept_Design\01_Common_Data\03_Standard units"
 to_mokc mokc, "FD_TPF", "$", "*?.?????.???.??.??*$TIF$tif"
 to_mokc mokc, "FD_DEL", "$", "2D_CATIA$3D_CATIA$3DXML$FEM$MDM$PPT$3D"
 to_mokc mokc, "FL_TPF", "$", "?.?????.???.??.??*.xls*"
 
str1 = get_file_infolder(mokc, mfso)
Do While Len(str1) > 0
MsgBox str1
DoEvents
str1 = get_file_infolder(mokc, mfso)
Loop
End Sub

 Function to_mokc(mokc As OneKeyCls, s_item As String, ssep As String, str1 As String)
If mokc.Item(s_item) Is Nothing Then mokc.Add s_item, s_item
Dim b_c As Boolean
Dim i As Integer
Dim skey As String
If Right(str1, 1) <> ssep Then str1 = str1 & ssep
Do While InStr(str1, ssep) > 0
skey = Left(str1, InStr(str1, ssep))
str1 = Right(str1, Len(str1) - Len(skey))
skey = Replace(skey, ssep, "")
If Len(skey) > 0 Then
b_c = True
For i = 1 To mokc.Item(s_item).Count
If mokc.Item(s_item).Item(i).Key = skey Then
b_c = False
Exit For
End If
Next
If b_c Then
If mokc.Item(s_item).Item(skey) Is Nothing Then
mokc.Item(s_item).Add skey, skey
Else
mokc.Item(s_item).Add skey
End If
End If
End If
Loop
End Function

