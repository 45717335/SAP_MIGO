VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_AJJL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Const MOUSEEVENTF_MOVE = &H1 '  mouse move
Private Const MOUSEEVENTF_LEFTDOWN = &H2 '  left button down
Private Const MOUSEEVENTF_LEFTUP = &H4 '  left button up
Private Const MOUSEEVENTF_RIGHTDOWN = &H8 '  right button down
Private Const MOUSEEVENTF_RIGHTUP = &H10 '  right button up
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20 '  middle button down
Private Const MOUSEEVENTF_MIDDLEUP = &H40 '  middle button up
Private Const MOUSEEVENTF_ABSOLUTE = &H8000 '  absolute move


Private Declare Function ShowWindow Lib "user32" (ByVal hwnd _
As Long, ByVal nCmdShow As Long) As Long



Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long



Private Type PicBmp
Size As Long
Type As Long
hBmp As Long
hPal As Long
Reserved As Long
End Type


Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'获取光标形状（用光标句柄标志光标形状）

Private Declare Function GetCursorInfo Lib "user32.dll" (ByRef pci As CURSORINFO) As Boolean


Private Type POINT
    x As Long
    y As Long
End Type

Private Type CURSORINFO
    cbSize As Long
    flags As Long
    hCursor As Long
    ptScreenPos As POINT
End Type



'获取光标形状（用光标句柄标志光标形状）


'光标特征值

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, icoinfo As ICONINFO) As Long
'Private Declare Function GetCursorInfo Lib "user32.dll" (ByRef pci As CURSORINFO) As Boolean
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long

'Private Type POINT
'    x As Long
'    y As Long
'End Type

'Private Type CURSORINFO
'    cbSize As Long
'    flags As Long
'    hCursor As Long
'    ptScreenPos As POINT
'End Type

Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hBMMask As Long
    hBMColor As Long
End Type

'光标特征值






'获取光标形状（用光标句柄标志光标形状）


'光标特征值

'Private Type POINT
'    x As Long
'    y As Long
'End Type

'Private Type CURSORINFO
'    cbSize As Long
'    flags As Long
'    hCursor As Long
'    ptScreenPos As POINT
'End Type



'光标特征值






Private Const SW_SHOW = &H5

'20190107
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hwnd1 As Long, ByVal hwnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'20190107


Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Enum JpMode
theScreen = 0 '全屏截图
theform = 1 '当前焦点窗口截图
End Enum
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Const CF_BITMAP = 2


Private Type GUID
data1 As Long
data2 As Integer
data3 As Integer
data4(0 To 7) As Byte
End Type







'Private Const S_OK As Long = &H0
'Private Declare Function CoCreateInstance Lib "ole32.dll" (clsid As GUID, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, riid As GUID, pvObj As Any) As Long



'Private Const CLSCTX_INPROC_SERVER As Long = &H1
'Private Declare Function IIDFromString Lib "ole32.dll" (ByVal lpszIID As Long, iid As GUID) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszCLSID As Long, clsid As GUID) As Long

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long '获取句柄
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long '释放DC
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long





Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source




'颜色表
Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbAlpha As Byte   '透明通道
End Type
Private Type BITMAPINFOHEADER
    biSize As Long          '位图大小
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer   '信息头长度
    biCompression As Long   '压缩方式
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type
'图片文件头
Private Const DIB_RGB_COLORS = 0 '  color table in RGBs
'Dim BI As BITMAPINFO
'Dim BI1 As BITMAPINFO
'Dim PP As New Form1
Private Const HIMETRIC_PER_PIXEL = 96 / 2540
Private Const vbPicTypeBitmap = 1
Private Enum EnumPicMode
    BlackWhite = 0
    GrayScale = 1
End Enum


'==================================================================================================================================================================================

Public mCol As New Collection
'Private mokc As New OneKeyCls


Function L_CLICK_WIN(Win_name As String, i_x As Integer, i_y As Integer) As String

Dim hwnd As Long
hwnd = my_findwindow(Win_name)
If hwnd = 0 Then
    
    Exit Function
End If


ShowWindow hwnd, SW_SHOW
delay 200
ShowWindow hwnd, SW_SHOW
delay 200
                     
    'AppActivate win_tit
    '清空PDM查询框
    'AppActivate Win_tit
    'showwindow 代替AppActivate,因为后者在处理IE多窗口时候报错5
    Dim L As Long
    Dim R As Long
    Dim T As Long
    Dim b As Long
    '为什么重复此段代码
    hwnd = my_findwindow(Win_name)
    ' hwnd = FindWindow(vbNullString, win_tit)
   ' If hwnd = 0 Then
   ' If hwnd = 0 Then
   ' 'MEMO_Win_Pic_Leftclick = False
   ' 'Exit Function
   ' hwnd = my_findwindow(win_tit)
   ' If hwnd = 0 Then
   ' MEMO_Win_Pic_Leftclick = False
    'Exit Function
    ''End If
    'End If
   ' Else
  '  End If
    '为什么重复此段代码
    '因为ShowWindow 和 AppActivate 引起了 hwnd 的变化，原因暂未知(20150525xuefeng.gao@thyssenkrupp.com)
    Get_WR hwnd, L, T, R, b
 
    Dim s_x As Integer
    Dim s_y As Integer
    s_x = i_x
    s_y = i_y
    SetCursorPos L + s_x, T + s_y
    mouse_leftclick L + s_x, T + s_y
   

End Function



Function L_CLICK_PIC(Win_name As String, pic_flfp As String, i_x As Integer, i_y As Integer) As Boolean

Add_Pic pic_flfp
L_CLICK_PIC = MEMO_LClick_PIC(Win_name, mCol.Item(P_SPLIT(pic_flfp, "\", -1)), i_x, i_y)
 
 
End Function


Public Function Add_Pic(flfp As String)
Dim m_StdPicture As StdPicture
Dim fdn As String, fln As String
fln = P_SPLIT(flfp, "\", -1)
fdn = Left(flfp, Len(flfp) - Len(flfp))
If haveItem(mCol, fln) Is Nothing Then
Set m_StdPicture = LoadPicture(flfp)
mCol.Add m_StdPicture, fln
End If
End Function



Private Function haveItem(mCol As Collection, vntIndexKey As Variant) As StdPicture
    Err.Clear
    On Error GoTo GetItemError
    '引用集合中的一个元素时使用。
    'vntIndexKey 包含集合的索引或关键字，
    '这是为什么要声明为 Variant 的原因
    '语法：Set foo = x.Item(xyz) or Set foo = x.Item(5)
    Set haveItem = mCol(vntIndexKey)
    Exit Function
    
GetItemError:
    'Debug.Print Err.Source & " : Error Was found in Item(vntIndexKey as Variant) as oneKeyCls . " & vbCrLf & Err.Description & " : " & Err.Number
    Set haveItem = Nothing
End Function


Private Function P_SPLIT(ByVal txtRange, ByVal splitter As String, Optional ByVal get_index As Integer = 0)
'拆分字符串，
'get_index=0    则返回被拆分出来的第一个字符串
'get_index=1，2，3    则返回被拆分出来的第二、三、四个字符传
'get_index=-1，-2，-3    则返回被拆分出来的倒数第一、二、三个字符传
'越界返回 ""
    Dim a As Variant
    Dim b As Variant
    Dim c As Variant
    a = Split(txtRange, splitter)
    b = LBound(a)
    c = UBound(a)
    If get_index = -1 Then
    P_SPLIT = a(c)
    ElseIf get_index = 0 Then
    P_SPLIT = a(b)
    ElseIf get_index >= b And get_index <= c Then
    P_SPLIT = a(get_index)
    ElseIf get_index <= -1 And get_index >= -1 * c - 1 Then
    P_SPLIT = a(c + 1 + get_index)
    Else
    P_SPLIT = ""
    End If
End Function


Private Function MEMO_LClick_PIC(win_tit As String, Pic_s As StdPicture, x As Integer, y As Integer) As Boolean
'函数功能
'点击 窗口名称为Win_name,的窗口中，的一张图片
'这张图片 格式为stdPicture

'111111111111111111
'my_findwindow
'222222222222222222
'ShowWindow
'333333333333333333
'MEMO_Check_Win_Open
'444444444444444444
'SetCursorPos
'5555555555555555555
'mouse_leftclick


If Pic_s Is Nothing Then

MsgBox "MEMO_LClick_PIC" & "Can not find pic"
MEMO_LClick_PIC = False
Exit Function
End If


    
Dim hwnd As Long
hwnd = my_findwindow(win_tit)
If hwnd = 0 Then
    MEMO_LClick_PIC = False
    Exit Function
End If


ShowWindow hwnd, SW_SHOW
delay 200
ShowWindow hwnd, SW_SHOW
delay 200
                     
    'AppActivate win_tit
    '清空PDM查询框
    'AppActivate Win_tit
    'showwindow 代替AppActivate,因为后者在处理IE多窗口时候报错5
    Dim L As Long
    Dim R As Long
    Dim T As Long
    Dim b As Long
    '为什么重复此段代码
    hwnd = my_findwindow(win_tit)
    ' hwnd = FindWindow(vbNullString, win_tit)
   ' If hwnd = 0 Then
   ' If hwnd = 0 Then
   ' 'MEMO_Win_Pic_Leftclick = False
   ' 'Exit Function
   ' hwnd = my_findwindow(win_tit)
   ' If hwnd = 0 Then
   ' MEMO_Win_Pic_Leftclick = False
    'Exit Function
    ''End If
    'End If
   ' Else
  '  End If
    '为什么重复此段代码
    '因为ShowWindow 和 AppActivate 引起了 hwnd 的变化，原因暂未知(20150525xuefeng.gao@thyssenkrupp.com)
    Get_WR hwnd, L, T, R, b
 
    Dim s_x As Integer
    Dim s_y As Integer
    
    If MEMO_Check_Win_Open(win_tit, Pic_s, s_x, s_y) Then
    SetCursorPos L + s_x + x, T + s_y + y
    mouse_leftclick L + s_x + x, T + s_y + y
    'delay 500
    Else
    MEMO_LClick_PIC = False
    Exit Function
    End If
    MEMO_LClick_PIC = True
    
MEMO_LClick_PIC = True
End Function

'1111111111111111111111111
 Function my_findwindow(ByRef winname As String) As Long
Dim str2 As String



'第一次尝试： 使用最传统的函数，返回窗口句柄，如果失败，则找类似的进行匹配
my_findwindow = FindWindow(vbNullString, winname)
If my_findwindow > 0 Then Exit Function
'使用最传统的函数，返回窗口句柄，如果失败，则找类似的进行匹配



'第二次尝试： 相似或者包含
'处理窗口名中有特殊字符的窗口
Dim str1 As String
'Dim all_windows() As wininfo, arr_out(1 To 1000, 1 To 1)
'Dim all_windows() As wininfo, arr_out(1 To 1000, 1 To 1)
'Dim my_handle As Long, my_title As String * 100, title_len As Long, i As Long, j As Long
Dim my_handle As Long, my_title As String * 200, title_len As Long, i As Long, j As Long

'=================如果本身就是long并且有窗口名称则直接返回该long
If c_long(winname) > 0 Then
If GetWindowText(c_long(winname), my_title, Len(my_title)) > 0 Then
my_findwindow = c_long(winname)
Exit Function
End If
End If
'=================如果本身就是long并且有窗口名称则直接返回该long



my_handle = FindWindowEx(0, 0, vbNullString, vbNullString) '搜索第一个窗口,第一个0代表windows桌面，第二个0代表找第0个子窗口的下一个窗口
  Do While my_handle <> 0 '找到窗口的话
    DoEvents
    'my_title = String$(100, Chr$(0)) 'API调用只接受定长字符串，所以要么声明为 as string * 长度，要么就用这里的语句格式化一下
    title_len = GetWindowText(my_handle, my_title, Len(my_title))
    If title_len > 0 Then    '如果窗口标题存在，那么就会返回窗口标题的长度
        j = j + 1
        str1 = ""
        str1 = Left$(my_title, title_len) '按长度截取结果即可
        'If Len(str1) > Len(winname) Then
        'str1 = Left(str1, Len(winname))
       ' End If
      
        If str1 Like winname Then
        my_findwindow = my_handle
        Exit Function
        ElseIf InStr(str1, winname) > 0 Then
        my_findwindow = my_handle
        winname = str1
        Exit Function
        Else
       
          
        End If
    End If
    my_handle = FindWindowEx(0, my_handle, vbNullString, vbNullString) '继续找下一个窗口
Loop

If my_handle > 0 Then Exit Function


'第三次尝试： 紧紧留文件名进行匹配 删除Excel
str2 = Replace(winname, "Microsoft Excel - ", "")
my_handle = FindWindowEx(0, 0, vbNullString, vbNullString) '搜索第一个窗口,第一个0代表windows桌面，第二个0代表找第0个子窗口的下一个窗口
  Do While my_handle <> 0 '找到窗口的话
    DoEvents
    my_title = String$(100, Chr$(0)) 'API调用只接受定长字符串，所以要么声明为 as string * 长度，要么就用这里的语句格式化一下
    title_len = GetWindowText(my_handle, my_title, Len(my_title))
    If title_len > 0 Then    '如果窗口标题存在，那么就会返回窗口标题的长度
        j = j + 1
        str1 = ""
        str1 = Left$(my_title, title_len) '按长度截取结果即可
        'If Len(str1) > Len(winname) Then
        'str1 = Left(str1, Len(winname))
       ' End If
      
        If str1 Like str2 Then
        my_findwindow = my_handle
        Exit Function
        ElseIf InStr(str1, str2) > 0 Then
        winname = str1
        my_findwindow = my_handle
        Exit Function
        Else
       
          
        End If
    End If
    my_handle = FindWindowEx(0, my_handle, vbNullString, vbNullString) '继续找下一个窗口
Loop
'If my_handle = 0 Then MsgBox "Can Not Find Window:" & winname
End Function


Private Function c_long(str As String) As Long
On Error Resume Next
c_long = CLng(str)
End Function

'333333333333333333333333333333333
Function MEMO_Check_Win_Open(win_tit As String, Pic_s As StdPicture, Optional ByRef x As Integer, Optional ByRef y As Integer) As Boolean

'333_111
'Print_window_Memo
'333_222
'333_333



'Dim mpic_L As New C_Pic
'本函数用于判断指定窗口是否存在
MEMO_Check_Win_Open = False
Dim continue_b As Boolean
continue_b = True
Dim hwnd As Long
If Len(win_tit) = 0 Then
continue_b = False
End If
If continue_b Then
hwnd = FindWindow(vbNullString, win_tit)
If hwnd = 0 Then


    If hwnd = 0 Then
    'MEMO_Win_Pic_Leftclick = False
    'Exit Function
    hwnd = my_findwindow(win_tit)
    If hwnd = 0 Then
    MEMO_Check_Win_Open = False
    Exit Function
    End If
    End If
    


End If
End If
Dim L_pic As StdPicture
Set L_pic = Print_window_Memo(win_tit)


If L_pic Is Nothing Then
Set L_pic = Print_window_Memo()
End If




If L_pic.Handle = 0 Then continue_b = False




If continue_b Then

'mpic_L.loadpic "", L_pic
'If mpic_L.FindPic_S(Pic_s, x, y) = False Then
'continue_b = False
'End If

continue_b = GetPosi(Pic_s, L_pic, x, y)


End If

If continue_b = False Then
Exit Function
End If
MEMO_Check_Win_Open = True
End Function


Private Function GetPosi(Pic_s As StdPicture, Pic_L As StdPicture, Optional ByRef x As Integer, Optional ByRef y As Integer) As Boolean
'打开两张图
GetPosi = False
'Dim Pic_S As StdPicture
'Dim Pic_L As StdPicture
'Set Pic_L = LoadPicture(flfp_large_pic)
'Set Pic_S = LoadPicture(flfp_small_pic)
Dim iWidth_S As Integer    '以像素为单位的图形宽度
Dim iHeight_S As Integer    '以像素为单位的图形高度
Dim iWidth_L As Integer    '以像素为单位的图形宽度
Dim iHeight_L As Integer    '以像素为单位的图形高度

    Dim bits_L() As Byte
    Dim bits_S() As Byte    '三维数组，用于获取原彩色图像中各像素的RGB数值以及存放转化后的灰度值
    Dim bitsBW() As Byte    '三维数组，用于存放转化为黑白图后各像素的值
    
    
    '获取图形的宽度和高度
    iWidth_L = Pic_L.Width * HIMETRIC_PER_PIXEL
    iHeight_L = Pic_L.Height * HIMETRIC_PER_PIXEL
    
    iWidth_S = Pic_s.Width * HIMETRIC_PER_PIXEL
    iHeight_S = Pic_s.Height * HIMETRIC_PER_PIXEL
    
    
    Dim bi24BitInfo_L As BITMAPINFO
    With bi24BitInfo_L.bmiHeader
        .biBitCount = 32
        .biCompression = 0&
        .biPlanes = 1
        .biSize = Len(bi24BitInfo_L.bmiHeader)
        .biWidth = iWidth_L
        .biHeight = iHeight_L
    End With
    '重新定义数组大小
    ReDim bits_L(0 To 3, 0 To iWidth_L - 1, 0 To iHeight_L - 1) As Byte
    
    
    Dim bi24BitInfo_S As BITMAPINFO
    With bi24BitInfo_S.bmiHeader
        .biBitCount = 32
        .biCompression = 0&
        .biPlanes = 1
        .biSize = Len(bi24BitInfo_S.bmiHeader)
        .biWidth = iWidth_S
        .biHeight = iHeight_S
    End With
    '重新定义数组大小
    ReDim bits_S(0 To 3, 0 To iWidth_S - 1, 0 To iHeight_S - 1) As Byte
    
    
   Dim hdc As Long, hDCmem As Long
   
    
    
    hdc = GetDC(0)
    hDCmem = CreateCompatibleDC(hdc)
    

     
    '使用GetDIBits方法一次性获取picture1中各点的rgb值，比point方法或getPixel函数逐像素获取像素rgb要快出一个数量级
    GetDIBits hDCmem, Pic_L.Handle, 0&, iHeight_L, bits_L(0, 0, 0), bi24BitInfo_L, DIB_RGB_COLORS
    
    'DeleteDC hDCmem
    
    'hdc = GetDC(0)
    'hDCmem = CreateCompatibleDC(hdc)
    GetDIBits hDCmem, Pic_s.Handle, 0&, iHeight_S, bits_S(0, 0, 0), bi24BitInfo_S, DIB_RGB_COLORS
    
    DeleteDC hDCmem
    
    ReleaseDC 0, hdc
    
    
    Dim i As Integer, j As Integer
    Dim i2 As Integer, j2 As Integer
    
    
    Dim key1_x As Integer
    Dim key1_y As Integer
    Dim D_x As Integer
    Dim D_y As Integer
    
    GetSpecial_Point bits_S, iWidth_S - 1, iHeight_S - 1, 0, 0, key1_x, key1_y, D_x, D_y
        
    
    Dim key2_x As Integer
    Dim key2_y As Integer
    Dim D2_x As Integer
    Dim D2_y As Integer
    
    GetSpecial_Point bits_S, iWidth_S - 1, iHeight_S - 1, key1_x, key1_y, key2_x, key2_y, D2_x, D2_y
        
        
        
        
    Dim key3_x As Integer
    Dim key3_y As Integer
    Dim D3_x As Integer
    Dim D3_y As Integer
    GetSpecial_Point bits_S, iWidth_S - 1, iHeight_S - 1, iWidth_S / 2, iHeight_S / 2, key3_x, key3_y, D3_x, D3_y
        
        
                
    Dim key4_x As Integer
    Dim key4_y As Integer
    Dim D4_x As Integer
    Dim D4_y As Integer
    GetSpecial_Point bits_S, iWidth_S - 1, iHeight_S - 1, iWidth_S / 2, 0, key4_x, key4_y, D4_x, D4_y
        
        
    Dim key5_x As Integer
    Dim key5_y As Integer
    Dim D5_x As Integer
    Dim D5_y As Integer
    GetSpecial_Point bits_S, iWidth_S - 1, iHeight_S - 1, 0, iHeight_S / 2, key5_x, key5_y, D5_x, D5_y
        

    
     For j = 0 To iHeight_L - iHeight_S - 1
    
       VBA.DoEvents
        For i = 0 To iWidth_L - iWidth_S - 1
        
        
            '判断 点 Key1_x,Key1_y 是否相同
            
            If bits_L(2, i + key1_x, j + key1_y) <> bits_S(2, key1_x, key1_y) Then GoTo ExitLine: 'R
            If bits_L(1, i + key1_x, j + key1_y) <> bits_S(1, key1_x, key1_y) Then GoTo ExitLine: 'G
            If bits_L(0, i + key1_x, j + key1_y) <> bits_S(0, key1_x, key1_y) Then GoTo ExitLine: 'B
            
            '判断 点 Key2_x,Key2_y 是否相同
            
            If bits_L(2, i + key2_x, j + key2_y) <> bits_S(2, key2_x, key2_y) Then GoTo ExitLine: 'R
            If bits_L(1, i + key2_x, j + key2_y) <> bits_S(1, key2_x, key2_y) Then GoTo ExitLine: 'G
            If bits_L(0, i + key2_x, j + key2_y) <> bits_S(0, key2_x, key2_y) Then GoTo ExitLine: 'B
            
            '判断 点 Key3_x,Key3_y 是否相同
            
            If bits_L(2, i + key3_x, j + key3_y) <> bits_S(2, key3_x, key3_y) Then GoTo ExitLine: 'R
            If bits_L(1, i + key3_x, j + key3_y) <> bits_S(1, key3_x, key3_y) Then GoTo ExitLine: 'G
            If bits_L(0, i + key3_x, j + key3_y) <> bits_S(0, key3_x, key3_y) Then GoTo ExitLine: 'B
            
            '判断 点 Key4_x,Key4_y 是否相同
            
            If bits_L(2, i + key4_x, j + key4_y) <> bits_S(2, key4_x, key4_y) Then GoTo ExitLine: 'R
            If bits_L(1, i + key4_x, j + key4_y) <> bits_S(1, key4_x, key4_y) Then GoTo ExitLine: 'G
            If bits_L(0, i + key4_x, j + key4_y) <> bits_S(0, key4_x, key4_y) Then GoTo ExitLine: 'B
            
            
            
            '判断 点 Key5_x,Key5_y 是否相同
            
            If bits_L(2, i + key5_x, j + key5_y) <> bits_S(2, key5_x, key5_y) Then GoTo ExitLine: 'R
            If bits_L(1, i + key5_x, j + key5_y) <> bits_S(1, key5_x, key5_y) Then GoTo ExitLine: 'G
            If bits_L(0, i + key5_x, j + key5_y) <> bits_S(0, key5_x, key5_y) Then GoTo ExitLine: 'B
            
            
            
            
            
            
        
            '如果特征点相同，才开始 循环判断小图
            If bits_L(2, i, j) <> bits_S(2, 0, 0) Then GoTo ExitLine:   'R
            If bits_L(1, i, j) <> bits_S(1, 0, 0) Then GoTo ExitLine:   'G
            If bits_L(0, i, j) <> bits_S(0, 0, 0) Then GoTo ExitLine:   'B
            
            If bits_L(2, i + iWidth_S - 2, j) <> bits_S(2, iWidth_S - 2, 0) Then GoTo ExitLine: 'R
            If bits_L(1, i + iWidth_S - 2, j) <> bits_S(1, iWidth_S - 2, 0) Then GoTo ExitLine: 'G
            If bits_L(0, i + iWidth_S - 2, j) <> bits_S(0, iWidth_S - 2, 0) Then GoTo ExitLine: 'B
            
            If bits_L(2, i, j + iHeight_S - 2) <> bits_S(2, 0, iHeight_S - 2) Then GoTo ExitLine: 'R
            If bits_L(1, i, j + iHeight_S - 2) <> bits_S(1, 0, iHeight_S - 2) Then GoTo ExitLine: 'G
            If bits_L(0, i, j + iHeight_S - 2) <> bits_S(0, 0, iHeight_S - 2) Then GoTo ExitLine: 'B
            
            If bits_L(2, i + iWidth_S - 2, j + iHeight_S - 2) <> bits_S(2, iWidth_S - 2, iHeight_S - 2) Then GoTo ExitLine: 'R
            If bits_L(1, i + iWidth_S - 2, j + iHeight_S - 2) <> bits_S(1, iWidth_S - 2, iHeight_S - 2) Then GoTo ExitLine: 'G
            If bits_L(0, i + iWidth_S - 2, j + iHeight_S - 2) <> bits_S(0, iWidth_S - 2, iHeight_S - 2) Then GoTo ExitLine: 'B
            
  
            
            
            
            For j2 = 0 To iHeight_S - 2    '循环判断小图片
                For i2 = 0 To iWidth_S - 2


                    If bits_L(2, i + i2, j + j2) <> bits_S(2, i2, j2) Then GoTo ExitLine:    'R
                    If bits_L(1, i + i2, j + j2) <> bits_S(1, i2, j2) Then GoTo ExitLine:    'G
                    If bits_L(0, i + i2, j + j2) <> bits_S(0, i2, j2) Then GoTo ExitLine:    'B
                Next i2
            Next j2
            
            
            
            GetPosi = True
            x = i
            y = iHeight_L - j - iHeight_S
            Exit Function

ExitLine:
    Next i
Next j


'比较两张图
'给出返回值
End Function
Public Function delay(milliseconds As Long)
Dim delay_time As Double, start, now
start = Timer
delay_time = milliseconds / 1000
Do While now - start < delay_time
    now = Timer
    If now < start Then now = now + 86400
    DoEvents
Loop
End Function

Private Function Get_WR(ByVal hwnd As Long, ByRef L As Long, ByRef T As Long, ByRef R As Long, ByRef b As Long) As Boolean
Dim myrec As RECT
myrec.Bottom = b
myrec.Left = L
myrec.Top = T
myrec.Right = R
GetWindowRect hwnd, myrec
b = myrec.Bottom
L = myrec.Left
T = myrec.Top
R = myrec.Right
Get_WR = True

End Function
Private Function mouse_leftclick(x As Long, y As Long)
SetCursorPos x, y
mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
delay 100
End Function


Private Function Print_window_Memo(Optional ByVal win_title As String = Empty) As StdPicture

'fln_bmp,目标文件名合法性检测
Dim continue_b As Boolean
continue_b = True


If win_title = Empty Then
KeyJp (theScreen)


'SavePicture ApiGetClipBmp, fln_bmp
Dim mIPicD As IPictureDisp
Set mIPicD = ApiGetClipBmp
Set Print_window_Memo = CreateBitmapPicture(mIPicD.Handle)




Else
'============================================
    Dim hwnd As Long
    hwnd = FindWindow(vbNullString, win_title)
    If hwnd = 0 Then hwnd = my_findwindow(win_title)
    
    If hwnd <> 0 Then
    'AppActivate win_title
    Do While my_actwindow(win_title) = False
    Loop
    
    
    '窗体名称必须存在并且正确
    
    'AppActivate "Microsoft Excel (Product Activation Failed) - PX0495_CN.305872_TKSE_20190521_NEWSAP.xlsm"
    
    delay 600
    KeyJp (theform)
    delay 700
    'SavePicture ApiGetClipBmp, fln_bmp
    Set mIPicD = ApiGetClipBmp
    Set Print_window_Memo = CreateBitmapPicture(mIPicD.Handle)
        
    Else
    Exit Function
    End If
'============================================
End If

End Function


Function my_actwindow(wdn As String) As Boolean
'激活窗口,
On Error GoTo ErrorHand
AppActivate wdn
my_actwindow = True
Exit Function
ErrorHand:
delay 5000
my_actwindow = False
End Function

Private Function KeyJp(Optional ByVal TheMode As JpMode = theScreen) As IPictureDisp
Call keybd_event(vbKeySnapshot, TheMode, 0, 0)
DoEvents
'Set KeyJp = Clipboard.GetData
End Function

Private Function ApiGetClipBmp() As IPicture
'On Error Resume Next
Dim pic As PicBmp, IID_IDispatch As GUID
OpenClipboard 0 'OpenClipboard
With IID_IDispatch
.data1 = &H20400
.data4(0) = &HC0
.data4(7) = &H46
End With
With pic
.Size = Len(pic)
.Type = 1
.hBmp = GetClipboardData(CF_BITMAP)
End With
OleCreatePictureIndirect pic, IID_IDispatch, 1, ApiGetClipBmp
'stdole.SavePicture ApiGetClipBmp, "c:\clipboard.bmp"
CloseClipboard
End Function





Public Function CreateBitmapPicture(ByVal hBmp As Long) As StdPicture
     Dim pic As PicBmp
     'Dim IID_IDispatch(0 To 15) As Byte
    
        Dim clsid As GUID, iid As GUID
        
        
     CLSIDFromString StrPtr("{00020400-0000-0000-C000-000000000046}"), clsid
    
     With pic
         .Size = Len(pic)
         '.Type = vbPicTypeBitmap
         .Type = 1
         .hBmp = hBmp
         .hPal = 0
     End With
     OleCreatePictureIndirect pic, clsid, 1, CreateBitmapPicture
End Function


 Private Function GetSpecial_Point(bits, max_x As Integer, max_y As Integer, Zero_X As Integer, Zero_Y As Integer, ByRef dif_x As Integer, ByRef dif_y As Integer, ByRef x_distence As Integer, ByRef y_distence As Integer) As Boolean
Dim i, j  As Integer
Dim continue_b As Boolean
Dim x_stop_b As Boolean
Dim y_stop_b As Boolean
continue_b = True
x_stop_b = False
y_stop_b = False
x_distence = 0
y_distence = 0
i = 0
j = 0
Do While continue_b

If x_stop_b = False Then
x_distence = x_distence + 1
If x_distence + Zero_X >= max_x Then
x_stop_b = True
x_distence = x_distence - 1
Else
For i = Zero_Y To y_distence + Zero_Y
            If bits(2, Zero_X, Zero_Y) <> bits(2, Zero_X + x_distence, i) Then x_stop_b = True 'R
            If bits(1, Zero_X, Zero_Y) <> bits(1, Zero_X + x_distence, i) Then x_stop_b = True 'G
            If bits(0, Zero_X, Zero_Y) <> bits(0, Zero_X + x_distence, i) Then x_stop_b = True 'B
            If x_stop_b Then
            dif_x = Zero_X + x_distence
            dif_y = i
    
            x_distence = x_distence - 1
            Exit For
            End If
Next
End If
End If
If y_stop_b = False Then
y_distence = y_distence + 1
If y_distence + Zero_Y >= max_y Then
y_stop_b = True
y_distence = y_distence - 1
Else
For i = Zero_X To x_distence + Zero_X

            If bits(2, Zero_X, Zero_Y) <> bits(2, i, Zero_Y + y_distence) Then y_stop_b = True 'R
            If bits(1, Zero_X, Zero_Y) <> bits(1, i, Zero_Y + y_distence) Then y_stop_b = True 'G
            If bits(0, Zero_X, Zero_Y) <> bits(0, i, Zero_Y + y_distence) Then y_stop_b = True 'B
            If y_stop_b Then
            dif_x = i
            dif_y = Zero_Y + y_distence
    
            y_distence = y_distence - 1
            Exit For
            End If
Next
End If
End If
If y_stop_b And x_stop_b Then continue_b = False
Loop
End Function
