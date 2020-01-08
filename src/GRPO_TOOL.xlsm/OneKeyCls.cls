VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OneKeyCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'FileName : OneKeyCls.cls

Public Key As String




'�ֲ����������漯��
Private mCol As Collection

Public Function Add(Key As String, Optional skey As String) As OneKeyCls
    '�����¶���
    Dim objNewMember As OneKeyCls
    Set objNewMember = New OneKeyCls
    
    
    '���ô��뷽��������
    objNewMember.Key = Key
    If Len(skey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, skey
    End If
    
    
    '�����Ѵ����Ķ���
    Set Add = objNewMember
    Set objNewMember = Nothing
    
    
End Function

'���� ���� ���������� û�� �ؼ��� ��ʱ�� �ͻ᷵��Nothing�ˡ�
Public Property Get Item(vntIndexKey As Variant) As OneKeyCls
    Err.Clear
    On Error GoTo GetItemError
    '���ü����е�һ��Ԫ��ʱʹ�á�
    'vntIndexKey �������ϵ�������ؼ��֣�
    '����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��Set foo = x.Item(xyz) or Set foo = x.Item(5)
    Set Item = mCol(vntIndexKey)
    Exit Property
    
GetItemError:
    Debug.Print Err.Source & " : Error Was found in Item(vntIndexKey as Variant) as oneKeyCls . " & vbCrLf & Err.Description & " : " & Err.Number
    Set Item = Nothing
End Property



Public Property Get Count() As Long
    '���������е�Ԫ����ʱʹ�á��﷨��Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'ɾ�������е�Ԫ��ʱʹ�á�
    'vntIndexKey ����������ؼ��֣�����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
    '������������ For...Each �﷨ö�ٸü��ϡ�
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    '������󴴽�����
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    '����ֹ���ƻ�����
    Set mCol = Nothing
End Sub

Public Function To_excel(fdn As String)
'���������ڽ����ݵ�����ָ���ļ���
Dim mfso As New CFSO
If mfso.FileExists(fdn) = False Then
mfso.CreateFolder fdn
End If
Set mfso = Nothing
Dim mokc0 As OneKeyCls
Dim wb As Workbook
Dim i As Integer
If Right(fdn, 1) <> "\" Then fdn = fdn & "\"
For Each mokc0 In mCol

If mokc0.Count >= 1 Then
    
    Set wb = Workbooks.Add
    wb.SaveAs fdn & Replace(mokc0.Key, ":", "") & ".xlsx"
    
    If mokc0.Count > 0 Then
    
    For i = 1 To mokc0.Count
    wb.Worksheets(1).Range("A" & i) = mokc0.Item(i).Key
    Next
    End If
    
    If Len(mokc0.Item(1).Key) > 5 Then
    wb.Worksheets(1).Columns("A:A").WrapText = False
    End If
    
        
    wb.Save
    wb.Saved = True
    wb.Close
    Set wb = Nothing
    
End If


Next mokc0
End Function