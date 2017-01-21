Attribute VB_Name = "CellInfo2KML"
Public Opt_Point As Boolean
Public Opt_Cell As Boolean

Sub Opt_Point_Click()
Opt_Point = True
Opt_Cell = False
End Sub
Sub Opt_Cell_Click()
Opt_Point = False
Opt_Cell = True
End Sub



Sub BtnClick()

Sheets(1).Range("E12").Value = ""
Sheets(1).Range("E13").Value = ""

Sheets(1).Range("F12").Value = ""
Sheets(1).Range("F13").Value = ""

If (Opt_Point) Then
    Call Point2KML
Else
    Call Cell2Kml
End If

End Sub


Sub Point2KML()

'Dim headFile As String
Dim tmpFile As String
Dim outFile As String

'headFile = ThisWorkbook.Path + "\headPoint.kml"
tmpFile = ThisWorkbook.Path + "\tmp.kml"
outFile = ThisWorkbook.Path + "\Point.kml"

Dim lat, lon As Double
Dim rows As Integer
Dim tmpStr As String

If Dir(tmpFile) <> "" Then Kill (tmpFile)

'Open headFile For Binary As #1
'tmpStr = InputB(LOF(1), #1)
'Close #1

tmpStr = ""
For i = 131 To 173
    tmpStr = tmpStr + Sheets("head").Range("A" & i).Value + vbLf
Next

Open tmpFile For Binary As #2
Put #2, , tmpStr

rows = Sheets(2).UsedRange.rows.Count

   

tmpStr = ""
tmpStr = tmpStr + vbTab + vbTab + vbLf + "<name>��վ</name>" + vbLf
Put #2, , tmpStr
tmpStr = ""

Dim row As Integer

row = 2
Do While rows - 1 > 0

    name = Sheets(2).Range("A" & row)
    lon = Sheets(2).Range("B" & row)
    lat = Sheets(2).Range("C" & row)
    
    tmpStr = tmpStr + vbTab + vbTab + "<Placemark>" + vbLf
    tmpStr = tmpStr + vbTab + vbTab + vbTab + "<name>" + name + "</name>" + vbLf
    tmpStr = tmpStr + vbTab + vbTab + vbTab + "<styleUrl>#m_ylw-pushpin</styleUrl>" + vbLf
    tmpStr = tmpStr + vbTab + vbTab + vbTab + "<Point>" + vbLf
    tmpStr = tmpStr + vbTab + vbTab + vbTab + vbTab + "<gx:drawOrder>1</gx:drawOrder>" + vbLf
    
    tmpStr = tmpStr + vbTab + vbTab + vbTab + vbTab + "<coordinates>" + CStr(lon) + "," + CStr(lat) + ",0</coordinates>" + vbLf
    
    tmpStr = tmpStr + vbTab + vbTab + vbTab + "</Point>" + vbLf
    tmpStr = tmpStr + vbTab + vbTab + "</Placemark>" + vbLf
    row = row + 1
    rows = rows - 1
    Sheets(1).Range("E12").Value = "�����" + CStr(row - 2) + "�������"
    Put #2, , tmpStr
    tmpStr = ""
Loop


tmpStr = ""
tmpStr = tmpStr + vbTab + "</Folder>" + vbLf
tmpStr = tmpStr + "</Document>" + vbLf
tmpStr = tmpStr + "</kml>"

Put #2, , tmpStr

Close #2

If Dir(outFile) <> "" Then Kill (outFile)

Call FileZM(tmpFile, "GB2312", outFile, "utf-8")

If Dir(tmpFile) <> "" Then Kill (tmpFile)



End Sub

Sub Cell2Kml()
Dim tmpFile As String
Dim outFile As String

tmpFile = ThisWorkbook.Path + "\tmp.kml"
If Dir(tmpFile) <> "" Then Kill (tmpFile)


outFile = ThisWorkbook.Path + "\Cell.kml"
Dim tmpStr As String
tmpStr = ""



Dim lat, lon, ang, radius As Double
Dim name As String
Dim bIO, bPwnDiv As Boolean

Dim row As Integer
Dim tmpStrInfo As String
Dim tmpStrPolygon As String

tmpStrInfo = ""
tmpStrPolygon = ""
name = ""

For i = 1 To 129
    tmpStr = tmpStr + Sheets("head").Range("A" & i).Value + vbLf
Next

Open tmpFile For Binary As #1
Put #1, , tmpStr
tmpStr = ""




tmpStr = tmpStr + vbLf + vbTab + vbTab + "<name>��վͼ��</name>" + vbLf
tmpStr = tmpStr + vbTab + vbTab + "<open>1</open>" + vbLf
Put #1, , tmpStr
tmpStr = ""

'����С����ϢFolder
tmpStr = ""
tmpStr = tmpStr + vbTab + vbTab + "<Folder>" + vbLf
tmpStr = tmpStr + vbTab + vbTab + vbTab + "<name>С����Ϣ</name>" + vbLf
tmpStr = tmpStr + vbTab + vbTab + vbTab + "<Style>" + vbLf
tmpStr = tmpStr + vbTab + vbTab + vbTab + vbTab + "<ListStyle>" + vbLf
tmpStr = tmpStr + vbTab + vbTab + vbTab + vbTab + vbTab + "<listItemType>check</listItemType>" + vbLf
tmpStr = tmpStr + vbTab + vbTab + vbTab + vbTab + vbTab + "<bgColor>00ffffff</bgColor>" + vbLf
tmpStr = tmpStr + vbTab + vbTab + vbTab + vbTab + vbTab + "<maxSnippetLines>2</maxSnippetLines>" + vbLf
tmpStr = tmpStr + vbTab + vbTab + vbTab + vbTab + "</ListStyle>" + vbLf
tmpStr = tmpStr + vbTab + vbTab + vbTab + "</Style>" + vbLf
Put #1, , tmpStr
tmpStr = ""

For i = 2 To Sheets(2).UsedRange.rows.Count

    name = Sheets(2).Range("A" & i)

    lon = Sheets(2).Range("B" & i)
    If IsNumeric(lon) = False Then
        Sheets(1).Range("F12").Value = "��" + CStr(i - 1) + "�о��ȴ���"
        GoTo Err
    End If
    lat = Sheets(2).Range("C" & i)
    If IsNumeric(lat) = False Then
        Sheets(1).Range("F12").Value = "��" + CStr(i - 1) + "��γ�ȴ���"
       GoTo Err
    End If
    ang = Sheets(2).Range("D" & i)
    If IsNumeric(ang) = False Then
        Sheets(1).Range("F12").Value = "��" + CStr(i - 1) + "�з�λ�Ǵ���"
        GoTo Err
    End If
    radius = Sheets(2).Range("F" & i)
    If IsNumeric(radius) = False Then
        Sheets(1).Range("F12").Value = "��" + CStr(i - 1) + "�а뾶����"
        GoTo Err
    End If


    'С����Ϣ����
    tmpStrInfo = vbTab + vbTab + "<Placemark>" + vbLf
    tmpStrInfo = tmpStrInfo + vbTab + vbTab + vbTab + "<name>" + name + "</name>" + vbLf
    tmpStrInfo = tmpStrInfo + vbTab + vbTab + vbTab + "<styleUrl>#msn_wht-blank</styleUrl>" + vbLf
    tmpStrInfo = tmpStrInfo + vbTab + vbTab + vbTab + "<Point>" + vbLf
    tmpStrInfo = tmpStrInfo + vbTab + vbTab + vbTab + vbTab + "<gx:drawOrder>1</gx:drawOrder>" + vbLf
    strInOrOut = Sheets(2).Range("E" & i)
    If (strInOrOut = "����") Then
        tmpStrInfo = tmpStrInfo + vbTab + vbTab + vbTab + vbTab + "<coordinates>" + CStr(lon) + "," + CStr(lat) + ",0 " + "</coordinates>" + vbLf
    Else
        tmpStrInfo = tmpStrInfo + vbTab + vbTab + vbTab + vbTab + "<coordinates>" + Computation.Computation(lat, lon, ang, 70) + ",0 " + "</coordinates>" + vbLf
    End If
    tmpStrInfo = tmpStrInfo + vbTab + vbTab + vbTab + "</Point>" + vbLf
    tmpStrInfo = tmpStrInfo + vbTab + vbTab + "</Placemark>" + vbLf
    Put #1, , tmpStrInfo
    tmpStrInfo = ""
    Sheets(1).Range("E12").Value = "�����" + CStr(i - 1) + "��С����Ϣ"
    
Next

tmpStr = ""
tmpStr = tmpStr + vbTab + vbTab + "</Folder>" + vbLf
Put #1, , tmpStr


'����С��ͼ��Folder
tmpStr = ""
tmpStr = tmpStr + vbTab + vbTab + "<Folder>" + vbLf
tmpStr = tmpStr + vbTab + vbTab + vbTab + "<name>С��ͼ��</name>" + vbLf
Put #1, , tmpStr

For i = 2 To Sheets(2).UsedRange.rows.Count

    name = Sheets(2).Range("A" & i)

    lon = Sheets(2).Range("B" & i)
    If IsNumeric(lon) = False Then
        Sheets(1).Range("F12").Value = "��" + CStr(i - 1) + "�о��ȴ���"
        GoTo Err
    End If
    lat = Sheets(2).Range("C" & i)
    If IsNumeric(lat) = False Then
        Sheets(1).Range("F12").Value = "��" + CStr(i - 1) + "��γ�ȴ���"
       GoTo Err
    End If
    ang = Sheets(2).Range("D" & i)
    If IsNumeric(ang) = False Then
        Sheets(1).Range("F12").Value = "��" + CStr(i - 1) + "�з�λ�Ǵ���"
        GoTo Err
    End If
    radius = Sheets(2).Range("F" & i)
    If IsNumeric(radius) = False Then
        Sheets(1).Range("F12").Value = "��" + CStr(i - 1) + "�а뾶����"
        GoTo Err
    End If
    'С��ͼ�δ���
    tmpStrPolygon = tmpStrPolygon + vbTab + vbTab + vbTab + "<Placemark>" + vbLf
    tmpStrPolygon = tmpStrPolygon + vbTab + vbTab + vbTab + "<name>" + name + "</name>" + vbLf
    tmpStrPolygon = tmpStrPolygon + vbTab + vbTab + vbTab + "<description>"
    tmpStrPolygon = tmpStrPolygon + "<![CDATA[<table border=1 width=360>"
    tmpStrPolygon = tmpStrPolygon + "<tr><th>С������</th><th>����</th><th>ΰ��</th></tr>"
    tmpStrPolygon = tmpStrPolygon + "<tr><td>" + name + "</td><td>" + CStr(lon) + "</td><td>" + CStr(lat) + "</td></tr>"
    tmpStrPolygon = tmpStrPolygon + "<tr><th>��λ��</th><th>վ������</th><th>�뾶</th></tr>"
    tmpStrPolygon = tmpStrPolygon + "<tr><td>" + CStr(ang) + "</td><td>" + Sheets(2).Range("E" & i) + "</td><td>" + CStr(Sheets(2).Range("F" & i)) + "</td></tr>"
    tmpStrPolygon = tmpStrPolygon + "</table>]]> " + "</description>" + vbLf

    cellid = Sheets(2).Range("G" & i)
    ID = cellid Mod 3
    tmpStrPolygon = tmpStrPolygon + vbTab + vbTab + vbTab + "<styleUrl>#msn_ylw-pushpin" + CStr(ID) + "</styleUrl>" + vbLf
    tmpStrPolygon = tmpStrPolygon + vbTab + vbTab + vbTab + "<Polygon>" + vbLf
    tmpStrPolygon = tmpStrPolygon + vbTab + vbTab + vbTab + vbTab + "<tessellate>1</tessellate>" + vbLf
    tmpStrPolygon = tmpStrPolygon + vbTab + vbTab + vbTab + vbTab + "<outerBoundaryIs>" + vbLf
    tmpStrPolygon = tmpStrPolygon + vbTab + vbTab + vbTab + vbTab + "<LinearRing>" + vbLf
    tmpStrPolygon = tmpStrPolygon + vbTab + vbTab + vbTab + vbTab + vbTab + "<coordinates>" + vbLf

    strInOrOut = Sheets(2).Range("E" & i)
    radius = Sheets(2).Range("F" & i)
    If (radius < 100) Then radius = 100
    If (radius > 3000) Then radius = 3000

    A1 = (ang - 30) Mod 360
    A2 = (ang - 15) Mod 360
    A3 = ang
    A4 = (ang + 15) Mod 360
    A5 = (ang + 30) Mod 360
    If (strInOrOut = "����") Then
        tmpStrPolygon = tmpStrPolygon + vbTab + vbTab + vbTab + vbTab + vbTab
        For j = 1 To 36
            tmpStrPolygon = tmpStrPolygon + Computation.Computation(lat, lon, j * 10 - 1, 30) + ",0 "
        Next
        tmpStrPolygon = tmpStrPolygon + vbLf
    Else
            tmpStrPolygon = tmpStrPolygon + vbTab + vbTab + vbTab + vbTab + vbTab + CStr(lon) + "," + CStr(lat) + "," + ",0 " 'ԭ������
            tmpStrPolygon = tmpStrPolygon + Computation.Computation(lat, lon, A1, radius) + ",0 " '��һ��
            tmpStrPolygon = tmpStrPolygon + Computation.Computation(lat, lon, A2, radius) + ",0 "
            tmpStrPolygon = tmpStrPolygon + Computation.Computation(lat, lon, A3, radius) + ",0 "
            tmpStrPolygon = tmpStrPolygon + Computation.Computation(lat, lon, A4, radius) + ",0 "
            tmpStrPolygon = tmpStrPolygon + Computation.Computation(lat, lon, A5, radius) + ",0 "
            tmpStrPolygon = tmpStrPolygon + CStr(lon) + "," + CStr(lat) + "," + ",0 " 'ԭ������
            tmpStrPolygon = tmpStrPolygon + vbLf
    End If

    tmpStrPolygon = tmpStrPolygon + vbTab + vbTab + vbTab + vbTab + vbTab + "</coordinates>" + vbLf
    tmpStrPolygon = tmpStrPolygon + vbTab + vbTab + vbTab + vbTab + "</LinearRing>" + vbLf
    tmpStrPolygon = tmpStrPolygon + vbTab + vbTab + vbTab + vbTab + "</outerBoundaryIs>" + vbLf
    tmpStrPolygon = tmpStrPolygon + vbTab + vbTab + vbTab + "</Polygon>" + vbLf
    tmpStrPolygon = tmpStrPolygon + vbTab + vbTab + vbTab + "</Placemark>" + vbLf
    Put #1, , tmpStrPolygon
    tmpStrPolygon = ""
    Sheets(1).Range("E13").Value = "�����" + CStr(i - 1) + "��С��ͼ��"


Next

tmpStr = ""
tmpStr = tmpStr + vbTab + vbTab + "</Folder>" + vbLf
Put #1, , tmpStr

'KML�ļ�β
tmpStr = ""
tmpStr = tmpStr + "</Folder>" + vbLf
tmpStr = tmpStr + "</Document>" + vbLf
tmpStr = tmpStr + "</kml>"
Put #1, , tmpStr
Close #1

If Dir(outFile) <> "" Then Kill (outFile)
Call FileZM(tmpFile, "GB2312", outFile, "utf-8")
If Dir(tmpFile) <> "" Then Kill (tmpFile)
Err:
    Close #1
End Sub

Sub FileZM(sFile As String, sCode As String, dFile As String, dCode As String)
'������Դ�ļ���Դ�ļ����룬Ŀ���ļ���Ŀ���ļ����롣�������----"gb2312"��"UTF-8"��
Dim ObjStream As Object

Set ObjStream = CreateObject("Adodb.Stream")
With ObjStream
    .Mode = 3         'adModeReadWrite = 3 ' ָʾ��/дȨ�ޡ�
    .Type = 1         'adTypeBinary = 1
    .Open
    .LoadFromFile sFile   'Դ�ļ�

    .Position = 0
    .Type = 2         'adTypeText = 2
    .Charset = sCode
    sCode = .ReadText '��ȡ�ı���sCode
    
    .Position = 0     ' ��ֻ�Ƕ�λ���ļ�ͷ������
    .SetEOS           ' ��ȫ��д��Ҫ©�����,ͨ��ʹ��ǰ Position ��Ϊ���Ľ�β������ EOS ���Ե�ֵ����ǰλ�ú���������ֽڻ��ַ��������ض�
    .Type = 2         'adTypeText = 2
    .Charset = dCode       'ָ���������
    .WriteText sCode       'д��ָ�����ı����ݵ�Adodb.Stream
     .SaveToFile dFile, 2
    .Close
End With
Set ObjStream = Nothing
End Sub
