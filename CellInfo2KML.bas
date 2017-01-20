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

If (Opt_Point) Then
    Call Point2KML
Else
    Call Cell2KML
End If

End Sub


Sub Point2KML()

Dim headFile As String
Dim tmpFile As String
Dim outFile As String

headFile = ThisWorkbook.Path + "\headPoint.kml"
tmpFile = ThisWorkbook.Path + "\tmp.kml"
outFile = ThisWorkbook.Path + "\Point.kml"

Dim lat, lon As Double
Dim rows As Integer

If Dir(tmpFile) <> "" Then Kill (tmpFile)

Open headFile For Binary As #1
Open tmpFile For Binary As #2
Dim tmpStr As String

tmpStr = InputB(LOF(1), #1)
Close #1

Put #2, , tmpStr

rows = Sheets(2).UsedRange.rows.Count

tmpStr = ""
tmpStr = tmpStr + vbTab + vbTab + vbLf + "<name>��վ</name>" + vbLf

Put #2, , tmpStr

Dim row As Integer

row = 2

    tmpStr = ""

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

MsgBox row - 2

End Sub

Sub Cell2KML()

Dim headFile As String
Dim tmpFile As String
Dim outFile As String

headFile = ThisWorkbook.Path + "\headCell.kml"
tmpFile = ThisWorkbook.Path + "\tmp.kml"
outFile = ThisWorkbook.Path + "\Cell.kml"

Dim lat, lon As Double
Dim rows As Integer

If Dir(tmpFile) <> "" Then Kill (tmpFile)

Open headFile For Binary As #1
Open tmpFile For Binary As #2
Dim tmpStr As String

tmpStr = InputB(LOF(1), #1)
Close #1
Put #2, , tmpStr

rows = Sheets(2).UsedRange.rows.Count

tmpStr = ""
tmpStr = tmpStr + vbTab + vbTab + "<name>��վͼ��</name>" + vbLf
tmpStr = tmpStr + vbTab + vbTab + "<open>1</open>" + vbLf
tmpStr = tmpStr + vbTab + vbTab + "<Folder>" + vbLf
tmpStr = tmpStr + vbTab + vbTab + vbTab + "<name>С����Ϣ</name>" + vbLf
tmpStr = tmpStr + vbTab + vbTab + vbTab + "<Style>" + vbLf
tmpStr = tmpStr + vbTab + vbTab + vbTab + vbTab + "<ListStyle>" + vbLf
tmpStr = tmpStr + vbTab + vbTab + vbTab + vbTab + vbTab + "<listItemType>check</listItemType>" + vbLf
tmpStr = tmpStr + vbTab + vbTab + vbTab + vbTab + vbTab + "<bgColor>00ffffff</bgColor>" + vbLf
tmpStr = tmpStr + vbTab + vbTab + vbTab + vbTab + vbTab + "<maxSnippetLines>2</maxSnippetLines>" + vbLf
tmpStr = tmpStr + vbTab + vbTab + vbTab + vbTab + "</ListStyle>" + vbLf
tmpStr = tmpStr + vbTab + vbTab + vbTab + "</Style>" + vbLf
Put #2, , tmpStr
tmpStr = ""

row = 2
Do While rows - 1 > 0

    name = Sheets(2).Range("A" & row)
    lon = Sheets(2).Range("B" & row)
    lat = Sheets(2).Range("C" & row)
    
    tmpStr = tmpStr + vbTab + vbTab + "<Placemark>" + vbLf
    tmpStr = tmpStr + vbTab + vbTab + vbTab + "<name>" + name + "</name>" + vbLf
    tmpStr = tmpStr + vbTab + vbTab + vbTab + "<styleUrl>#msn_wht-blank</styleUrl>" + vbLf
    tmpStr = tmpStr + vbTab + vbTab + vbTab + "<Point>" + vbLf
    tmpStr = tmpStr + vbTab + vbTab + vbTab + vbTab + "<gx:drawOrder>1</gx:drawOrder>" + vbLf
    
    tmpStr = tmpStr + vbTab + vbTab + vbTab + vbTab + "<coordinates>" + CStr(lon) + "," + CStr(lat) + ",0</coordinates>" + vbLf
    
    tmpStr = tmpStr + vbTab + vbTab + vbTab + "</Point>" + vbLf
    tmpStr = tmpStr + vbTab + vbTab + "</Placemark>" + vbLf
    row = row + 1
    rows = rows - 1
    Put #2, , tmpStr
    tmpStr = ""
Loop

tmpStr = ""
tmpStr = tmpStr + vbTab + vbTab + "</Folder>" + vbLf
Put #2, , tmpStr

tmpStr = ""
tmpStr = tmpStr + vbTab + vbTab + "<Folder>" + vbLf
tmpStr = tmpStr + vbTab + vbTab + vbTab + "<name>С��ͼ��</name>" + vbLf
Put #2, , tmpStr
tmpStr = ""


row = 2
rows = Sheets(2).UsedRange.rows.Count
Do While rows - 1 > 0

    name = Sheets(2).Range("A" & row)
    lon = Sheets(2).Range("B" & row)
    lat = Sheets(2).Range("C" & row)
    ang = Sheets(2).Range("D" & row)
    A1 = (ang - 30) Mod 360
    A2 = (ang - 15) Mod 360
    A3 = ang
    A4 = (ang + 15) Mod 360
    A5 = (ang + 30) Mod 360
    
    tmpStr = tmpStr + vbTab + vbTab + vbTab + "<Placemark>" + vbLf
    tmpStr = tmpStr + vbTab + vbTab + vbTab + "<name>" + name + "</name>" + vbLf
    tmpStr = tmpStr + vbTab + vbTab + vbTab + "<description>" + name + "</description>" + vbLf
    tmpStr = tmpStr + vbTab + vbTab + vbTab + "<styleUrl>#msn_ylw-pushpin0</styleUrl>" + vbLf
    tmpStr = tmpStr + vbTab + vbTab + vbTab + "<Polygon>" + vbLf
    tmpStr = tmpStr + vbTab + vbTab + vbTab + vbTab + "<tessellate>1</tessellate>" + vbLf
    tmpStr = tmpStr + vbTab + vbTab + vbTab + vbTab + "<outerBoundaryIs>" + vbLf
    tmpStr = tmpStr + vbTab + vbTab + vbTab + vbTab + "<LinearRing>" + vbLf
    tmpStr = tmpStr + vbTab + vbTab + vbTab + vbTab + vbTab + "<coordinates>" + vbLf
    
    tmpStr = tmpStr + vbTab + vbTab + vbTab + vbTab + vbTab + CStr(lon) + "," + CStr(lat) + "," + ",0 " 'ԭ������
    tmpStr = tmpStr + Computation.Computation(lat, lon, A1, 200) + ",0 " '��һ��
    tmpStr = tmpStr + Computation.Computation(lat, lon, A2, 200) + ",0 "
    tmpStr = tmpStr + Computation.Computation(lat, lon, A3, 200) + ",0 "
    tmpStr = tmpStr + Computation.Computation(lat, lon, A4, 200) + ",0 "
    tmpStr = tmpStr + Computation.Computation(lat, lon, A5, 200) + ",0 "
    
    tmpStr = tmpStr + vbTab + vbTab + vbTab + vbTab + vbTab + CStr(lon) + "," + CStr(lat) + "," + ",0 " 'ԭ������
    
    tmpStr = tmpStr + vbTab + vbTab + vbTab + vbTab + vbTab + "</coordinates>" + vbLf
    tmpStr = tmpStr + vbTab + vbTab + vbTab + vbTab + "</LinearRing>" + vbLf
    tmpStr = tmpStr + vbTab + vbTab + vbTab + vbTab + "</outerBoundaryIs>" + vbLf
    tmpStr = tmpStr + vbTab + vbTab + vbTab + "</Polygon>" + vbLf
    tmpStr = tmpStr + vbTab + vbTab + vbTab + "</Placemark>" + vbLf
    row = row + 1
    rows = rows - 1
    Put #2, , tmpStr
    tmpStr = ""
Loop

tmpStr = ""
tmpStr = tmpStr + vbTab + vbTab + "</Folder>" + vbLf
Put #2, , tmpStr


tmpStr = ""
tmpStr = tmpStr + vbTab + "</Folder>" + vbLf
tmpStr = tmpStr + "</Document>" + vbLf
tmpStr = tmpStr + "</kml>"
Put #2, , tmpStr
Close #2

If Dir(outFile) <> "" Then Kill (outFile)
Call FileZM(tmpFile, "GB2312", outFile, "utf-8")
If Dir(tmpFile) <> "" Then Kill (tmpFile)

MsgBox row - 2

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
