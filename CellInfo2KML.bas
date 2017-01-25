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

Sheets(1).Range("E15").Value = ""


If (Opt_Point) Then
    Call Point2KML
Else
    Call Cell2Kml
End If

End Sub


Sub Point2KML()


Dim tmpFile As String
Dim outFile As String

tmpFile = ThisWorkbook.Path + "\tmp.kml"
outFile = ThisWorkbook.Path + "\Point.kml"

Dim lat, lon As Double
Dim rows As Integer
Dim columns As Integer

Dim tmpStr As String

If Dir(tmpFile) <> "" Then Kill (tmpFile)


tmpStr = ""
For i = 131 To 173
    tmpStr = tmpStr + Sheets("head").Range("A" & i).Value + vbLf
Next

Open tmpFile For Binary As #2
Put #2, , tmpStr

rows = Sheets(2).UsedRange.rows.Count
columns = Sheets(2).UsedRange.columns.Count

   

tmpStr = ""
tmpStr = tmpStr + vbTab + vbTab + vbLf + "<name>��վ</name>" + vbLf
Put #2, , tmpStr
tmpStr = ""



For i = 2 To rows

    name = Sheets(2).Range("A" & i)
    lon = Sheets(2).Range("B" & i)
    lat = Sheets(2).Range("C" & i)
    
    tmpStr = tmpStr + vbTab + vbTab + "<Placemark>" + vbLf
    tmpStr = tmpStr + vbTab + vbTab + vbTab + "<name>" + name + "</name>" + vbLf
    tmpStr = tmpStr + vbTab + vbTab + vbTab + "<styleUrl>#m_ylw-pushpin</styleUrl>" + vbLf
    tmpStr = tmpStr + vbTab + vbTab + vbTab + "<Point>" + vbLf
    tmpStr = tmpStr + vbTab + vbTab + vbTab + vbTab + "<gx:drawOrder>1</gx:drawOrder>" + vbLf
    
    tmpStr = tmpStr + vbTab + vbTab + vbTab + vbTab + "<coordinates>" + CStr(lon) + "," + CStr(lat) + ",0</coordinates>" + vbLf
    
    tmpStr = tmpStr + vbTab + vbTab + vbTab + "</Point>" + vbLf
    tmpStr = tmpStr + vbTab + vbTab + "</Placemark>" + vbLf

    Sheets(1).Range("E12").Value = "�����" + CStr(i - 1) + "�������"
    

    
    Put #2, , tmpStr
    tmpStr = ""
Next


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

Dim rows, columns As Integer

rows = Sheets(2).UsedRange.rows.Count
columns = Sheets(2).UsedRange.columns.Count

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

For i = 2 To rows

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
    
        
        Dim angles As Variant
        Dim angCount As Variant
        angles = Split(ang, "/")
        angCount = UBound(angles)
        Dim bAngle As Boolean
        bAngle = True
        For j = 0 To angCount
            If (angles(j) < 0) Then bAngle = False
            If (angles(j) > 360) Then bAngle = False
            
            bAngle = bAngle And (IsNumeric(angles(j)))
        Next
        If (bAngle = False) Then
            Sheets(1).Range("F12").Value = "��" + CStr(i - 1) + "�з�λ�Ǵ���"
            GoTo Err
        End If
    
    radius = Sheets(2).Range("F" & i)
    If IsNumeric(radius) = False Then
        Sheets(1).Range("F12").Value = "��" + CStr(i - 1) + "�а뾶����"
        GoTo Err
    End If


    'С����Ϣ����
    For j = 0 To angCount 'С����angCount������
    
        tmpStrInfo = vbTab + vbTab + "<Placemark>" + vbLf
        tmpStrInfo = tmpStrInfo + vbTab + vbTab + vbTab + "<name>" + name + "</name>" + vbLf
        tmpStrInfo = tmpStrInfo + vbTab + vbTab + vbTab + "<styleUrl>#msn_wht-blank</styleUrl>" + vbLf
        tmpStrInfo = tmpStrInfo + vbTab + vbTab + vbTab + "<Point>" + vbLf
        tmpStrInfo = tmpStrInfo + vbTab + vbTab + vbTab + vbTab + "<gx:drawOrder>1</gx:drawOrder>" + vbLf
        strInOrOut = Sheets(2).Range("E" & i)
        If (strInOrOut = "����" Or strInOrOut = "�ҷ�") Then
            tmpStrInfo = tmpStrInfo + vbTab + vbTab + vbTab + vbTab + "<coordinates>" + CStr(lon) + "," + CStr(lat) + ",0 " + "</coordinates>" + vbLf
            tmpStrInfo = tmpStrInfo + vbTab + vbTab + vbTab + "</Point>" + vbLf
            tmpStrInfo = tmpStrInfo + vbTab + vbTab + "</Placemark>" + vbLf
            Put #1, , tmpStrInfo
            tmpStrInfo = ""
            Exit For
        Else
            
            tmpStrInfo = tmpStrInfo + vbTab + vbTab + vbTab + vbTab + "<coordinates>" + Computation.Computation(lat, lon, angles(j), 70) + ",0 " + "</coordinates>" + vbLf
        End If
        tmpStrInfo = tmpStrInfo + vbTab + vbTab + vbTab + "</Point>" + vbLf
        tmpStrInfo = tmpStrInfo + vbTab + vbTab + "</Placemark>" + vbLf
        Put #1, , tmpStrInfo
        tmpStrInfo = ""
        
    Next
    
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

For i = 2 To rows 'Sheets(2).UsedRange.rows.Count

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
        'Dim angles As Variant
        'Dim angCount As Variant
        angles = Split(ang, "/")
        angCount = UBound(angles)
        'Dim bAngle As Boolean
        bAngle = True
        For j = 0 To angCount
            If (angles(j) < 0) Then bAngle = False
            If (angles(j) > 360) Then bAngle = False
            
            bAngle = bAngle And (IsNumeric(angles(j)))
        Next
        If (bAngle = False) Then
            Sheets(1).Range("F12").Value = "��" + CStr(i - 1) + "�з�λ�Ǵ���"
            GoTo Err
        End If
    radius = Sheets(2).Range("F" & i)
    
    If IsNumeric(radius) = False Then
        Sheets(1).Range("F12").Value = "��" + CStr(i - 1) + "�а뾶����"
        GoTo Err
    End If
    'С��ͼ�δ���
    
    For k = 0 To angCount '��angCount�����֣���angCount������
    
        tmpStrPolygon = tmpStrPolygon + vbTab + vbTab + vbTab + "<Placemark>" + vbLf
        tmpStrPolygon = tmpStrPolygon + vbTab + vbTab + vbTab + "<name>" + name + "</name>" + vbLf
        tmpStrPolygon = tmpStrPolygon + vbTab + vbTab + vbTab + "<description>"
        tmpStrPolygon = tmpStrPolygon + "<![CDATA[<table border=1 width=650>"
        Dim col As Integer
        col = 1
        For n = 1 To (Int(columns / 4))
                   
            tmpStrPolygon = tmpStrPolygon + "<tr><th>" + CStr(Sheets(2).Cells(1, col)) + "</th><th>" + CStr(Sheets(2).Cells(1, col + 1)) + "</th><th>" + CStr(Sheets(2).Cells(1, col + 2)) + "</th><th>" + CStr(Sheets(2).Cells(1, col + 3)) + "</th></tr>"
            'tmpStrPolygon = tmpStrPolygon + "<tr><td>" + Sheets(2).Cells(2, col) + "</td><td>" + Sheets(2).Cells(2, col + 1) + "</td><td>" + Sheets(2).Cells(2, col + 2) + "</td><td>" + Sheets(2).Cells(2, col + 3) + "</td></tr>"
            tmpStrPolygon = tmpStrPolygon + "<tr><td>" + CStr(Sheets(2).Cells(i, col)) + "</td><td>" + CStr(Sheets(2).Cells(i, col + 1)) + "</td><td>" + CStr(Sheets(2).Cells(i, col + 2)) + "</td><td>" + CStr(Sheets(2).Cells(i, col + 3)) + "</td></tr>"
                       
             col = col + 4
        Next
        
        tmpStrPolygon = tmpStrPolygon + "</table>]]> " + "</description>" + vbLf
    
        cellid = Sheets(2).Range("G" & i)
        ID = cellid Mod 3
        tmpStrPolygon = tmpStrPolygon + vbTab + vbTab + vbTab + "<styleUrl>#msn_ylw-pushpin" + CStr(ID) + "</styleUrl>" + vbLf
        tmpStrPolygon = tmpStrPolygon + vbTab + vbTab + vbTab + "<Polygon>" + vbLf
        tmpStrPolygon = tmpStrPolygon + vbTab + vbTab + vbTab + vbTab + "<tessellate>1</tessellate>" + vbLf
        tmpStrPolygon = tmpStrPolygon + vbTab + vbTab + vbTab + vbTab + "<outerBoundaryIs>" + vbLf
        tmpStrPolygon = tmpStrPolygon + vbTab + vbTab + vbTab + vbTab + "<LinearRing>" + vbLf
        tmpStrPolygon = tmpStrPolygon + vbTab + vbTab + vbTab + vbTab + vbTab + "<coordinates>" + vbLf
    
       
        radius = Sheets(2).Range("F" & i)
        coverType = Sheets(2).Range("H" & i)
        Select Case coverType
            Case "����"
                radius = radius * 25
                
            Case "����"
                radius = radius * 50
                
            Case "ũ��"
                radius = radius * 100
            Case Else
                radius = radius * 35
        End Select
        
        'radius = radius * 40
        If (radius < 100) Then radius = 100
        If (radius > 3000) Then radius = 3000
    
        A1 = (angles(k) - 30) Mod 360
        
        A2 = (angles(k) - 15) Mod 360
        A3 = angles(k)
        A4 = (angles(k) + 15) Mod 360
        A5 = (angles(k) + 30) Mod 360
        
        strInOrOut = Sheets(2).Range("E" & i)
        If (strInOrOut = "����" Or strInOrOut = "�ҷ�") Then
            tmpStrPolygon = tmpStrPolygon + vbTab + vbTab + vbTab + vbTab + vbTab
            For j = 1 To 36
                tmpStrPolygon = tmpStrPolygon + Computation.Computation(lat, lon, j * 10 - 1, 30) + ",0 "
            Next
            tmpStrPolygon = tmpStrPolygon + vbLf
            tmpStrPolygon = tmpStrPolygon + vbTab + vbTab + vbTab + vbTab + vbTab + "</coordinates>" + vbLf
            tmpStrPolygon = tmpStrPolygon + vbTab + vbTab + vbTab + vbTab + "</LinearRing>" + vbLf
            tmpStrPolygon = tmpStrPolygon + vbTab + vbTab + vbTab + vbTab + "</outerBoundaryIs>" + vbLf
            tmpStrPolygon = tmpStrPolygon + vbTab + vbTab + vbTab + "</Polygon>" + vbLf
            tmpStrPolygon = tmpStrPolygon + vbTab + vbTab + vbTab + "</Placemark>" + vbLf
            Put #1, , tmpStrPolygon
            tmpStrPolygon = ""
            Exit For
            
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
    Next
    
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


Sheets(1).Range("E15").Value = "����д��KML�ļ���������"

If Dir(outFile) <> "" Then Kill (outFile)
Call FileZM(tmpFile, "GB2312", outFile, "utf-8")
If Dir(tmpFile) <> "" Then Kill (tmpFile)

Sheets(1).Range("E15").Value = "д��KML�ļ����"

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
