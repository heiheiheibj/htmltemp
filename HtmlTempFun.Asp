<%

Dim strHtmlTempHtmlContent
Dim objHtmlTempfs 
Dim objHtmlTempf 
Dim TempSub    '��������
Dim TempSubStr '�����滻�������
Dim SubName    '��������
Dim StartSub      '���鿪ʼ��
Dim EndSub        '���������
Dim TempSubStrAll  '�����滻��������
Dim ErrMsg         '������ʾ��Ϣ
Dim ModuleFilePath 'ģ���ļ�·��
Dim HtmlTitle      'Html�ļ���������
Dim Zt



'/************************** ģ���ļ��滻����***********************/
'ʹ�÷���:
'Set b2b = Server.CreateObject("etb2c114.modulefun")
'b2b.OpenHtmlFile (FileName)            ��ģ���ļ�
'b2b.ReplaceVar "title", "�ٺ�"         ��ReplaceVar�������滻ȫ�ֱ���
'b2b.ReplaceVar "username", "hater"
'b2b.ListSub ("testmk")                 ��ListSub(ListSubname)������
'For s = 1 To 10
'    b2b.ReplaceSub "i", s              ��ReplaceSub�������滻�����еı���
'    b2b.ReplaceSub "v", s
'    b2b.ReplaceSubAll                  ��ReplaceSubAll�������滻�����еı������ۼ�
'Next
'b2b.ReplaceSubHtml                     ��ReplaceSubHtml��������Html�ļ������е����������滻��
'b2b.ListSub ("testmk1")
'    b2b.ReplaceSub "k", 1000
'    b2b.ReplaceSubAll
'b2b.ReplaceSubHtml
'b2b.ListSub ("testmk2")
'    b2b.ReplaceSub "k", 100
'    b2b.ReplaceSubAll
'b2b.ReplaceSubHtml
'b2b.PrintHtml                          ��PrintHtml���������


Sub OpenHtmlFile(FileName)
    Set objHtmlTempfs = Server.CreateObject("Scripting.FileSystemObject")
    If objHtmlTempfs.FileExists(FileName) Then        '��htmlģ���ļ�
        strHtmlTempHtmlContent = ""
        Set objHtmlTempf = objHtmlTempfs.opentextfile(FileName, 1, False)
        Do While Not objHtmlTempf.AtEndOfStream
            strHtmlTempHtmlContent = strHtmlTempHtmlContent + objHtmlTempf.readline() + Chr(13)
        Loop
        objHtmlTempf.Close
        'Response.Write strHtmlTempHtmlContent
   Else
        Response.Write FileName + "�ļ������ڣ�"
   End If
   CloseFile
End Sub

Sub ReplaceVar(VarBl, ValueBl)
    strHtmlTempHtmlContent = Replace(strHtmlTempHtmlContent, "%%" + VarBl + "%%", ValueBl)
End Sub

Sub CloseFile()    
    Set objHtmlTempf = Nothing
    Set objHtmlTempfs = Nothing
End Sub




Sub ListSub(SubNameTmp)    '��List��
SubName = SubNameTmp
Zt = 0
If SubName = "" Then
    Response.Write "LIST����������������Ϊ�գ�"
Else
    StartSub = InStr(strHtmlTempHtmlContent, "<!--start " + SubName + " -->")
    EndSub = InStr(strHtmlTempHtmlContent, "<!--end " + SubName + " -->")
    If StartSub = 0 Or EndSub = 0 Or EndSub - StartSub < 0 Then
        Response.Write "����" + SubName + "�����ڻ���ȫ����" + CStr(StartSub) + CStr(EndSub)
    Else
        TempSub = Mid(strHtmlTempHtmlContent, StartSub, EndSub - StartSub)
    End If
End If

End Sub


Sub ReplaceSub(VarBl, ValueBl)    '��List���еı��������滻
    If Zt = 0 Then
        TempSubStr = TempSub    '��ztΪ0ʱtempsubstrֵΪList���е�����
    End If
    If IsNull(TempSub) Then
        Response.Write "������list���������飡"
    Else
        TempSubStr = Replace(TempSubStr, "%%" + VarBl + "%%", ValueBl)
        Zt = 1  '��ztΪ1ʱtempsubstrֵΪtempsubstr�е�����
    End If
End Sub

Sub ReplaceSubAll()  '��List������ݽ����ۼ�
        TempSubStrAll = TempSubStrAll + TempSubStr
        Zt = 0
End Sub

Sub ReplaceSubHtml()   '��Htmlҳ���е�List���滻��
    'If SubName = "" Or TempSubStrAll = "" Then
    If SubName = "" Then
        Response.Write "���������������ݲ���Ϊ�գ���"
        Response.Write "<br>subname=" + SubName
        Response.Write "<br>tempsubstrall=" + TempSubStrAll + "<br>"
    Else
        If StartSub = 0 Or EndSub = 0 Or EndSub - StartSub < 0 Then
            Response.Write "����" + SubName + "�����ڻ���ȫ����"
        Else
            TempSub = Mid(strHtmlTempHtmlContent, StartSub, EndSub - StartSub)
            strHtmlTempHtmlContent = Left(strHtmlTempHtmlContent, StartSub - 1) + TempSubStrAll + Right(strHtmlTempHtmlContent, Len(strHtmlTempHtmlContent) - EndSub + 1)
            strHtmlTempHtmlContent = Replace(strHtmlTempHtmlContent, "<!--start " + SubName + " -->", "")
            strHtmlTempHtmlContent = Replace(strHtmlTempHtmlContent, "<!--end " + SubName + " -->", "")
        End If
    End If
    TempSubStr = ""
    TempSubStrAll = ""
    SubName = ""
    Zt = 0
End Sub

Sub ReplaceSubNull() '��Html�е�List���е��������β���ʾ
            strHtmlTempHtmlContent = Left(strHtmlTempHtmlContent, StartSub - 1) + Right(strHtmlTempHtmlContent, Len(strHtmlTempHtmlContent) - EndSub + 1)
            strHtmlTempHtmlContent = Replace(strHtmlTempHtmlContent, "<!--start " + SubName + " -->", "")
            strHtmlTempHtmlContent = Replace(strHtmlTempHtmlContent, "<!--end " + SubName + " -->", "")
End Sub

Sub PrintHtml()
    Response.Write strHtmlTempHtmlContent
    strHtmlTempHtmlContent = ""
End Sub
%>