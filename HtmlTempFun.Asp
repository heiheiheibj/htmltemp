<%

Dim strHtmlTempHtmlContent
Dim objHtmlTempfs 
Dim objHtmlTempf 
Dim TempSub    '区块内容
Dim TempSubStr '区块替换后的内容
Dim SubName    '区块名称
Dim StartSub      '区块开始处
Dim EndSub        '区块结束处
Dim TempSubStrAll  '区块替换完后的内容
Dim ErrMsg         '错误提示信息
Dim ModuleFilePath '模板文件路径
Dim HtmlTitle      'Html文件标题内容
Dim Zt



'/************************** 模板文件替换程序***********************/
'使用方法:
'Set b2b = Server.CreateObject("etb2c114.modulefun")
'b2b.OpenHtmlFile (FileName)            打开模板文件
'b2b.ReplaceVar "title", "嘿嘿"         用ReplaceVar方法来替换全局变量
'b2b.ReplaceVar "username", "hater"
'b2b.ListSub ("testmk")                 用ListSub(ListSubname)打开区块
'For s = 1 To 10
'    b2b.ReplaceSub "i", s              用ReplaceSub方法来替换区块中的变量
'    b2b.ReplaceSub "v", s
'    b2b.ReplaceSubAll                  用ReplaceSubAll方法来替换区块中的变量并累加
'Next
'b2b.ReplaceSubHtml                     用ReplaceSubHtml方法来将Html文件内容中的区块内容替换掉
'b2b.ListSub ("testmk1")
'    b2b.ReplaceSub "k", 1000
'    b2b.ReplaceSubAll
'b2b.ReplaceSubHtml
'b2b.ListSub ("testmk2")
'    b2b.ReplaceSub "k", 100
'    b2b.ReplaceSubAll
'b2b.ReplaceSubHtml
'b2b.PrintHtml                          用PrintHtml方法来输出


Sub OpenHtmlFile(FileName)
    Set objHtmlTempfs = Server.CreateObject("Scripting.FileSystemObject")
    If objHtmlTempfs.FileExists(FileName) Then        '打开html模板文件
        strHtmlTempHtmlContent = ""
        Set objHtmlTempf = objHtmlTempfs.opentextfile(FileName, 1, False)
        Do While Not objHtmlTempf.AtEndOfStream
            strHtmlTempHtmlContent = strHtmlTempHtmlContent + objHtmlTempf.readline() + Chr(13)
        Loop
        objHtmlTempf.Close
        'Response.Write strHtmlTempHtmlContent
   Else
        Response.Write FileName + "文件不存在！"
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




Sub ListSub(SubNameTmp)    '打开List块
SubName = SubNameTmp
Zt = 0
If SubName = "" Then
    Response.Write "LIST方法中区块名不能为空！"
Else
    StartSub = InStr(strHtmlTempHtmlContent, "<!--start " + SubName + " -->")
    EndSub = InStr(strHtmlTempHtmlContent, "<!--end " + SubName + " -->")
    If StartSub = 0 Or EndSub = 0 Or EndSub - StartSub < 0 Then
        Response.Write "区块" + SubName + "不存在或不完全！！" + CStr(StartSub) + CStr(EndSub)
    Else
        TempSub = Mid(strHtmlTempHtmlContent, StartSub, EndSub - StartSub)
    End If
End If

End Sub


Sub ReplaceSub(VarBl, ValueBl)    '将List块中的变量进行替换
    If Zt = 0 Then
        TempSubStr = TempSub    '当zt为0时tempsubstr值为List块中的数据
    End If
    If IsNull(TempSub) Then
        Response.Write "请先用list方法打开区块！"
    Else
        TempSubStr = Replace(TempSubStr, "%%" + VarBl + "%%", ValueBl)
        Zt = 1  '当zt为1时tempsubstr值为tempsubstr中的数据
    End If
End Sub

Sub ReplaceSubAll()  '将List块的内容进行累加
        TempSubStrAll = TempSubStrAll + TempSubStr
        Zt = 0
End Sub

Sub ReplaceSubHtml()   '将Html页面中的List块替换掉
    'If SubName = "" Or TempSubStrAll = "" Then
    If SubName = "" Then
        Response.Write "区块名或区块内容不能为空！！"
        Response.Write "<br>subname=" + SubName
        Response.Write "<br>tempsubstrall=" + TempSubStrAll + "<br>"
    Else
        If StartSub = 0 Or EndSub = 0 Or EndSub - StartSub < 0 Then
            Response.Write "区块" + SubName + "不存在或不完全！！"
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

Sub ReplaceSubNull() '将Html中的List块中的内容屏蔽不显示
            strHtmlTempHtmlContent = Left(strHtmlTempHtmlContent, StartSub - 1) + Right(strHtmlTempHtmlContent, Len(strHtmlTempHtmlContent) - EndSub + 1)
            strHtmlTempHtmlContent = Replace(strHtmlTempHtmlContent, "<!--start " + SubName + " -->", "")
            strHtmlTempHtmlContent = Replace(strHtmlTempHtmlContent, "<!--end " + SubName + " -->", "")
End Sub

Sub PrintHtml()
    Response.Write strHtmlTempHtmlContent
    strHtmlTempHtmlContent = ""
End Sub
%>