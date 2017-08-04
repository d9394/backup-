<%
  dim d,date1
  dd=trim(request("d"))
  if dd="" then
	d=date()
  else
	d=cdate(left(dd,4)&"-"&mid(dd,5,2)&"-"&right(dd,2))
  end if
  
  date1=year(d) & iif (month(d)<10,"0","") & month(d) & iif(day(d)<10 ,"0","") & day(d)
  date2=year(d) & "-" & month(d) & "-" & day(d)
  date3=iif (month(d)<10,"0","") & month(d) & iif(day(d)<10 ,"0","") & day(d)
  date4=year(d) & "-" & iif (month(d)<10,"0","") & month(d) & "-" & iif(day(d)<10 ,"0","") & day(d)
  date5=year(date()) & iif (month(date())<10,"0","") & month(date()) & iif(day(date())<10 ,"0","") & day(date())
  
  filepath=server.mappath("./")
  
  dim source(100,7)

'  reqfunc=lcase(trim(request("reqfunc")))
  cfgfile="./backup.cfg"

  Dim FSO
  Set FSO = Server.CreateObject("Scripting.FileSystemObject")

  n=0
  if FSO.FileExists(Server.MapPath(cfgfile) ) then
    Set cfgFileObj = fso.opentextfile(server.mappath(cfgfile),1,true)
    While not cfgFileObj.AtEndOfStream
      line=trim(cfgFileObj.ReadLine)
'	  response.write line & "<br>"
      if len(line)>0 then
        linetext=split(line,",")
        if left(trim(linetext(0)),1) <> "#" then
		  if lcase(trim(linetext(0)))="file" then
            linetext(2)=replace(linetext(2),"%YYYYMMDD%",date1)
            linetext(2)=replace(linetext(2),"%YYYY-MM-DD%",date4)
            linetext(2)=replace(linetext(2),"%MMDD%",date3)
            linetext(2)=replace(linetext(2),"%YYYY-M-D%",date2)
 
            linetext(3)=replace(linetext(3),"%YYYYMMDD%",date1)
            linetext(3)=replace(linetext(3),"%YYYY-MM-DD%",date4)
            linetext(3)=replace(linetext(3),"%MMDD%",date3)
            linetext(3)=replace(linetext(3),"%YYYY-M-D%",date2)
		  end if
          source(n,0)=trim(linetext(1))  '说明，系统
          source(n,1)=trim(linetext(2))  '路径，SID
          source(n,2)=trim(linetext(3))  '文件名，用户名
          source(n,3)=trim(linetext(4))  'null，密码
          source(n,4)=trim(linetext(5))  '压缩前缀
		  source(n,5)=trim(getTaskNextTime(source(n,4)))  '传递备份提交情况
		  source(n,6)=trim(linetext(6))  '全备份延迟开始
		  source(n,7)=lcase(trim(linetext(0)))  'file，ocl
          n=n+1
        end if
      end if
    Wend
    cfgFileObj.Close

  end if
  target="\DataBackup\" & date1 & "\"
  req=trim(request("a"))
  action=trim(request("c"))
  if not FSO.FolderExists(Server.MapPath( target ))then
     FSO.CreateFolder(Server.MapPath( target ))
  end if

  if len(req)>0 and action="backup" then
    if req="all" then
      for yy=0 to n-1
		time2=formatdatetime(DateAdd("n",1+source(yy,6),time()),4)
		if lcase(source(yy,7)) ="ocl" then
			aaa=DATA_BACKUP(source(yy,2)&"/"&source(yy,3)&"@"&source(yy,1), Server.MapPath(target), source(yy,4),date1,time2,"ocl")
		else
			if lcase(source(yy,7))= "file" then
				aaa=DATA_BACKUP(source(yy,1)&"\"&source(yy,2),Server.MapPath(target),source(yy,4),date1,time2,"file")
			else
			end if
		end if
      next
    else
		time2=formatdatetime(DateAdd("n",1,time()),4)
		if lcase(source(req,7)) ="ocl" then
			aaa=DATA_BACKUP(source(req,2)&"/"&source(req,3)&"@"&source(req,1), Server.MapPath(target), source(req,4),date1,time2,"ocl")
		else
			if lcase(source(req,7))= "file" then
				aaa=DATA_BACKUP(source(req,1)&"\"&source(req,2),Server.MapPath(target),source(req,4),date1,time2,"file")
			else
			end if
		end if
    end if
  end if

  url = Request.ServerVariables("SCRIPT_NAME")
  urlParts = Split(url,"/")
  pageName = urlParts(UBound(urlParts))
  
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--<meta http-equiv="refresh" content="300; url=<%=pageName%>?reqfunc=<%=reqfunc%>"/>-->
<title>XXXX公司BACKUP工具</title>
<style type="text/css">
<!--
body {
	color:#000000;
	background-color:#B3B3B3;
	margin:0;
}

#container {
	margin-left:auto;
	margin-right:auto;
	text-align:center;
	}

a img {
	border:none;
}

-->
</style>
</head>
<body>
<div id="container">
<center>
<table>
	<tr><td align=center>
	<table border=1>
	  <tr>
	    <form method="get" action="<%=pageName%>" name="inputdate">
		<td colspan=3>备份日期是：<input type="text" name="d" value="<%=date1%>"><input type="submit" value="改日期">&nbsp;<input type="button" value="今天日期" onclick="todayclick('<%=date5%>')">&nbsp;<%=formatdatetime(now,4)%>&nbsp;<%=formatdatetime(DateAdd("n",1,time()),4)%>, 
	<%
		if reqfunc="qs" then
		  response.write "<b>清算</b>"
		else
		  
		end if
	%>
		任务数: <%=n%>, a=<%=req%>, c=<%=action%></td>
		</form>
		<td align=center><a href="http://<%=request.ServerVariables("LOCAL_ADDR")%>/">回主页</a></td>
	  </tr>
	  <tr>
		<td align=center>系统</td><td align=center>目标文件</td><td align=center>操作</td><td>下次计划-上次执行</td>
	  </tr>
	<%
		for xx=0 to n-1
		  response.write "<tr>" & chr(13)
		  target_file= Server.MapPath( target & source(xx,4) & "_" & date1 & ".rar")

		  target_status=FSO.FileExists( target_file )

		  response.write "<td>" & source(xx,0) & "</td>"
		  response.write "<td align=left>" & iif(target_status,"<font color='blue'>","<font color='red'>") & target_file &  "</font></td>" & chr(13)
		  response.write "<td>"
		  if ( not target_status) then 
			response.write "<a href='" & pageName & "?a=" & xx & "&c=backup&d=" &date1 & "'>备份</a></td>" & chr(13)
		  else
			  response.write "<a href='" & pageName & "?a=" & xx & "&c=backup&d=" & date1 & "'>重新备份</a></td>" & chr(13)
		  end if
		  response.write "<td>" & source(xx,5) & "</td>"
		  response.write "</tr>" & chr(13)
		next
		response.write "<tr><td align=center><a href='" & pageName & "?d=" & date1 &"'>刷新</a></td>" & chr(13)
		response.write "<td colspan=2 align=right><a href='" & pageName & "?a=all&c=backup&d="& date1 &"'>全部备份</a></td>" & chr(13)
		response.write "<td></td></tr>" & chr(13)
	%>
	</table>
	</td></tr>
	<tr><td align=center>备份日志</td></tr>
	<tr><td valign=top>
		<%
		log_file=Server.MapPath( target & "backup_" & date1 & "_log.htm")
		if not FSO.FileExists(log_file) then
			set fd=FSO.createtextfile(log_file,true)
'			fd.writeline "<metahttp-equiv="&chr(34)&"refresh"&chr(34)&" content="&chr(34)&"300; url="&pageName&"?reqfunc="&reqfunc&chr(34)&"/>"
			fd.writeline now()&"<br/>"
			set fd=nothing
		end if
		response.write "<iframe src='"& target & "backup_" & date1 & "_log.htm'  width='900' height='400' id='ghrzFrame' frameborder='1' scrolling='auto' name='ghrzFrame' >" & chr(13)
		response.write "<p>Your browser does not support iframes.</p>" & chr(13)
		response.write "</iframe>" & chr(13)
		%>
	</td></tr>
</table>
</center>
</div>
</body>
<script>
function todayclick(datetodayclick)
{ 
 this.inputdate.d.value=datetodayclick;
 this.inputdate.submit();
} 
</script> 
</html>
<%
Function IIf(bExp1, sVal1, sVal2)
    If (bExp1) Then
        IIf = sVal1
    Else
        IIf = sVal2
    End If
End Function
Function DATA_BACKUP(uid,tag_path,tag_system,tag_date,time1,func1)
  Dim FSO1
  Set FSO1 = Server.CreateObject("Scripting.FileSystemObject")
  
  if not FSO1.FileExists( tag_path&"\"&tag_system&"_"&tag_date) then
'    tag_path=mid(tag,1,instrrev(tag,"\"))
   
    if not FSO.FolderExists(tag_path)then
       FSO.CreateFolder(tag_path)
    end if
	
	if func1="ocl" then
		command = "schtasks /create /tn T_" & tag_system & " /F /sc ONCE /st " & time1 & " /tr " & chr(34) & "%windir%\system32\cmd /c " & Server.MapPath("./backup.bat") & " " & uid & " " & tag_path & " " & tag_system & " " & tag_date & " " & func1 & chr(34)
	else
		if func1="file" then
			command = "schtasks /create /tn T_" & tag_system & " /F /sc ONCE /st " & time1 & " /tr " & chr(34) & "%windir%\system32\cmd /c " & Server.MapPath("./backup.bat") & " " & uid & "  " & tag_path & " " & tag_system & " " & tag_date & " " & func1 & chr(34)
		end if
	end if
'response.write command & "<br/>"
    Set WshShell = server.CreateObject("Wscript.Shell")
    Set IsSuccess = WshShell.exec ("%windir%\system32\cmd.exe")
	IsSuccess.StdIn.WriteLine command
''	IsSuccess.StdIn.WriteLine "yw-123"
	Issuccess.StdIn.WriteLine "exit"
	
    DATA_BACKUP=IsSuccess.stdout.readall()
    Set IsSuccess = Nothing
    Set WshShell = Nothing

  end if
  Set FSO1 = Nothing
End Function

Function getTaskNextTime(taskname1)
	Set WshShell = server.CreateObject("Wscript.Shell")
	Set IsSuccess = WshShell.exec ("%windir%\system32\cmd.exe /c chcp 437 | schtasks /query /fo csv /v /tn " & chr(34) & "T_" & taskname1 & chr(34))
	If left(IsSuccess.StdOut.ReadLine,5) <> "ERROR" then
		getLine = trim(IsSuccess.StdOut.ReadLine)
	else
		getLine =""
	end If
'response.write taskname1 & "," & getLine & "<BR/>"

	If getLine <> "" then
		getInfo = split(getLine,",")
		getTaskNextTime = getInfo(2) & " - " & getInfo(5)
	else
		getTaskNextTime = ""
	end if
	Set IsSuccess = Nothing
    Set WshShell = Nothing
End Function

%>
