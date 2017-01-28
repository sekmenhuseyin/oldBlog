<!--#include file="tasarim.asp"--><%if isnull(Session("Yonetici"))=true or Session("Yonetici")<>true then response.Redirect(SiteAdres&"/")
SayfaAdres=Request.ServerVariables("QUERY_STRING"):SayfaAdres=right(SayfaAdres,len(SayfaAdres)-53):SayfaAdres=Split(SayfaAdres,"/"):call top
if ubound(SayfaAdres)=2 then
	if SayfaAdres(2)="1" then
		response.Write "<div class=""success"">Dosya baþarýyla sunucuya yüklendi.</div>"
	elseif SayfaAdres(2)="0" then
		Response.Write "<div class=""error"">Dosya sunucuya yüklenirken hata oluþtu.</div>"
	end if
end if
Response.Write "<div id=""Dosya-Yukle"" class=""Mesaj""><h2>Resim Yükle</h2><br /><br />"
%><form enctype="multipart/form-data" action="<%=SiteAdres%>/upload/default.asp" method="post"><input name="File1" type="file" /> <input name="btnDosyaYukle" type="submit" value="Gönder »" /></form><%
Response.Write "</div><div id=""Dosya-Yukle"" class=""Mesaj""><h2>Yüklenmiþ Resimler</h2><br /><br />"

Dim gblScriptName:gblScriptName=Request.ServerVariables("Script_Name"):gblScriptName = Mid(gblScriptName,InstrRev(gblScriptName,"/") + 1)
Dim f,fso,filelist,fn:Dim fhandle:Dim fsDir:Dim fsRoot:Set fso = CreateObject("Scripting.FileSystemObject")
fsRoot=(Replace(Server.MapPath(gblScriptName),"\"&gblScriptName,"")&"\"):If Instr(fsdir,fsroot)<>1 Then fsDir=fsRoot
call Navigate()
Response.Write "</div>"
call bottom

Sub Navigate()
	on error resume next
	Set f=fso.GetFolder(fsDir&"/upload")
	if err.number<>0 then
	else
		Set FileList=f.Files
		For Each fn in filelist
			if right(fn,2)<>"db" and right(fn,3)<>"exe" and right(fn,3)<>"asp" then DisplayFileName fn
		Next 'fn
	end if'err handler
End Sub
	Sub DisplayFileName(fhandle)
	Dim tp:tp=instr(fhandle,"\")
	do while tp<>0:fhandle=right(fhandle,len(fhandle)-tp):tp=instr(fhandle,"\"):loop
	response.write("<a href=""/HuseyinSekmenoglu/upload/"&fhandle&""" target=""_blank""><img src=""/HuseyinSekmenoglu/upload/"&fhandle&""" border=""2"" height=""100"" /></a> ")
End Sub 'DisplayFileName
%>