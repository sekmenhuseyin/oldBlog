<!-- #Include File="Contents/abt-yapilandirma.asp" --><%SayfaAdres=Request.ServerVariables("QUERY_STRING")
If SayfaAdres<>"" then
  if len(SayfaAdres)<53 then response.redirect(SiteAdres)
	SayfaAdres=right(SayfaAdres,len(SayfaAdres)-53):SayfaAdres=Split(SayfaAdres,"/")
	If isnumeric(SayfaAdres(1))=true then
		Server.Execute("/HuseyinSekmenoglu/contents/index.asp")
	ElseIf SayfaAdres(1)="iletisim" then
		Server.Execute("/HuseyinSekmenoglu/contents/iletisim.asp")
	ElseIf SayfaAdres(1)="projelerim" or SayfaAdres(1)="CV" then
		Server.Execute("/HuseyinSekmenoglu/contents/xtra.asp")
	ElseIf SayfaAdres(1)="blog" then
		Server.Execute("/HuseyinSekmenoglu/contents/oku.asp")
	ElseIf SayfaAdres(1)="etiket" then
		Server.Execute("/HuseyinSekmenoglu/contents/etiket.asp")
	ElseIf SayfaAdres(1)="kategori" then
		Server.Execute("/HuseyinSekmenoglu/contents/kategoriler.asp")
	ElseIf SayfaAdres(1)="sayfa" then
		Server.Execute("/HuseyinSekmenoglu/contents/sayfa.asp")
	ElseIf SayfaAdres(1)="arsiv" then
		Server.Execute("/HuseyinSekmenoglu/contents/arsiv.asp")
	ElseIf SayfaAdres(1)="ara" then
		Server.Execute("/HuseyinSekmenoglu/contents/ara.asp")
	ElseIf SayfaAdres(1)="rss" then
		Server.Execute("/HuseyinSekmenoglu/contents/rss.asp")
	ElseIf SayfaAdres(1)="Yeni" or SayfaAdres(1)="Duzenle" or SayfaAdres(1)="TemizDuzenle" then'sadece admin
		Server.Execute("/HuseyinSekmenoglu/contents/yonetim_yeni.asp")
	ElseIf SayfaAdres(1)="TumYorumlar" then'sadece admin
		Server.Execute("/HuseyinSekmenoglu/contents/yonetim_yorum.asp")
	ElseIf SayfaAdres(1)="TumLinkler" then'sadece admin
		Server.Execute("/HuseyinSekmenoglu/contents/yonetim_link.asp")
	ElseIf SayfaAdres(1)="TumKategoriler" then'sadece admin
		Server.Execute("/HuseyinSekmenoglu/contents/yonetim_kategori.asp")
	ElseIf SayfaAdres(1)="TumEtiketler" then'sadece admin
		Server.Execute("/HuseyinSekmenoglu/contents/yonetim_etiket.asp")
	ElseIf SayfaAdres(1)="Ayarlar" then'sadece admin
		Server.Execute("/HuseyinSekmenoglu/contents/yonetim_ayar.asp")
	ElseIf SayfaAdres(1)="DosyaYukle" then'sadece admin
		Server.Execute("/HuseyinSekmenoglu/contents/yonetim_upload.asp")
	ElseIf SayfaAdres(1)="TumMesajlar" then'sadece admin
		Server.Execute("/HuseyinSekmenoglu/contents/yonetim_mesaj.asp")
	Else
		Server.Execute("/HuseyinSekmenoglu/contents/404.asp")
	End If
else
	Server.Execute("/HuseyinSekmenoglu/contents/index.asp")
end if%>