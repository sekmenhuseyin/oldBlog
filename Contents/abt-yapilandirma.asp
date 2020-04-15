<%DB=Server.Mappath("\Contents\DataSource\data2KM37FN.mdb"):Dim SayfaAdres:Dim SiteAdres:session.lcid=1055
Set ObjCon=Server.CreateObject ("ADODB.Connection"):ObjCon.Open "Provider=Microsoft.JET.OLEDB.4.0; Data Source="&DB&";"
Set SiteAyar=ObjCon.Execute("Select * From SiteAyar")
AdSoyad=SiteAyar("AdSoyad")
EPosta=SiteAyar("EPosta")
SiteAdi=SiteAyar("SiteAdi")
SiteBaslik=SiteAyar("SiteBaslik")
SiteAdres=SiteAyar("SiteAdres")
SiteTelifHakki=SiteAyar("SiteTelifHakki")
SiteMetaDes=SiteAyar("SiteMetaDes")
SiteMetaKey=SiteAyar("SiteMetaKey")
SiteSayfalama=SiteAyar("SiteSayfalama")
ZamanFark=SiteAyar("ZamanFark")
SiteTarih=DateAdd("h" ,ZamanFark,Now())
SiteAyar.close:Set SiteAyar=Nothing%>