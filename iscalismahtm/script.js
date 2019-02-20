function isimCek(id) 
{ 
document.getElementById(id.toString()).innerHTML='';
// Code to connect the MS Access database using java Script
// " C:/my_db.mdb " is the MS Access database
var cn = new ActiveXObject("ADODB.Connection");
var strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = iscalisma/is.mdb"; 
var rs = new ActiveXObject("ADODB.Recordset");
// Assume there is a table in MS Access database with the name " user_master “, below is the query for authentication
var SQL = "select distinct(ad) from isi"; 
cn.Open(strConn); 
rs.Open(SQL, cn);
 while (!rs.eof){

document.getElementById(id.toString()).innerHTML+='<option>'+rs.fields(0).value+'</option>';
        rs.MoveNext();
    }

rs.Close(); 	
}
function firmaEkle() 
{ 
var isim=document.getElementById("isim");
isim=isim.options[isim.selectedIndex].text;
var tarih=document.getElementById("tarih").value;
if(tarih.toString().length<8)alert("Lütfen gecerli bir tarih giriniz!");
else{ 
// Code to connect the MS Access database using java Script
// " C:/my_db.mdb " is the MS Access database
var cn = new ActiveXObject("ADODB.Connection");
var strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = iscalismahtm/is.mdb"; 
var rs = new ActiveXObject("ADODB.Command");
// Assume there is a table in MS Access database with the name " user_master “, below is the query for authentication 
cn.Open(strConn); 
rs.ActiveConnection = cn;
rs.CommandText = "INSERT INTO isi(ad,bTarih)VALUES('"+isim.toString()+"','"+tarih.toString()+"')";
if(rs.Execute()){
var x=document.getElementById("ekleme");
var y=document.getElementById("isekleme");
 x.style.display="none";
 y.style.display="block";
 //rs.Close();
 rs = new ActiveXObject("ADODB.Recordset");
// Assume there is a table in MS Access database with the name " user_master “, below is the query for authentication
var SQL = "select max(iNo) from isi";  
rs.Open(SQL, cn);
if(!rs.eof)
document.getElementById("id").value=rs.fields(0).value;
alert("success!");
}
else alert("unsuccess!");	
rs.Close();
}
}


function calismaEkle() 
{ 

var calisma=document.getElementById("calisma");
calisma=calisma.options[calisma.selectedIndex].value;
var id=document.getElementById("id").value;
var adet=document.getElementById("adet").value;
if(adet<0)alert("Lütfen gecerli bir miktar giriniz!");
else{ 
// Code to connect the MS Access database using java Script
// " C:/my_db.mdb " is the MS Access database
var cn = new ActiveXObject("ADODB.Connection");
var strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = iscalismahtm/is.mdb"; 
var rs = new ActiveXObject("ADODB.Command");
// Assume there is a table in MS Access database with the name " user_master “, below is the query for authentication 
cn.Open(strConn); 
rs.ActiveConnection = cn;
rs.CommandText = "INSERT INTO miktar(iNo,cMiktar,cNo)VALUES("+id+","+adet+","+calisma+")";
if(rs.Execute()){

 //rs.Close();
 rs = new ActiveXObject("ADODB.Recordset");
// Assume there is a table in MS Access database with the name " user_master “, below is the query for authentication
var SQL = "select ID,(select cAd from calisma where cNo=miktar.cNo) as a,(select birim from calisma where cNo=miktar.cNo) as b,(select bfiyat from calisma where cNo=miktar.cNo) as c,cMiktar from miktar where iNo in( select max(iNo) from isi)";  
rs.Open(SQL, cn);
document.getElementById("table").innerHTML=' ';
document.getElementById("table").innerHTML+=' <tr><th>Sil</th><th>Calisma</th><th>Birim</th><th>Fiyat</th><th>Miktar</th></tr> ';
var total=0;
while(!rs.eof){
document.getElementById("table").innerHTML+='<tr><td><input type=button value='+rs.fields(0).value+' onclick=calismaSil(this.value)></td><td>'+rs.fields(1).value+'</td><td>'+rs.fields(2).value+'</td><td>'+rs.fields(3).value+'</td><td>'+rs.fields(4).value+'</td></tr>';
total=total+rs.fields(3).value*rs.fields(4).value;
rs.MoveNext();	
}
rs = new ActiveXObject("ADODB.Command");
// Assume there is a table in MS Access database with the name " user_master “, below is the query for authentication 
rs.ActiveConnection = cn;
rs.CommandText = "update isi set ucret="+total+" where iNo="+id+"";
if(rs.Execute()){
document.getElementById("toplam").innerHTML=total.toString();
alert("success!");
}
else alert("unsuccess!");	
}
}	
}	 
function calismaSil(id){
// Code to connect the MS Access database using java Script
// " C:/my_db.mdb " is the MS Access database
var iNo=document.getElementById("id").value;
var cn = new ActiveXObject("ADODB.Connection");
var strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = iscalismahtm/is.mdb"; 
var rs = new ActiveXObject("ADODB.Command");
// Assume there is a table in MS Access database with the name " user_master “, below is the query for authentication 
cn.Open(strConn); 
rs.ActiveConnection = cn;
rs.CommandText = "delete from miktar where ID="+id+"";
if(rs.Execute()){
 //rs.Close();
 rs = new ActiveXObject("ADODB.Recordset");
// Assume there is a table in MS Access database with the name " user_master “, below is the query for authentication
var SQL = "select ID,(select cAd from calisma where cNo=miktar.cNo) as a,(select birim from calisma where cNo=miktar.cNo) as b,(select bfiyat from calisma where cNo=miktar.cNo) as c,cMiktar from miktar where iNo in( select max(iNo) from isi)";  
rs.Open(SQL, cn);
document.getElementById("table").innerHTML=' ';
document.getElementById("table").innerHTML+=' <tr><th>Sil</th><th>Calisma</th><th>miktar</th><th>Birim</th></tr> ';
var total=0;
while(!rs.eof){
document.getElementById("table").innerHTML+='<tr><td><input type=button value='+rs.fields(0).value+' onclick=calismaSil(this.value)></td><td>'+rs.fields(1).value+'</td><td>'+rs.fields(2).value+'</td><td>'+rs.fields(3).value+'</td><td>'+rs.fields(4).value+'</td></tr>';
total=total+rs.fields(3).value*rs.fields(4).value;
rs.MoveNext();
}
rs = new ActiveXObject("ADODB.Command");
// Assume there is a table in MS Access database with the name " user_master “, below is the query for authentication 
rs.ActiveConnection = cn;
rs.CommandText = "update isi set ucret="+total+" where iNo="+iNo+"";
if(rs.Execute()){
document.getElementById("toplam").innerHTML=total.toString();
alert("success!");
}
else alert("unsuccess!");	

}
}
function yeniEkle(){
var string=document.getElementById("yeni").value;
 document.getElementById("isim").innerHTML+='<option>'+string.toString()+'</option>';
}
function listele(ids){
var strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = iscalismahtm/is.mdb"; 
var cn = new ActiveXObject("ADODB.Connection");
rs = new ActiveXObject("ADODB.Recordset");
cn.Open(strConn); 
var SQL = "select m.cMiktar as mik,c.cAd as ad,c.birim as birim,c.bfiyat as fiyat,(select bTarih from isi where iNo=m.iNo) " +
						"from calisma c,miktar m " +
						"where m.cNo=c.cNo and m.iNo="+ids+""; 
rs.Open(SQL, cn);
while(!rs.eof){
		var d=new Date(rs.fields(4).value);
		var d=d.toLocaleDateString();
document.getElementById("table").innerHTML+='<tr><td>'+d+'</td><td>'+rs.fields(1).value+'</td><td>'+rs.fields(0).value+'</td><td>'+rs.fields(2).value+'</td><td>'+rs.fields(3).value+'</td><td>'+rs.fields(3).value*rs.fields(0).value+' TL</td></tr>';
//ucret_real=ucret_real+rs.fields(0).value*rs.fields(3).value;
rs.MoveNext();
}
rs.Close();
rs = new ActiveXObject("ADODB.Recordset");
SQL = "select *from isi where iNo="+ids+""; 

rs.Open(SQL, cn);
while(!rs.eof){
		var d=new Date(rs.fields(2).value);
		var d=d.toLocaleDateString();
		document.getElementById("table").innerHTML+='<tr><td>'+d+'</td><td></td><td></td><td></td><td></td><th>+'+rs.fields(3)+'</th></tr>';
		ucret=ucret+rs.fields(3);
		rs.MoveNext();
    }
rs.Close(); 
rs = new ActiveXObject("ADODB.Recordset");
SQL = "select *from tahsil where iNo="+ids+"";
rs.Open(SQL, cn);
while(!rs.eof){
		var d=new Date(rs.fields(2).value);
		var d=d.toLocaleDateString();
		document.getElementById("table").innerHTML+='<tr><td>'+d+'</td><td></td><td></td><td></td><td></td><th>-'+rs.fields(3)+'</th></tr>';
		ucret=ucret-rs.fields(3);
		rs.MoveNext();
    }
rs.Close(); 
}
	function isFiltrele()
	{
	ucret=0;
	ucret_real=0;
	document.getElementById("is").innerHTML='<option value=0>EMPTY</option>';
	var count=0;
	var a=0;
	var b=0;
	var lp=new Array();
	
	var cn = new ActiveXObject("ADODB.Connection");
var strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = iscalismahtm/is.mdb"; 
	var sel=document.getElementById("firma");
	sel=sel.options[sel.selectedIndex].text;
	var	rs = new ActiveXObject("ADODB.Recordset");
SQL = "select iNo from isi where ad='"+sel+"'"; 
cn.Open(strConn); 
rs.Open(SQL, cn);
 while (!rs.eof){
		lp[count++]=rs.fields(0).value;
//document.getElementById("is").innerHTML+='<option value='+rs.fields(0).value+'>'+rs.fields(0).value+'</option>';
        rs.MoveNext();
    }
//document.getElementById("is").innerHTML+='<option value=0>empty</option>';
rs.Close(); 

document.getElementById("table").innerHTML=' <tr><th>Numara</th><th>Calisma</th><th>Miktar</th><th>Birim</th><th>Birim Fiyat</th><th>Tutar</th></tr> ';
	for(var j=0;j<count;j++)
	{
			a=0;
			b=0;
			rs = new ActiveXObject("ADODB.Recordset");
SQL = "select ucret from isi where iNo="+lp[j]+""; 
rs.Open(SQL, cn);
if(!rs.eof){
		a=rs.fields(0).value;
    }
rs.Close(); 
			rs = new ActiveXObject("ADODB.Recordset");
SQL = "select sum(kBakiye) from tahsil where iNo="+lp[j]+""; 
rs.Open(SQL, cn);
if(!rs.eof){
		b=rs.fields(0).value;
    }
rs.Close(); 
if(a>b){
document.getElementById("is").innerHTML+='<option value='+lp[j]+'>'+lp[j]+'</option>';
	listele(lp[j]);
}


}
document.getElementById("table").innerHTML+='<tr><td></td><td></td><td></td><td></td><td>Guncel Borcunuz:</td><th>'+ucret.toString()+' TL</th></tr>';	
//document.getElementById("odeme").innerHTML+='<tr><td></td><td></td><th>'+ucret.toString()+' TL</th></tr>';

	}

	function filter_id(id){
	/*alert(id.toString());
	var count=0;
	var a=0;
	var b=0;
	var row=new Array();*/
	ucret=0;
	ucret_real=0;
	document.getElementById("table").innerHTML=' ';
document.getElementById("table").innerHTML+=' <tr><th>Tarih</th><th>Calisma</th><th>Miktar</th><th>Birim</th><th>Birim Fiyat</th><th>Ara Toplam</th></tr> ';
//document.getElementById("odeme").innerHTML="<tr><th>numara</th><th>tarih</th><th>ucret</th></tr>";
	var cn = new ActiveXObject("ADODB.Connection");
var strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = iscalismahtm/is.mdb";
 cn.Open(strConn); 
listele(id);
document.getElementById("table").innerHTML+='<tr><td></td><td></td><td></td><td></td><td></td><th>'+ucret.toString()+' TL</th></tr>';
//document.getElementById("odeme").innerHTML+='<tr><td></td><td></td><th>'+ucret.toString()+' TL</th></tr>';
	}           
function tahsilEt(){
	var sel=document.getElementById("is");
	sel=sel.options[sel.selectedIndex].text;
	var ucret=document.getElementById("ucret").value;
	var tarih=document.getElementById("tarih").value;
	var cn = new ActiveXObject("ADODB.Connection");
	var strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = iscalismahtm/is.mdb"; 
	var rs = new ActiveXObject("ADODB.Command");
// Assume there is a table in MS Access database with the name " user_master “, below is the query for authentication 
	cn.Open(strConn); 
	rs.ActiveConnection = cn;
		rs.CommandText = "insert into tahsil(iNo,aTarih,kBakiye) values ("+sel.toString()+",'"+tarih.toString()+"',"+ucret.toString()+")";
		if(rs.Execute()){
		alert("Tahsil edildi! "+ucret+" TL");
		document.getElementById("table").innerHTML+='<tr><td>'+tarih.toString()+'</td><td></td><td></td><td></td><td></td><th>-'+ucret.toString()+'</th></tr>';
		}
}