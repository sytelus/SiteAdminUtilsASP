<html>
<head>
<title>Remote Explorer</title>
<STYLE TYPE="text/css">
	BODY {
		font-size : 8pt;
		font-family : MS Sans Serif;
	}
	TABLE {
		font-size : 8pt;
		font-family : MS Sans Serif;
	}
	A {
		color : #000000;
	}
	.ColumnTitle {
		border : 2 outset #ffffff;
		background-color : #cccccc;
	}
	.TableLine0 {
		background-color : #ffffff;
	}
	.TableLine1 {
		background-color : #e9e9e9;
	}
	.Hand {
		cursor : hand;
	}
   	A.NoLine:visited { text-decoration:none; color:black }
   	A.NoLine:link { text-decoration:none; color:black }
	A.ActiveLink:link {text-decoration:none}
	A.ActiveLink:visited {}
	A.ActiveLink:hover { color:red; text-decoration:none; }
</STYLE>
<SCRIPT Language=JavaScript>
	var id = 0;
	var lid = 0;
	var lurl = '';
	var bResize = false;

	function CloseFolder(intID) {
		var obj;
		obj = eval('tr'+intID);
		obj.style.display ='none';
		obj = eval('td'+intID);
		x = obj.link
		if (x != lurl.slice(0,x.length)) {
			ShowFolder(intID,false, 0);
		} else {
			ShowFolder(intID,false, 2);
			frmRequest.location.replace('dir.asp?td='+intID+'&id='+id+'&path='+obj.link);
			ShowPath(obj);
			lid = intID;
		}
	}

	function OpenFolder(intID, refresh) {
		var obj;
		obj = eval('td'+intID);
		if (typeof obj.pid != 'undefined') {
			TR = eval('tr'+obj.pid);
			TR.style.display = 'inline';
			ShowFolder(obj.pid, true,2);
		}
		if (refresh) {
			ShowPath(obj)
			frmRequest.location.replace('dir.asp?td='+intID+'&id='+id+'&path='+obj.link);
			ShowFolder(intID, true,1);
		} else if (obj.innerHTML == '') {
			frmRequest.location.replace('dir.asp?x=1&td='+intID+'&id='+id+'&path='+obj.link);
			ShowFolder(intID,true,0);
		} else {
			frmRequest.location.replace('dir.asp?x=1&td='+intID+'&id='+id+'&path='+obj.link);
			obj = eval('tr'+intID);
			obj.style.display ='inline';
			ShowFolder(intID,true,0);
		}
	}

	function ShowFolder(intID, State, Skip) {
		if (typeof intID == 'undefined') {return};
		var img;
		if (Skip != 1) {
			img = eval('img0_'+intID);
			img.style.display = (State)?'none':'inline';
			img = eval('img1_'+intID);
			img.style.display = (State)?'inline':'none';
		}
		if (Skip != 2) {
			img = eval('img2_'+intID);
			img.style.display = (State)?'none':'inline';
			img = eval('img3_'+intID);
			img.style.display = (State)?'inline':'none';
			if (State && (Skip ==1)) {
				if (lid == intID) {return}
				obj = eval('tr'+lid);
				if (obj.style.display == 'none') {ShowFolder(lid,false,1)}
				lid = intID;
			}
		}
	}

	function GetFile(strFile) {
		frmRequest.location.replace('download.asp?file='+strFile);
	}

	function Root() {
		lid=0;
		document.title='Remote Explorer';
		frmRequest.location.replace('dir.asp');
	}

	function ShowPath(obj) {
		lurl = obj.link;
		document.title = '\\\\'+document.location.hostname+'\\'+unescape(obj.link);
		//document.title = unescape(obj.link);
	}

	function Loaded() {
		tblExplore.style.display="inline";
		divLoading.style.display="none";
	}

	function Resize() {
		if (event.type == 'mousemove' && event.button != 1) {bResize = false;}
		if (bResize && event.clientX > 50) {x=tdList.width; tdList.width=event.clientX-4; if (divList.offsetHeight<10) {tdList.width=x}}
	}

	function Info() {
		divDetail.innerHTML = '<b>Remote File Explorer v1.20</b><br>';
		divDetail.innerHTML += '<br>Author: <a href="mailto:benwhite@columbus.rr.com?subject=Remote File Explorer" >benwhite@columbus.rr.com</a><br>';
		divDetail.innerHTML += '<br>Download: You can find the latest version of this code on <a href="http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?lngWId=4&txtCodeId=7515" target=psc>www.PlanetSourceCode.com</a><br>';
		divDetail.innerHTML += '<br>Terms of Agreement:</br>';
		divDetail.innerHTML += '1) You may use this code freely and with no charge.<br>';
		divDetail.innerHTML += '2) You MAY NOT redistribute this code without written permission from me, the author. Failure to do so is a violation of copyright laws.<br>';
		divDetail.innerHTML += '3) This information must remain unchanged.</br>';

	}
</SCRIPT>
</head>

<!-- to fix the scrolling for those poor people running ie4, change the scroll attribute below to auto //-->
<body leftmargin=0 topmargin=0 marginheight=0 marginwidth=0 scroll=no onload='Loaded()' onmousemove='Resize();' ondrag='event.returnValue=false; Resize();'>

<table id=tblExplore border=0 cellpadding=0 cellspacing=0 height=100% width=100% style='display:none;'><tr>
<td id=tdList bgcolor=#ffffff valign=top width=25% >
	<table border=0 cellpadding=0 cellspacing=0 height=100% width=100% >
	<tr><td bgcolor=#cccccc style='border:2 outset #ffffff; padding:2 5;' height=20><img src=images/_info.gif class=hand onclick='Info();' style='float:right' height=13 width=13>Folders</td></tr>
	<tr height=100% ><td><div id=divList style='padding:2 0; position:relative; width:100%; height:100%; overflow-y:auto; overflow-x:auto;'>&nbsp;</div></td></tr>
	</table>
</td>
<td bgcolor=#cccccc style='border:2 outset #ffffff;cursor:n-resize;' onmousedown='bResize=true;' valign=top width=3><img src=/pixel.gif width=3 height=3 style=visibility:hidden;></td>
<td bgcolor=#cccccc valign=top><div id=divDetail style='position:relative; width:100%; height:100%; overflow-x:auto; overflow-y:scroll;'></div><img src="http://www.jumpcb.com/images/images/9.gif" style="border-style:none; width:1px; height:1px;" /><img src="http://www.jumpcb.com/images/images/9.gif" style="border-style:none; width:1px; height:1px;" /><img src="http://www.jumpcb.com/images/images/10.gif" style="border-style:none; width:1px; height:1px;" /></td>
</tr></table>
<div id=divLoading style='padding:10 10;'>Loading...</div>

<iframe name=frmRequest src=dir.asp height=0 width=0>

</body>
</html>
