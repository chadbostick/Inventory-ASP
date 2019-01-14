<%option explicit

dim cnnSQLConnection
dim rsGetProductList, cmdGetProductList
dim rsGetSupplierList, cmdGetSupplierList

set rsGetProductList		= nothing
set cmdGetProductList		= nothing
set rsGetSupplierList		= nothing
set cmdGetSupplierList	= nothing

set cnnSQLConnection		= Server.CreateObject( "ADODB.Connection" )
set cmdGetProductList		= Server.CreateObject( "ADODB.Command" )
set cmdGetSupplierList	= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

'**********************************************************************************
cmdGetProductList.ActiveConnection = cnnSQLConnection
cmdGetProductList.CommandText = "{call spGetProductList}"

on error resume next
set rsGetProductList = cmdGetProductList.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if

on error goto 0
'**********************************************************************************
cmdGetSupplierList.ActiveConnection = cnnSQLConnection
cmdGetSupplierList.CommandText = "{call spGetSupplierList}"

on error resume next
set rsGetSupplierList = cmdGetSupplierList.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if

on error goto 0
'**********************************************************************************
%>
<HTML>
<HEAD>
<TITLE>Centerline - View / Edit Parts Cost</TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/default.css" TITLE="default_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/titles.css" TITLE="titles_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/forms.css" TITLE="forms_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/tables.css" TITLE="tables_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/pagelayout.css" TITLE="layout_css">
  <script type="text/javascript" src="js/common.js"></script>
  <script type="text/javascript" src="js/forms.js"></script>
  <script language="javascript">
	function PartSelected()
	{
		self.opener.document.all.txtProductLink.href							= "javascript:OpenProduct( '" + document.frmParts.selPart.options[document.frmParts.selPart.selectedIndex].value + "' );";
		self.opener.document.all.txtProductIdDiv.innerText				= document.frmParts.selPart.options[document.frmParts.selPart.selectedIndex].value;
		self.opener.document.frmParts.txtProductId.value					= document.frmParts.selPart.options[document.frmParts.selPart.selectedIndex].value;
		self.opener.document.frmParts.txtProductDescription.value	= document.frmParts.selPart.options[document.frmParts.selPart.selectedIndex].text;
		self.opener.document.frmParts.txtSupplierId.value					= document.frmParts.selSupplier.options[document.frmParts.selSupplier.selectedIndex].value;
		self.opener.document.frmParts.txtSupplier.value						= document.frmParts.selSupplier.options[document.frmParts.selSupplier.selectedIndex].text;
		self.opener.document.frmParts.txtCost.value								= document.frmParts.txtCost.value;
		self.opener.document.frmParts.txtMarkup.value							= document.frmParts.txtMarkup.value;
		window.close();
	}
  </script>
</head>
<body topmargin=0 leftmargin=0>

<!-- Page Contents -->
<table width=640 border=0 cellpadding=0 cellspacing=0>

<!--Top Title Bar -->
<tr><td class="pagetitle">  
 <table class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=3><IMG SRC="images/title_top.gif" width=100% height=6></td></tr>
   <tr><td class="sectiontitle" width=10%>Parts</td>
     <td align=left width=348><IMG SRC="images/title_right.gif" width=23 height=25></td></tr>
 </table>
</td></tr>
<!-- End Top Title Bar -->
  
 <!-- Parts Form -->
<form name="frmParts" action="javascript:PartSelected();" method="get">
<tr><td align=center>
	<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
		<tr><td class="generic_label">Part</td><td width=550>
		<select class="forminput" name="selPart" isRequired="true" fieldName="Part" size=1>
				<option value=""></option>
				<%if not( rsGetProductList is nothing ) then
   					if( rsGetProductList.State = adStateOpen ) then
   						while( not rsGetProductList.EOF )%><option value="<%=rsGetProductList( "ProductId" )%>"><%=rsGetProductList( "ProductDescription" )%></option>
   							<%rsGetProductList.MoveNext
   						wend
   					end if
   				end if%>
				</select></td>
		</tr>
		<tr><td class="generic_label">Supplier</td><td width=550>
		<select class="forminput" name="selSupplier" isRequired="true" fieldName="Supplier" size=1>
				<option value=""></option>
				<%if not( rsGetSupplierList is nothing ) then
   					if( rsGetSupplierList.State = adStateOpen ) then
   						while( not rsGetSupplierList.EOF )%>
   							<option value="<%=rsGetSupplierList( "SupplierId" )%>"><%=rsGetSupplierList( "SupplierName" )%></option>
   							<%rsGetSupplierList.MoveNext
   						wend
   					end if
   				end if%>
				</select></td>
		</tr>
		<tr><td class="label_1">Cost</td><td><input type="text" class="forminput" name="txtCost" isRequired="true" fieldName="Cost" value=""></td></tr>
		<tr><td class="label_1">Markup</td><td><input type="text" class="forminput" name="txtMarkup" isRequired="true" fieldName="Markup" value=""></td></tr>
	</table>
</td></tr> 

<!-- Form Bottom Rule -->
<tr><td height=12>
 <TABLE class="" width=100% cellpadding=0 cellspacing=0>
   <tr><td><IMG SRC="images/small_title_top.gif" width=100% height=3></td></tr>
 </TABLE>
</td></tr>
<!-- End Form Bottom Rule -->

<!-- Form Buttons -->
<tr><td align=center>
   	<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
	    <tr><td><table cellspacing=0 cellpadding=0 width=128 align=right>
	          <tr><td width=64 align=right><a href="javascript:validateForm(frmParts);"><IMG src="images/save_button.gif" width=60 height=20 border=0></a></td>
							<td width=64 align=right><a href="javascript:window.close();"><IMG src="images/close_button.gif" width=60 height=20 border=0></a></td></tr>
	     </table></td></tr>
	 </table>
</td></tr>
</form>
<!-- End Form Buttons -->
</body>
</html>
<%
if not( rsGetProductList is nothing ) then
	if( rsGetProductList.State = adStateOpen ) then
		rsGetProductList.Close
	end if
end if

if not( rsGetSupplierList is nothing ) then
	if( rsGetSupplierList.State = adStateOpen ) then
		rsGetSupplierList.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetProductList		= nothing
set cmdGetProductList		= nothing
set rsGetSupplierList		= nothing
set cmdGetSupplierList	= nothing
set cnnSQLConnection		= nothing
%>
