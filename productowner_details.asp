<%option explicit

dim cnnSQLConnection
dim rsGetProductOwnerByProductOwnerId, cmdGetProductOwnerByProductOwnerId, cmdEditProductOwner
dim intRowNumber, intProductOwnerId
dim bSaved

set cnnSQLConnection													= nothing
set rsGetProductOwnerByProductOwnerId		= nothing
set cmdGetProductOwnerByProductOwnerId	= nothing
set cmdEditProductOwner										= nothing

set cnnSQLConnection													= Server.CreateObject( "ADODB.Connection" )
set cmdGetProductOwnerByProductOwnerId	= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

bSaved= false

if( Request.Form.Count > 0 ) then

	set cmdEditProductOwner	= Server.CreateObject( "ADODB.Command" )
	
	cmdEditProductOwner.ActiveConnection	= cnnSQLConnection
	cmdEditProductOwner.CommandText				= "{? = call spEditProductOwner( ?,?,? ) }"
	
	cmdEditProductOwner.Parameters.Append cmdEditProductOwner.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
	cmdEditProductOwner.Parameters.Append cmdEditProductOwner.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
	cmdEditProductOwner.Parameters.Append cmdEditProductOwner.CreateParameter( "@ProductOwnerId", adInteger, adParamInput, 4 )
	cmdEditProductOwner.Parameters.Append cmdEditProductOwner.CreateParameter( "@ProductOwner", adVarChar, adParamInput, 50 )
		
	if( Request.Form( "txtProductOwnerId" ) <> "" ) then cmdEditProductOwner.Parameters( "@ProductOwnerId" ) = Request.Form( "txtProductOwnerId" )
	if( Request.Form( "txtProductOwner" ) <> "" ) then cmdEditProductOwner.Parameters( "@ProductOwner" ) = Request.Form( "txtProductOwner" )
		
	on error resume next
	cmdEditProductOwner.Execute
	if( Err.number <> 0 ) then
		Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
		Response.Write( Err.Description )
		Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
		Response.Write( "</font></body></html>" )
		Response.End
	end if
							
	on error goto 0

	intProductOwnerId = cmdEditProductOwner.Parameters( "@Return" ).Value
	bSaved = true
else
	intProductOwnerId = Request.QueryString( "intProductOwnerId" )
end if

cmdGetProductOwnerByProductOwnerId.ActiveConnection = cnnSQLConnection
cmdGetProductOwnerByProductOwnerId.CommandText = "{call spGetProductOwnerByProductOwnerId( ?,? ) }"
cmdGetProductOwnerByProductOwnerId.Parameters.Append cmdGetProductOwnerByProductOwnerId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetProductOwnerByProductOwnerId.Parameters.Append cmdGetProductOwnerByProductOwnerId.CreateParameter( "@ProductOwnerId", adInteger, adParamInput, 4 )
cmdGetProductOwnerByProductOwnerId.Parameters( "@ProductOwnerId" )	= intProductOwnerId

on error resume next
set rsGetProductOwnerByProductOwnerId = cmdGetProductOwnerByProductOwnerId.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
							
on error goto 0%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Centerline - Add Product Owner</TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/default.css" TITLE="default_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/titles.css" TITLE="titles_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/forms.css" TITLE="forms_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/tables.css" TITLE="tables_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/pagelayout.css" TITLE="layout_css">
  <script type="text/javascript" src="js/forms.js"></script>
</HEAD>
  <script LANGUAGE=JavaScript>
	<!-- Hide script from old browsers	
	function UpdateOpener()
	{
		window.opener.ProductOwnerSaved( <%=intProductOwnerId%>, "<%=rsGetProductOwnerByProductOwnerId( "ProductOwner" )%>" );
	}	
	// End hiding script from old browsers -->
  </script>
<BODY topmargin=0 leftmargin=0 <%if (bSaved) then%>onload="UpdateOpener();"<%end if%>>

<!-- Page Contents -->
<table width=640 border=0 cellpadding=0 cellspacing=0>

<%
if not( rsGetProductOwnerByProductOwnerId is nothing ) then
	if( rsGetProductOwnerByProductOwnerId.State = adStateOpen ) then
		if( not rsGetProductOwnerByProductOwnerId.EOF ) then%>
<!--Top Title Bar -->
<tr><td class="pagetitle">  
 <table class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=3><IMG SRC="images/title_top.gif" width=100% height=6></td></tr>
   <%if( intProductOwnerId = 0 ) then%>
   <tr><td class="sectiontitle">&nbsp;&nbsp;New Product Owner</td>
   <%else%>
   <tr><td class="sectiontitle">&nbsp;&nbsp;<%=rsGetProductOwnerByProductOwnerId( "ProductOwner" )%></td>
   <%end if%>
     <td align=left width=348><IMG SRC="images/title_right.gif" width=23 height=25></td></tr>
 </table>
</td></tr>  <!-- End Top Title Bar -->

<!-- Shipping Method Form -->
<tr><td align=center>
 <form id="frmProductOwner" name="frmProductOwner" action="" method="post">
 <input type="hidden" name="txtProductOwnerId" value="<%=intProductOwnerId%>">
 <table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
   <tr><td class="label_1">Product Owner</td><td class="input_1"><input type="text" class="forminput" name="txtProductOwner" onchange="this.isDirty=true;" maxlength=50 isRequired="true" fieldName="Product Owner" value="<%=rsGetProductOwnerByProductOwnerId( "ProductOwner" )%>"></td></tr>
   </table>
 
  <!-- Bottom Rule -->
 <tr><td height=12>
  <TABLE class="" width=100% cellpadding=0 cellspacing=0>
    <tr><td><IMG SRC="images/small_title_top.gif" width=100% height=3></td></tr>
  </TABLE>
 </td></tr>  <!-- End Bottom Rule -->
 
  <tr><td align=center>
   	<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
	    <tr><td colspan=4><table cellspacing=0 cellpadding=0 width=192 align=right>
	          <tr><td width=64 align=right><a href="javascript:validateForm(frmProductOwner);"><IMG src="images/save_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:frmProductOwner.reset();"><IMG src="images/clear_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:GetDirty(frmProductOwner);"><IMG src="images/close_button.gif" width=60 height=20 border=0></a></td></tr>
	     </table></td></tr>
	 </table>
 </td></tr>
 
 </form>   <!-- End Shipping Method Form -->
 <%	end if
	end if
end if%>
</body>
</html>
<%
if not( rsGetProductOwnerByProductOwnerId is nothing ) then
	if( rsGetProductOwnerByProductOwnerId.State = adStateOpen ) then
		rsGetProductOwnerByProductOwnerId.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetProductOwnerByProductOwnerId		= nothing
set cmdGetProductOwnerByProductOwnerId	= nothing
set cmdEditProductOwner										= nothing
set cnnSQLConnection													= nothing
%>