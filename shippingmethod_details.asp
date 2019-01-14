<%option explicit

dim cnnSQLConnection
dim rsGetShippingMethodByShippingMethodId, cmdGetShippingMethodByShippingMethodId, cmdEditShippingMethod
dim intRowNumber, intShippingMethodId

set cnnSQLConnection												= nothing
set rsGetShippingMethodByShippingMethodId		= nothing
set cmdGetShippingMethodByShippingMethodId	= nothing
set cmdEditShippingMethod										= nothing

set cnnSQLConnection												= Server.CreateObject( "ADODB.Connection" )
set cmdGetShippingMethodByShippingMethodId	= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

if( Request.Form.Count > 0 ) then

	set cmdEditShippingMethod	= Server.CreateObject( "ADODB.Command" )
	
	cmdEditShippingMethod.ActiveConnection	= cnnSQLConnection
	cmdEditShippingMethod.CommandText				= "{? = call spEditShippingMethod( ?,?,? ) }"
	
	cmdEditShippingMethod.Parameters.Append cmdEditShippingMethod.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
	cmdEditShippingMethod.Parameters.Append cmdEditShippingMethod.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
	cmdEditShippingMethod.Parameters.Append cmdEditShippingMethod.CreateParameter( "@ShippingMethodId", adInteger, adParamInput, 4 )
	cmdEditShippingMethod.Parameters.Append cmdEditShippingMethod.CreateParameter( "@ShippingMethod", adVarChar, adParamInput, 50 )
		
	if( Request.Form( "txtShippingMethodId" ) <> "" ) then cmdEditShippingMethod.Parameters( "@ShippingMethodId" ) = Request.Form( "txtShippingMethodId" )
	if( Request.Form( "txtShippingMethod" ) <> "" ) then cmdEditShippingMethod.Parameters( "@ShippingMethod" ) = Request.Form( "txtShippingMethod" )
	
	on error resume next
	cmdEditShippingMethod.Execute
	if( Err.number <> 0 ) then
		Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
		Response.Write( Err.Description )
		Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
		Response.Write( "</font></body></html>" )
		Response.End
	end if

	on error goto 0

	intShippingMethodId = cmdEditShippingMethod.Parameters( "@Return" ).Value
else
	intShippingMethodId = Request.QueryString( "intShippingMethodId" )
end if

cmdGetShippingMethodByShippingMethodId.ActiveConnection = cnnSQLConnection
cmdGetShippingMethodByShippingMethodId.CommandText = "{call spGetShippingMethodByShippingMethodId( ?,? ) }"
cmdGetShippingMethodByShippingMethodId.Parameters.Append cmdGetShippingMethodByShippingMethodId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetShippingMethodByShippingMethodId.Parameters.Append cmdGetShippingMethodByShippingMethodId.CreateParameter( "@ShippingMethodId", adInteger, adParamInput, 4 )
cmdGetShippingMethodByShippingMethodId.Parameters( "@ShippingMethodId" )	= intShippingMethodId

on error resume next
set rsGetShippingMethodByShippingMethodId = cmdGetShippingMethodByShippingMethodId.Execute
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
<TITLE>Centerline - Add Shipping Method</TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/default.css" TITLE="default_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/titles.css" TITLE="titles_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/forms.css" TITLE="forms_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/tables.css" TITLE="tables_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/pagelayout.css" TITLE="layout_css">
  <script type="text/javascript" src="js/forms.js"></script>
</HEAD>
<BODY topmargin=0 leftmargin=0>

<!-- Page Contents -->
<table width=640 border=0 cellpadding=0 cellspacing=0>

<%
if not( rsGetShippingMethodByShippingMethodId is nothing ) then
	if( rsGetShippingMethodByShippingMethodId.State = adStateOpen ) then
		if( not rsGetShippingMethodByShippingMethodId.EOF ) then%>
<!--Top Title Bar -->
<tr><td class="pagetitle">  
 <table class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=3><IMG SRC="images/title_top.gif" width=100% height=6></td></tr>
   <%if( intShippingMethodId = 0 ) then%>
   <tr><td class="sectiontitle">&nbsp;&nbsp;New Shipping Method</td>
   <%else%>
   <tr><td class="sectiontitle">&nbsp;&nbsp;<%=rsGetShippingMethodByShippingMethodId( "ShippingMethod" )%></td>
   <%end if%>
     <td align=left width=348><IMG SRC="images/title_right.gif" width=23 height=25></td></tr>
 </table>
</td></tr>  <!-- End Top Title Bar -->

<!-- Shipping Method Form -->
<tr><td align=center>
 <form id="frmShippingMethod" name="frmShippingMethod" action="" method="post">
 <input type="hidden" name="txtShippingMethodId" value="<%=intShippingMethodId%>">
 <table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
   <tr><td class="label_1">Shipping Method</td><td class="input_1"><input type="text" class="forminput" name="txtShippingMethod" onchange="this.isDirty=true;" maxlength=50 isRequired="true" fieldName="Shipping Method" value="<%=rsGetShippingMethodByShippingMethodId( "ShippingMethod" )%>"></td></tr>
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
	          <tr><td width=64 align=right><a href="javascript:validateForm(frmShippingMethod);"><IMG src="images/save_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:frmShippingMethod.reset();"><IMG src="images/clear_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:GetDirty(frmShippingMethod);"><IMG src="images/close_button.gif" width=60 height=20 border=0></a></td></tr>
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
if not( rsGetShippingMethodByShippingMethodId is nothing ) then
	if( rsGetShippingMethodByShippingMethodId.State = adStateOpen ) then
		rsGetShippingMethodByShippingMethodId.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetShippingMethodByShippingMethodId		= nothing
set cmdGetShippingMethodByShippingMethodId	= nothing
set cmdEditShippingMethod							= nothing
set cnnSQLConnection						= nothing
%>