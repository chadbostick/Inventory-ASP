<%option explicit

dim cnnSQLConnection
dim rsGetSpindleCategoryBySpindleCategoryId, cmdGetSpindleCategoryBySpindleCategoryId, cmdEditSpindleCategory
dim intRowNumber, intSpindleCategoryId
dim bSaved

set cnnSQLConnection													= nothing
set rsGetSpindleCategoryBySpindleCategoryId		= nothing
set cmdGetSpindleCategoryBySpindleCategoryId	= nothing
set cmdEditSpindleCategory										= nothing

set cnnSQLConnection													= Server.CreateObject( "ADODB.Connection" )
set cmdGetSpindleCategoryBySpindleCategoryId	= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

bSaved= false

if( Request.Form.Count > 0 ) then

	set cmdEditSpindleCategory	= Server.CreateObject( "ADODB.Command" )
	
	cmdEditSpindleCategory.ActiveConnection	= cnnSQLConnection
	cmdEditSpindleCategory.CommandText				= "{? = call spEditSpindleCategory( ?,?,? ) }"
	
	cmdEditSpindleCategory.Parameters.Append cmdEditSpindleCategory.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
	cmdEditSpindleCategory.Parameters.Append cmdEditSpindleCategory.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
	cmdEditSpindleCategory.Parameters.Append cmdEditSpindleCategory.CreateParameter( "@SpindleCategoryId", adInteger, adParamInput, 4 )
	cmdEditSpindleCategory.Parameters.Append cmdEditSpindleCategory.CreateParameter( "@SpindleCategory", adVarChar, adParamInput, 50 )
		
	if( Request.Form( "txtSpindleCategoryId" ) <> "" ) then cmdEditSpindleCategory.Parameters( "@SpindleCategoryId" ) = Request.Form( "txtSpindleCategoryId" )
	if( Request.Form( "txtSpindleCategory" ) <> "" ) then cmdEditSpindleCategory.Parameters( "@SpindleCategory" ) = Request.Form( "txtSpindleCategory" )
		
	on error resume next
	cmdEditSpindleCategory.Execute
	if( Err.number <> 0 ) then
		Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
		Response.Write( Err.Description )
		Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
		Response.Write( "</font></body></html>" )
		Response.End
	end if
							
	on error goto 0

	intSpindleCategoryId = cmdEditSpindleCategory.Parameters( "@Return" ).Value
	bSaved = true
else
	intSpindleCategoryId = Request.QueryString( "intSpindleCategoryId" )
end if

cmdGetSpindleCategoryBySpindleCategoryId.ActiveConnection = cnnSQLConnection
cmdGetSpindleCategoryBySpindleCategoryId.CommandText = "{call spGetSpindleCategoryBySpindleCategoryId( ?,? ) }"
cmdGetSpindleCategoryBySpindleCategoryId.Parameters.Append cmdGetSpindleCategoryBySpindleCategoryId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetSpindleCategoryBySpindleCategoryId.Parameters.Append cmdGetSpindleCategoryBySpindleCategoryId.CreateParameter( "@SpindleCategoryId", adInteger, adParamInput, 4 )
cmdGetSpindleCategoryBySpindleCategoryId.Parameters( "@SpindleCategoryId" )	= intSpindleCategoryId

on error resume next
set rsGetSpindleCategoryBySpindleCategoryId = cmdGetSpindleCategoryBySpindleCategoryId.Execute
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
<TITLE>Centerline - Add Spindle Category</TITLE>
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
		window.opener.SpindleCategorySaved( <%=intSpindleCategoryId%>, "<%=rsGetSpindleCategoryBySpindleCategoryId( "SpindleCategory" )%>" );
	}	
	// End hiding script from old browsers -->
  </script>
<BODY topmargin=0 leftmargin=0 <%if (bSaved) then%>onload="UpdateOpener();"<%end if%>>

<!-- Page Contents -->
<table width=640 border=0 cellpadding=0 cellspacing=0>

<%
if not( rsGetSpindleCategoryBySpindleCategoryId is nothing ) then
	if( rsGetSpindleCategoryBySpindleCategoryId.State = adStateOpen ) then
		if( not rsGetSpindleCategoryBySpindleCategoryId.EOF ) then%>
<!--Top Title Bar -->
<tr><td class="pagetitle">  
 <table class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=3><IMG SRC="images/title_top.gif" width=100% height=6></td></tr>
   <%if( intSpindleCategoryId = 0 ) then%>
   <tr><td class="sectiontitle">&nbsp;&nbsp;New Spindle Category</td>
   <%else%>
   <tr><td class="sectiontitle">&nbsp;&nbsp;<%=rsGetSpindleCategoryBySpindleCategoryId( "SpindleCategory" )%></td>
   <%end if%>
     <td align=left width=348><IMG SRC="images/title_right.gif" width=23 height=25></td></tr>
 </table>
</td></tr>  <!-- End Top Title Bar -->

<!-- Shipping Method Form -->
<tr><td align=center>
 <form id="frmSpindleCategory" name="frmSpindleCategory" action="" method="post">
 <input type="hidden" name="txtSpindleCategoryId" value="<%=intSpindleCategoryId%>">
 <table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
   <tr><td class="label_1">Spindle Category</td><td class="input_1"><input type="text" class="forminput" name="txtSpindleCategory" onchange="this.isDirty=true;" maxlength=50 isRequired="true" fieldName="Spindle Category" value="<%=rsGetSpindleCategoryBySpindleCategoryId( "SpindleCategory" )%>"></td></tr>
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
	          <tr><td width=64 align=right><a href="javascript:validateForm(frmSpindleCategory);"><IMG src="images/save_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:frmSpindleCategory.reset();"><IMG src="images/clear_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:GetDirty(frmSpindleCategory);"><IMG src="images/close_button.gif" width=60 height=20 border=0></a></td></tr>
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
if not( rsGetSpindleCategoryBySpindleCategoryId is nothing ) then
	if( rsGetSpindleCategoryBySpindleCategoryId.State = adStateOpen ) then
		rsGetSpindleCategoryBySpindleCategoryId.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetSpindleCategoryBySpindleCategoryId		= nothing
set cmdGetSpindleCategoryBySpindleCategoryId	= nothing
set cmdEditSpindleCategory										= nothing
set cnnSQLConnection													= nothing
%>