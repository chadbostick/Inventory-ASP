<%option explicit

dim cnnSQLConnection
dim cmdDeleteQuoteByQuoteId

set cnnSQLConnection								= nothing
set cmdDeleteQuoteByQuoteId		= nothing

set cnnSQLConnection								= Server.CreateObject( "ADODB.Connection" )
set cmdDeleteQuoteByQuoteId		= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

cmdDeleteQuoteByQuoteId.ActiveConnection	= cnnSQLConnection
cmdDeleteQuoteByQuoteId.CommandText				= "{call spDeleteQuoteByQuoteId( ?,? ) }"

cmdDeleteQuoteByQuoteId.Parameters.Append cmdDeleteQuoteByQuoteId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdDeleteQuoteByQuoteId.Parameters.Append cmdDeleteQuoteByQuoteId.CreateParameter( "@QuoteId", adInteger, adParamInput, 4)

cmdDeleteQuoteByQuoteId.Parameters( "@QuoteId" ) = Request.QueryString( "intQuoteId" )

on error resume next
cmdDeleteQuoteByQuoteId.Execute
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
<TITLE>Centerline - Delete</TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/default.css" TITLE="default_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/titles.css" TITLE="titles_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/forms.css" TITLE="forms_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/tables.css" TITLE="tables_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/pagelayout.css" TITLE="layout_css">
</HEAD>
<BODY topmargin=0 leftmargin=0 onunload="javascript:window.opener.RefreshCurrentPage();">

<!-- Page Contents -->
<table width=640 border=0 cellpadding=0 cellspacing=0>

 <!--Top Title Bar -->
 <tr><td class="pagetitle">  
  <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
    <tr><td colspan=3><IMG SRC="images/title_top.gif" width=100% height=6></td></tr>
    <tr><td class="sectiontitle" width=20%>Delete Record</td>
      <td align=left width=348><IMG SRC="images/title_right.gif" width=23 height=25></td></tr>
  </TABLE>
</td></tr>
<!-- End Top Title Bar -->

<!-- Print Form -->
<tr><td align=center>	
   <table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
     <tr><td colspan=2>The selected record was successfully removed from the system.</td></tr>
   </table>
</td></tr>
 
<!-- Bottom Rule -->
<tr><td height=12>
<TABLE class="" width=100% cellpadding=0 cellspacing=0>
	<tr><td><IMG SRC="images/small_title_top.gif" width=100% height=4></td></tr>
</TABLE>
</td></tr>
<!-- End Bottom Rule -->
 
<tr><td align=center>
 	<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
    <tr><td colspan=4><table cellspacing=0 cellpadding=0 width=64 align=right>
          <tr><td width=64 align=right><a href="javascript:window.close();"><IMG src="images/close_button.gif" width=60 height=20 border=0></a></td></tr>
     </table></td></tr>
 </table>
</td></tr>
</BODY>
</HTML>
<%
if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set cmdDeleteQuoteByQuoteId	= nothing
set cnnSQLConnection									= nothing
%>