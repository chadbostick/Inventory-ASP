<%option explicit

dim cnnSQLConnection
dim rsGetWorkOrderPriorityByWorkOrderPriorityId, cmdGetWorkOrderPriorityByWorkOrderPriorityId, cmdEditWorkOrderPriority
dim intRowNumber, intWorkOrderPriorityId
dim bSaved

set cnnSQLConnection													= nothing
set rsGetWorkOrderPriorityByWorkOrderPriorityId		= nothing
set cmdGetWorkOrderPriorityByWorkOrderPriorityId	= nothing
set cmdEditWorkOrderPriority										= nothing

set cnnSQLConnection													= Server.CreateObject( "ADODB.Connection" )
set cmdGetWorkOrderPriorityByWorkOrderPriorityId	= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

bSaved = false

if( Request.Form.Count > 0 ) then

	set cmdEditWorkOrderPriority	= Server.CreateObject( "ADODB.Command" )
	
	cmdEditWorkOrderPriority.ActiveConnection	= cnnSQLConnection
	cmdEditWorkOrderPriority.CommandText			= "{? = call spEditWorkOrderPriority( ?,?,? ) }"
	
	cmdEditWorkOrderPriority.Parameters.Append cmdEditWorkOrderPriority.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
	cmdEditWorkOrderPriority.Parameters.Append cmdEditWorkOrderPriority.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
	cmdEditWorkOrderPriority.Parameters.Append cmdEditWorkOrderPriority.CreateParameter( "@WorkOrderPriorityId", adInteger, adParamInput, 4 )
	cmdEditWorkOrderPriority.Parameters.Append cmdEditWorkOrderPriority.CreateParameter( "@WorkOrderPriority", adVarChar, adParamInput, 50 )
		
	if( Request.Form( "txtWorkOrderPriorityId" ) <> "" ) then cmdEditWorkOrderPriority.Parameters( "@WorkOrderPriorityId" ) = Request.Form( "txtWorkOrderPriorityId" )
	if( Request.Form( "txtWorkOrderPriority" ) <> "" ) then cmdEditWorkOrderPriority.Parameters( "@WorkOrderPriority" ) = Request.Form( "txtWorkOrderPriority" )
	
	on error resume next
	cmdEditWorkOrderPriority.Execute
	if( Err.number <> 0 ) then
		Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
		Response.Write( Err.Description )
		Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
		Response.Write( "</font></body></html>" )
		Response.End
	end if
	
	on error goto 0

	intWorkOrderPriorityId = cmdEditWorkOrderPriority.Parameters( "@Return" ).Value
	bSaved = true
else
	intWorkOrderPriorityId = Request.QueryString( "intWorkOrderPriorityId" )
end if

cmdGetWorkOrderPriorityByWorkOrderPriorityId.ActiveConnection = cnnSQLConnection
cmdGetWorkOrderPriorityByWorkOrderPriorityId.CommandText = "{call spGetWorkOrderPriorityByWorkOrderPriorityId( ?,? ) }"
cmdGetWorkOrderPriorityByWorkOrderPriorityId.Parameters.Append cmdGetWorkOrderPriorityByWorkOrderPriorityId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetWorkOrderPriorityByWorkOrderPriorityId.Parameters.Append cmdGetWorkOrderPriorityByWorkOrderPriorityId.CreateParameter( "@WorkOrderPriorityId", adInteger, adParamInput, 4 )
cmdGetWorkOrderPriorityByWorkOrderPriorityId.Parameters( "@WorkOrderPriorityId" )	= intWorkOrderPriorityId

on error resume next
set rsGetWorkOrderPriorityByWorkOrderPriorityId = cmdGetWorkOrderPriorityByWorkOrderPriorityId.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
	
on error goto 0
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Centerline - Add Work Order Priority</TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/default.css" TITLE="default_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/titles.css" TITLE="titles_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/forms.css" TITLE="forms_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/tables.css" TITLE="tables_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/pagelayout.css" TITLE="layout_css">
  <script type="text/javascript" src="js/forms.js"></script>
  <script LANGUAGE=JavaScript>
	<!-- Hide script from old browsers	
	function UpdateOpener()
	{
		window.opener.WorkOrderPrioritySaved( <%=intWorkOrderPriorityId%>, "<%=rsGetWorkOrderPriorityByWorkOrderPriorityId( "WorkOrderPriority" )%>" );
	}	
	// End hiding script from old browsers -->
  </script>
</HEAD>
<BODY topmargin=0 leftmargin=0 <%if (bSaved) then%>onload="UpdateOpener();"<%end if%>>

<!-- Page Contents -->
<table width=640 border=0 cellpadding=0 cellspacing=0>

<%
if not( rsGetWorkOrderPriorityByWorkOrderPriorityId is nothing ) then
	if( rsGetWorkOrderPriorityByWorkOrderPriorityId.State = adStateOpen ) then
		if( not rsGetWorkOrderPriorityByWorkOrderPriorityId.EOF ) then%>
<!--Top Title Bar -->
<tr><td class="pagetitle">  
 <table class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=3><IMG SRC="images/title_top.gif" width=100% height=6></td></tr>
   <%if( intWorkOrderPriorityId = 0 ) then%>
   <tr><td class="sectiontitle">&nbsp;&nbsp;New Work Order Priority</td>
   <%else%>
   <tr><td class="sectiontitle">&nbsp;&nbsp;<%=rsGetWorkOrderPriorityByWorkOrderPriorityId( "WorkOrderPriority" )%></td>
   <%end if%>
     <td align=left width=348><IMG SRC="images/title_right.gif" width=23 height=25></td></tr>
 </table>
</td></tr>  <!-- End Top Title Bar -->

<!-- Shipping Method Form -->
<tr><td align=center>
 <form id="frmWorkOrderPriority" name="frmWorkOrderPriority" action="" method="post">
 <input type="hidden" name="txtWorkOrderPriorityId" value="<%=intWorkOrderPriorityId%>">
 <table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
   <tr><td class="label_1">Work Order Priority</td><td class="input_1"><input type="text" class="forminput" onchange="this.isDirty=true;" name="txtWorkOrderPriority" maxlength=50 isRequired="true" fieldName="WorkOrder Priority" value="<%=rsGetWorkOrderPriorityByWorkOrderPriorityId( "WorkOrderPriority" )%>"></td></tr>
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
	          <tr><td width=64 align=right><a href="javascript:validateForm(frmWorkOrderPriority);"><IMG src="images/save_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:frmWorkOrderPriority.reset();"><IMG src="images/clear_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:GetDirty(frmWorkOrderPriority);"><IMG src="images/close_button.gif" width=60 height=20 border=0></a></td></tr>
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
if not( rsGetWorkOrderPriorityByWorkOrderPriorityId is nothing ) then
	if( rsGetWorkOrderPriorityByWorkOrderPriorityId.State = adStateOpen ) then
		rsGetWorkOrderPriorityByWorkOrderPriorityId.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetWorkOrderPriorityByWorkOrderPriorityId		= nothing
set cmdGetWorkOrderPriorityByWorkOrderPriorityId	= nothing
set cmdEditWorkOrderPriority										= nothing
set cnnSQLConnection													= nothing
%>