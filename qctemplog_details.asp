<%option explicit
%><!--#include file="global_functions.asp"--><%

dim cnnSQLConnection
dim rsGetQCTempLogByQCTempLogId, cmdGetQCTempLogByQCTempLogId
dim cmdEditQCTempLog
dim intQCTempLogId, intWorkOrderId, intWorkOrderNumber, intRowNumber, strCurrentPage

set cnnSQLConnection								= nothing
set rsGetQCTempLogByQCTempLogId			= nothing
set cmdGetQCTempLogByQCTempLogId		= nothing
set cmdEditQCTempLog								= nothing


set cnnSQLConnection								= Server.CreateObject( "ADODB.Connection" )
set cmdGetQCTempLogByQCTempLogId		= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

if( Request.Form.Count > 0 ) then
	if( Request.Form( "txtQCTempLogIdMaster" ) <> "" ) then
	
		set cmdEditQCTempLog	= Server.CreateObject( "ADODB.Command" )
	
		cmdEditQCTempLog.ActiveConnection	= cnnSQLConnection
		cmdEditQCTempLog.CommandText			= "{? = call spEditQCTempLog( ?,?,?,?,?,? ) }"
		
		cmdEditQCTempLog.Parameters.Append cmdEditQCTempLog.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
		cmdEditQCTempLog.Parameters.Append cmdEditQCTempLog.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
		cmdEditQCTempLog.Parameters.Append cmdEditQCTempLog.CreateParameter( "@QCTempLogId", adInteger, adParamInput, 4 )
		cmdEditQCTempLog.Parameters.Append cmdEditQCTempLog.CreateParameter( "@WorkOrderId", adInteger, adParamInput, 4 )
		cmdEditQCTempLog.Parameters.Append cmdEditQCTempLog.CreateParameter( "@QCDate", adVarChar, adParamInput, 10 )
		cmdEditQCTempLog.Parameters.Append cmdEditQCTempLog.CreateParameter( "@QCMaxSpeed", adVarChar, adParamInput, 100 )
		cmdEditQCTempLog.Parameters.Append cmdEditQCTempLog.CreateParameter( "@QCTotalRunTime", adVarChar, adParamInput, 100 )
			
		if( Request.Form( "txtQCTempLogIdMaster" ) <> "" ) then cmdEditQCTempLog.Parameters( "@QCTempLogId" ) = Request.Form( "txtQCTempLogIdMaster" )
		if( Request.Form( "txtWorkOrderId" ) <> "" ) then cmdEditQCTempLog.Parameters( "@WorkOrderId" ) = Request.Form( "txtWorkOrderId" )
		if( Request.Form( "txtQCDate" ) <> "" ) then cmdEditQCTempLog.Parameters( "@QCDate" ) = Request.Form( "txtQCDate" )
		if( Request.Form( "txtQCMaxSpeed" ) <> "" ) then cmdEditQCTempLog.Parameters( "@QCMaxSpeed" ) = Request.Form( "txtQCMaxSpeed" )
		if( Request.Form( "txtQCTotalRunTime" ) <> "" ) then cmdEditQCTempLog.Parameters( "@QCTotalRunTime" ) = Request.Form( "txtQCTotalRunTime" )
		
		
		on error resume next
		cmdEditQCTempLog.Execute
		if( Err.number <> 0 ) then
			Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
			Response.Write( Err.Description )
			Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
			Response.Write( "</font></body></html>" )
			Response.End
		end if

		on error goto 0
	
		intQCTempLogId = cmdEditQCTempLog.Parameters( "@Return" ).Value
		
	elseif( Request.Form( "txtQCTempLogIdDetail" ) <> "" ) then
	
		set cmdEditQCTempLog	= Server.CreateObject( "ADODB.Command" )
	
		cmdEditQCTempLog.ActiveConnection	= cnnSQLConnection
		cmdEditQCTempLog.CommandText			= "{? = call spEditQCTemp( ?,?,?,?,?,?,?,?,?,? ) }"
		
		cmdEditQCTempLog.Parameters.Append cmdEditQCTempLog.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
		cmdEditQCTempLog.Parameters.Append cmdEditQCTempLog.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
		cmdEditQCTempLog.Parameters.Append cmdEditQCTempLog.CreateParameter( "@QCTempId", adInteger, adParamInput, 4 )
		cmdEditQCTempLog.Parameters.Append cmdEditQCTempLog.CreateParameter( "@QCTempLogId", adInteger, adParamInput, 4 )
		cmdEditQCTempLog.Parameters.Append cmdEditQCTempLog.CreateParameter( "@QCTime", adVarChar, adParamInput, 100 )
		cmdEditQCTempLog.Parameters.Append cmdEditQCTempLog.CreateParameter( "@QCSpeed", adVarChar, adParamInput, 100 )
		cmdEditQCTempLog.Parameters.Append cmdEditQCTempLog.CreateParameter( "@QCFront", adVarChar, adParamInput, 100 )
		cmdEditQCTempLog.Parameters.Append cmdEditQCTempLog.CreateParameter( "@QCRear", adVarChar, adParamInput, 100 )
		cmdEditQCTempLog.Parameters.Append cmdEditQCTempLog.CreateParameter( "@QCShaft", adVarChar, adParamInput, 100 )
		cmdEditQCTempLog.Parameters.Append cmdEditQCTempLog.CreateParameter( "@QCTempLocation", adVarChar, adParamInput, 100 )
		cmdEditQCTempLog.Parameters.Append cmdEditQCTempLog.CreateParameter( "@QCNotes", adLongVarChar, adParamInput, 1048576 )

			
		if( Request.Form( "txtQCTempId" ) <> "" ) then cmdEditQCTempLog.Parameters( "@QCTempId" ) = Request.Form( "txtQCTempId" )
		if( Request.Form( "txtQCTempLogIdDetail" ) <> "" ) then cmdEditQCTempLog.Parameters( "@QCTempLogId" ) = Request.Form( "txtQCTempLogIdDetail" )
		if( Request.Form( "txtQCTime" ) <> "" ) then cmdEditQCTempLog.Parameters( "@QCTime" ) = Request.Form( "txtQCTime" )
		if( Request.Form( "txtQCSpeed" ) <> "" ) then cmdEditQCTempLog.Parameters( "@QCSpeed" ) = Request.Form( "txtQCSpeed" )
		if( Request.Form( "txtQCFront" ) <> "" ) then cmdEditQCTempLog.Parameters( "@QCFront" ) = Request.Form( "txtQCFront" )
		if( Request.Form( "txtQCRear" ) <> "" ) then cmdEditQCTempLog.Parameters( "@QCRear" ) = Request.Form( "txtQCRear" )
		if( Request.Form( "txtQCShaft" ) <> "" ) then cmdEditQCTempLog.Parameters( "@QCShaft" ) = Request.Form( "txtQCShaft" )
		if( Request.Form( "txtQCTempLocation" ) <> "" ) then cmdEditQCTempLog.Parameters( "@QCTempLocation" ) = Request.Form( "txtQCTempLocation" )
		if( Request.Form( "txtQCNotes" ) <> "" ) then cmdEditQCTempLog.Parameters( "@QCNotes" ) = Request.Form( "txtQCNotes" )
		
		on error resume next
		cmdEditQCTempLog.Execute
		if( Err.number <> 0 ) then
			Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
			Response.Write( Err.Description )
			Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
			Response.Write( "</font></body></html>" )
			Response.End
		end if

		on error goto 0	
		
		
		intQCTempLogId = Request.Form( "txtQCTempLogIdDetail" )
		
	end if
else
	intQCTempLogId = Request.QueryString( "intQCTempLogId" )
end if

cmdGetQCTempLogByQCTempLogId.ActiveConnection = cnnSQLConnection
cmdGetQCTempLogByQCTempLogId.CommandText = "{call spGetQCTempLogByQCTempLogId( ?,? ) }"
cmdGetQCTempLogByQCTempLogId.Parameters.Append cmdGetQCTempLogByQCTempLogId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetQCTempLogByQCTempLogId.Parameters.Append cmdGetQCTempLogByQCTempLogId.CreateParameter( "@QCTempLogId", adInteger, adParamInput, 4)
cmdGetQCTempLogByQCTempLogId.Parameters( "@QCTempLogId" )	= intQCTempLogId

on error resume next
set rsGetQCTempLogByQCTempLogId = cmdGetQCTempLogByQCTempLogId.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if

on error goto 0
'**********************************************************************************
if( Request.ServerVariables( "QUERY_STRING" ) <> "" ) then
	strCurrentPage = Request.ServerVariables( "URL" ) + "?" + Request.ServerVariables( "QUERY_STRING" )
else
	strCurrentPage = Request.ServerVariables( "URL" )
end if
%>
<HTML>
<HEAD>
<TITLE>Centerline - View / Edit Bearings Cost</TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/default.css" TITLE="default_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/titles.css" TITLE="titles_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/forms.css" TITLE="forms_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/tables.css" TITLE="tables_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/pagelayout.css" TITLE="layout_css">
  <script type="text/javascript" src="js/common.js"></script>
  <script type="text/javascript" src="js/forms.js"></script>
  <script language="javascript">
	function ConfirmDelete( intQCTempId, strTime )
	{
		if( confirm( 'Are you sure you want to remove this item?' ) )
		{
			OpenWindow( 'qctemp_delete.asp?intQCTempId='+intQCTempId,'DeleteRecord','width=657,height=120,left='+GetCenterX( 657 )+',top='+GetCenterY( 120 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 120 )+',scrollbars=yes,dependent=yes' );
		}
	}	
	function EditQCTemp( intQCTempId, intQCTempLogId)
	{
		OpenWindow( 'qctemp_details_edit.asp?intQCTempId='+intQCTempId + '&intQCTempLogId=' + intQCTempLogId,'EditTemp','width=700,height=160,left='+GetCenterX( 700 )+',top='+GetCenterY( 160 )+',screenX='+GetCenterX( 700 )+',screenY='+GetCenterY( 160 )+',scrollbars=yes,dependent=yes' );
	}
	function OpenWorkOrder( strWorkOrderId )
	{
	  OpenWindow( 'workorder_details.asp?intWorkOrderId='+strWorkOrderId,'CenterlineWorkOrder'+strWorkOrderId,'width=657,height=480,left='+GetCenterX( 657 )+',top='+GetCenterY( 480 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 480 )+',scrollbars=yes,dependent=yes' );
	}
	
	function RefreshCurrentPage()
	{
		document.location.href = '<%=strCurrentPage%>';
	}
  </script>
</head>
<body topmargin=0 leftmargin=0>

<!-- Page Contents -->
<table width=640 border=0 cellpadding=0 cellspacing=0>

<%if not( rsGetQCTempLogByQCTempLogId is nothing ) then
		if( rsGetQCTempLogByQCTempLogId.State = adStateOpen ) then
			if( not rsGetQCTempLogByQCTempLogId.EOF ) then%>
				<%if (intQCTempLogId <> 0) then
						intWorkOrderId = rsGetQCTempLogByQCTempLogId( "WorkOrderId" )
						intWorkOrderNumber = rsGetQCTempLogByQCTempLogId( "WorkOrderNumber" )
					else
						intWorkOrderId =  Request.QueryString( "intWorkOrderId" )
						intWorkOrderNumber =  Request.QueryString( "intWorkOrderNumber" )
					end if%>
						
<!--Top Title Bar -->
<tr><td class="pagetitle">  
 <table class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=3><IMG SRC="images/title_top.gif" width=100% height=6></td></tr>
   <tr><td class="sectiontitle" width=20%>QC Temp Log</td>
     <td align=left width=348><IMG SRC="images/title_right.gif" width=23 height=25></td>
     <td class="sectioncaption" align=right>W.O. #&nbsp;&nbsp;<a href="javascript:OpenWorkOrder('<%=intWorkOrderId%>');"><%=intWorkOrderNumber%></a></td></tr>
 </table>
</td></tr>
<!-- End Top Title Bar -->

<!-- Cost Form -->
<form id="frmQCTempLog" name="frmQCTempLog" action="" method="post">
<input type="hidden" name="txtQCTempLogIdMaster" value="<%=intQCTempLogId%>">
<input type="hidden" name="txtWorkOrderId" value="<%=intWorkOrderId%>">
<tr><td align=center>
	<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
	    <tr><td class="label_1">Work Order</td><td class="generic_input"><a href="javascript:OpenWorkOrder('<%=intWorkOrderId%>');"><%=intWorkOrderNumber%></a></td>
			  <td class="label_2">Date</td><td class="input_2"><input type="text" class="forminput" name="txtQCDate" maxlength=100 onchange="this.isDirty=true;" fieldName="Date" value="<%if( IsDate( rsGetQCTempLogByQCTempLogId( "QCDate" ) ) ) then Response.Write( FormatDateTime( rsGetQCTempLogByQCTempLogId( "QCDate" ), vbShortDate ) ) end if%>"></td></tr>
		
		<tr><td class="label_1">Max Speed</td><td class="input_1"><input type="text" class="forminput" name="txtQCMaxSpeed" onchange="this.isDirty=true;" fieldName="Max Speed" value="<%=rsGetQCTempLogByQCTempLogId( "QCMaxSpeed" )%>"></td>
			<td class="label_2">Total Run Time</td><td class="input_2"><input type="text" class="forminput" name="txtQCTotalRunTime" onchange="this.isDirty=true;" fieldName="Total Run Time" value="<%=rsGetQCTempLogByQCTempLogId( "QCTotalRunTime" )%>"></td></tr>
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
	    <tr><td colspan=4><table cellspacing=0 cellpadding=0 width=320 align=right>
	          <tr><td width=64 align=right><a href="javascript:window.navigate( '<%=Request.QueryString( "strPrevPage" )%>' );"><IMG src="images/back_button.gif" width=60 height=20 border=0></a></td>
							<td width=64 align=right><a href="javascript:validateForm(frmQCTempLog);"><IMG src="images/save_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:frmQCTempLog.reset();"><IMG src="images/clear_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:OpenPrintBox('quote_bearingorder.asp?intQCTempLogId=<%=intQCTempLogId%>');"><IMG src="images/print_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:GetDirty(frmQCTempLog);"><IMG src="images/close_button.gif" width=60 height=20 border=0></a></td></tr>
	     </table></td></tr>
	 </table>
</td></tr>
<!-- End Form Buttons -->
</form>
<%	set rsGetQCTempLogByQCTempLogId = rsGetQCTempLogByQCTempLogId.NextRecordset
		end if
	end if
end if%>
<!-- End QC Temp Log Form -->
 
 

<%if( intQCTempLogId <> 0 ) then%>
<!-- Middle Title Bar -->
<tr><td class="pagesmalltitle">
 <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=2><IMG SRC="images/small_title_top.gif" width=100% height=5></td></tr>
   <tr><td class="smallsectiontitle" width=15%>Temps</td>
     <td align=left width=448><IMG SRC="images/small_title_right.gif" width=15 height=17></td>
 </TABLE>
</td></tr>
<!-- End Middle Title Bar -->
  
 <!-- QCTemps Table -->
 <form id="frmQCTemps" name="frmQCTemps" action="" method="post">
 <input type="hidden" name="txtQCTempLogIdDetail" value="<%=intQCTempLogId%>">
 <tr><td align=center>
  <table width=600 border=0 cellpadding=0 cellspacing=3 align=middle>
  <tr>
  <td><table class="itemlist" width=600 border=1 cellpadding=2 cellspacing=0 rules="cols" align=middle>
  
  <!-- table heading -->
  <tr>
    <th class="itemlist" width="15" align="center">Del</th>
    <th class="itemlist" width="15" align="center">Edit</th>
    <th class="itemlist" width="60">Time</th>
    <th class="itemlist" width="60">Speed</th>
    <th class="itemlist" width="60">Front</th>
    <th class="itemlist" width="60">Rear</th>
    <th class="itemlist" width="60">Shaft</th>
    <th class="itemlist" width="100">Temp Location</th>
    <th class="itemlist">Notes</th>
  </tr>  
 <!-- table data -->
	<%if not( rsGetQCTempLogByQCTempLogId is nothing ) then
			if( rsGetQCTempLogByQCTempLogId.State = adStateOpen ) then
				intRowNumber = 1
				
				while( not rsGetQCTempLogByQCTempLogId.EOF )
					if( ( intRowNumber Mod 2 ) = 1 ) then%>
						<tr class="evenrow">
						<%else%>
						<tr class="oddrow">
						<%end if%>
							<td class="itemlist" width="15" align=middle><a href="javascript:ConfirmDelete('<%=rsGetQCTempLogByQCTempLogId( "QCTempId" )%>', '<%=rsGetQCTempLogByQCTempLogId( "QCTime" )%>');"><img src="images/form_delete.gif" width=15 height=15 border=0></a></td>
							<td class="itemlist" width="15" align=middle><a href="javascript:EditQCTemp('<%=rsGetQCTempLogByQCTempLogId( "QCTempId" )%>','<%=rsGetQCTempLogByQCTempLogId( "QCTempLogId" )%>');"><img src="images/form_edit.gif" width=15 height=15 border=0></a></td>
							<td class="itemlist" width="60"><%=rsGetQCTempLogByQCTempLogId( "QCTime" )%></td>
							<td class="itemlist" width="60"><%=rsGetQCTempLogByQCTempLogId( "QCSpeed" )%></td>
							<td class="itemlist" width="60"><%=rsGetQCTempLogByQCTempLogId( "QCFront" )%></td>
							<td class="itemlist" width="60"><%=rsGetQCTempLogByQCTempLogId( "QCRear" )%></td>
							<td class="itemlist" width="60"><%=rsGetQCTempLogByQCTempLogId( "QCShaft" )%></td>
							<td class="itemlist" width="100"><%=rsGetQCTempLogByQCTempLogId( "QCTempLocation" )%></td>
							<td class="itemlist"><%=rsGetQCTempLogByQCTempLogId( "QCNotes" )%></td>
						</tr>
						<%intRowNumber = intRowNumber + 1
						rsGetQCTempLogByQCTempLogId.MoveNext
						wend
					end if
				end if%>  
		<%if( ( intRowNumber Mod 2 ) = 1 ) then%>
		<tr class="evenrow">
		<%else%>
		<tr class="oddrow">
		<%end if%>
			<td width="15">&nbsp;</td>
			<td width="15">&nbsp;</td>
			<td width="60"><input type="text" class="forminput" name="txtQCTime" isRequired="true" fieldName="Time" value=""></td>
			<td width="60"><input type="text" class="forminput" name="txtQCSpeed" fieldName="Speed" value=""></td>
			<td width="60"><input type="text" class="forminput" name="txtQCFront" fieldName="Front" value=""></td>
			<td width="60"><input type="text" class="forminput" name="txtQCRear" fieldName="Rear" value=""></td>
			<td width="60"><input type="text" class="forminput" name="txtQCShaft" fieldName="Shaft" value=""></td>
			<td width="100"><input type="text" class="forminput" name="txtQCTempLocation" fieldName="Temp Location" value=""></td>
			<td><input type="text" class="forminput" name="txtQCNotes" fieldName="Notes" value=""></td>
		</tr>
 </table></td>
 </tr>
	<tr><table width=660><tr>
	
	<td><table cellspacing=0 cellpadding=0 width=192 align=right>
		<tr><td align=right width=64><a href="javascript:validateForm(frmQCTemps);"><img src="images/add_button.gif" width=60 height=20 border=0></a></td>
		  <td align=right width=64><a href="javascript:OpenPrintBox('quote_bearingorder.asp?intQCTempLogId=<%=intQCTempLogId%>');"><img src="images/print_button.gif" width=60 height=20 border=0></a></td>
		  <td align=right width=64><a href="javascript:RefreshCurrentPage();"><img src="images/refresh_button.gif" width=60 height=20 border=0></a></td></tr>
	</table></td></tr>

 </table>
 </form>
 
 </td></tr>  <!-- End Middle Title Bar -->
 
 <%end if%>
</body>
</html>
<%
if not( rsGetQCTempLogByQCTempLogId is nothing ) then
	if( rsGetQCTempLogByQCTempLogId.State = adStateOpen ) then
		rsGetQCTempLogByQCTempLogId.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetQCTempLogByQCTempLogId				= nothing
set cmdGetQCTempLogByQCTempLogId			= nothing
set cmdEditQCTempLog											= nothing
set cnnSQLConnection									= nothing
%>
