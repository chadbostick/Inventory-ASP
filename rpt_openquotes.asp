<%option explicit

dim cnnSQLConnection
dim rsReportOpenQuotes, cmdReportOpenQuotes
dim strPrint

set cnnSQLConnection				= nothing
set rsReportOpenQuotes	= nothing
set cmdReportOpenQuotes	= nothing
set strPrint = nothing

strPrint = Request.QueryString( "Print" )

set cnnSQLConnection				= Server.CreateObject( "ADODB.Connection" )
set cmdReportOpenQuotes	= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

cmdReportOpenQuotes.ActiveConnection	= cnnSQLConnection
cmdReportOpenQuotes.CommandText				= "{call spReportOpenQuotes }"

on error resume next
set rsReportOpenQuotes = cmdReportOpenQuotes.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if

on error goto 0%>
<HTML>
<TITLE>Open Quotes</TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/reports.css" TITLE="tables_css">
  <script type="text/javascript" src="js/common.js"></script>
  <script language="Javascript">
  function OpenQuote( strQuoteId )
	{
	  OpenWindow( 'quote_details.asp?intQuoteId='+strQuoteId,'CenterlineQuote'+strQuoteId,'width=657,height=440,left='+GetCenterX( 657 )+',top='+GetCenterY( 440 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 440 )+',scrollbars=yes,dependent=yes' );
	}
	
	function OpenWorkOrder( strWorkOrderId )
	{
	  OpenWindow( 'workorder_details.asp?intWorkOrderId='+strWorkOrderId,'CenterlineWorkOrder'+strWorkOrderId,'width=657,height=480,left='+GetCenterX( 657 )+',top='+GetCenterY( 480 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 480 )+',scrollbars=yes,dependent=yes' );
	}
	</script>
</HEAD>

<BODY<%if strPrint = "true" then%> onload="javascript: if (window.print) window.print();"<%end if%>>
  <table align=left width=800>
  
    <tr><td><table width=800 class="form_header">
      <tr><td valign=top><img src="images/centerline_logo_s.jpg" width=149 height=73></td>
       <td align=right><table>
        <tr><td class="internal_report_title">Open Quotes</td></tr>
        </table></td></tr>
    </table></td></tr>
    
    <tr><td class="internal_report_info"><table cellspacing=0 cellpadding=4 width=800>
      
      <tr><td class="internal_report_label" align="center" width=50>Id</td>
        <td class="internal_report_label" width="100">Work Order Number</td>
        <td class="internal_report_label" width="150">Customer</td>
        <td class="internal_report_label" width="380">Spindle Type</td>
        <td class="internal_report_label" width="60">Date Quoted</td>
      </tr>
      
      <%if strPrint <> "true" then%>
      <tr><td colspan=9>
				<div><!--  style="height:420px; overflow: auto;"  //-->
				<table width="100%" border=0 cellspacing = 0 cellpadding=4 ID="Table1">
			<%end if%>   
              
        <%if( not rsReportOpenQuotes is nothing ) then
						if( rsReportOpenQuotes.State = adStateOpen ) then
							while( not rsReportOpenQuotes.EOF )%>
								<tr><td class="internal_report_group_name" align="center" width="50"><a href="javascript:OpenQuote( '<%=rsReportOpenQuotes( "QuoteId" )%>' );"><%=rsReportOpenQuotes( "QuoteId" )%></a></td>
								  <td class="internal_report_group_name" width="100"><a href="javascript:OpenWorkOrder( '<%=rsReportOpenQuotes( "WorkOrderId" )%>' );"><%=rsReportOpenQuotes( "WorkOrderNumber" )%></a></td>
								  <td class="internal_report_group_name" width="150"><%=rsReportOpenQuotes( "Customer" )%></td>
								  <td class="internal_report_group_name" width="380"><%=rsReportOpenQuotes( "SpindleType" )%></td>
								  <%if( IsNull( rsReportOpenQuotes( "DateQuoted" ) ) ) then%>
								  <td class="internal_report_group_name" width="60">&nbsp;</td>
								  <%else%>
								  <td class="internal_report_group_name" width="60"><%=rsReportOpenQuotes( "DateQuoted" )%></td>
								  <%end if%>
								<%rsReportOpenQuotes.MoveNext
							wend
					end if
				end if%>  
				
			<%if strPrint <> "true" then%>
				</table></div>
			</td></tr>
			<%end if%> 
			
    </table></td></tr>
  </table>
</BODY>
</HTML>
<%
if not( rsReportOpenQuotes is nothing ) then
	if( rsReportOpenQuotes.State = adStateOpen ) then
		rsReportOpenQuotes.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsReportOpenQuotes	= nothing
set cmdReportOpenQuotes	= nothing
set cnnSQLConnection				= nothing
%>
