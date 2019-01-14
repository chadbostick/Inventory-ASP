<%option explicit

dim cnnSQLConnection
dim rsReportOpenWorkOrders, cmdReportOpenWorkOrders
dim strPrint

set cnnSQLConnection				= nothing
set rsReportOpenWorkOrders	= nothing
set cmdReportOpenWorkOrders	= nothing
set strPrint = nothing

strPrint = Request.QueryString( "Print" )

set cnnSQLConnection				= Server.CreateObject( "ADODB.Connection" )
set cmdReportOpenWorkOrders	= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

cmdReportOpenWorkOrders.ActiveConnection	= cnnSQLConnection
cmdReportOpenWorkOrders.CommandText				= "{call spReportOpenWorkOrders }"

on error resume next
set rsReportOpenWorkOrders = cmdReportOpenWorkOrders.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if

on error goto 0%>
<HTML>
<TITLE>Open Work Orders</TITLE>
  <script type="text/javascript" src="js/common.js"></script>
</HEAD>

<BODY<%if strPrint = "true" then%> onload="javascript: if (window.print) window.print();"<%end if%>>
<script type="text/javascript" src="/js/tooltips.js"></script>
  <table align=left width="100%">
    
    <tr><td class="internal_report_info"><table cellspacing=0 cellpadding=4>
      
      <tr><td width="11%" class="internal_report_label">WO#</td>
        <td width="11%" class="internal_report_label">Customer</td>
        <td width="11%" class="internal_report_label">Spindle Type</td>
        <td width="11%" class="internal_report_label">Log In Date</td>
        <td width="11%" class="internal_report_label">Prty</td>
        <td width="11%" class="internal_report_label">
			<span style="color:green" onmouseover="Tip('changed from Quoted RTS after conversation with customer. This is when the Customer expects the repair to ship if not the Quoted RTS')">Promised RTS</span>
		</td>
        <td width="11%" class="internal_report_label">
			<span style="color:green" onmouseover="Tip('Earliest Possible Assembly should be filled out according to supply time for the bearings, parts, subwork etc. As POs are adjusted/closed, EPAs should be updated')">EPA Date</span>
		</td>
        <td width="11%" class="internal_report_label">
			<span style="color:green" onmouseover="Tip('Ready To Ship date set after quote approval based on quoted delivery')">Quoted RTS</span>
		</td>
        <td width="11%" class="internal_report_label">Recvd Date</td>
      </tr>
      
      <%if strPrint <> "true" then%>

			<%end if%>          
              
        <%if( not rsReportOpenWorkOrders is nothing ) then
						if( rsReportOpenWorkOrders.State = adStateOpen ) then
							while( not rsReportOpenWorkOrders.EOF )%>
								<tr><td class="internal_report_group_name"><a href="workorder_details.asp?intWorkOrderId=<%=rsReportOpenWorkOrders( "WorkOrderId" )%>"><%=rsReportOpenWorkOrders( "WorkOrderNumber" )%></a></td>
								  <td class="internal_report_group_name"><%=Left(rsReportOpenWorkOrders( "Customer" ),15)%></td>
								  <td class="internal_report_group_name"><%=Left(rsReportOpenWorkOrders( "SpindleType" ),15)%></td>
								  <%if( IsNull( rsReportOpenWorkOrders( "DateIn" ) ) ) then%>
								  <td class="internal_report_group_name">&nbsp;</td>
								  <%else%>
								  <td class="internal_report_group_name"><%=rsReportOpenWorkOrders( "DateIn" )%></td>
								  <%end if%>
								  <td class="internal_report_group_name"><b><%=rsReportOpenWorkOrders( "Priority" )%></b></td>
								  <%if( IsNull( rsReportOpenWorkOrders( "PromiseDate" ) ) ) then%>
								  <td class="internal_report_group_name">&nbsp;</td>
								  <%else%>
								  <td class="internal_report_group_name"><%=rsReportOpenWorkOrders( "PromiseDate" )%></td>
								  <%end if%>
								  <%if( IsNull( rsReportOpenWorkOrders( "Date" ) ) ) then%>
								  <td class="internal_report_group_name">&nbsp;</td>
								  <%else%>
								  <td class="internal_report_group_name"><%=rsReportOpenWorkOrders( "Date" )%></td>
								  <%end if%>
								  <%if( IsNull( rsReportOpenWorkOrders( "ExpectedDelDate" ) ) ) then%>
								  <td class="internal_report_group_name">&nbsp;</td>
								  <%else%>
								  <td class="internal_report_group_name"><%=rsReportOpenWorkOrders( "ExpectedDelDate" )%></td>
								  <%end if%>
								  <%if( IsNull( rsReportOpenWorkOrders( "DateRec" ) ) ) then%>
								  <td class="internal_report_group_name">&nbsp;</td>
								  <%else%>
								  <td class="internal_report_group_name"><%=rsReportOpenWorkOrders( "DateRec" )%></td>
								  <%end if%>
								<%rsReportOpenWorkOrders.MoveNext
							wend
					end if
				end if%>  
				
			<%if strPrint <> "true" then%>
				</table>
			<%end if%> 
			
    </table></td></tr>
  </table>
</BODY>
</HTML>
<%
if not( rsReportOpenWorkOrders is nothing ) then
	if( rsReportOpenWorkOrders.State = adStateOpen ) then
		rsReportOpenWorkOrders.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsReportOpenWorkOrders	= nothing
set cmdReportOpenWorkOrders	= nothing
set cnnSQLConnection				= nothing
%>
