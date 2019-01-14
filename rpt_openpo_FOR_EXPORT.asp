<%option explicit

dim cnnSQLConnection
dim rsReportOpenPurchaseOrders, cmdReportOpenPurchaseOrders
dim strPrint

set cnnSQLConnection						= nothing
set rsReportOpenPurchaseOrders	= nothing
set cmdReportOpenPurchaseOrders	= nothing
set strPrint = nothing

strPrint = Request.QueryString( "Print" )

set cnnSQLConnection						= Server.CreateObject( "ADODB.Connection" )
set cmdReportOpenPurchaseOrders	= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

cmdReportOpenPurchaseOrders.ActiveConnection	= cnnSQLConnection
cmdReportOpenPurchaseOrders.CommandText				= "{call spReportOpenPurchaseOrders }"

on error resume next
set rsReportOpenPurchaseOrders = cmdReportOpenPurchaseOrders.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if

on error goto 0%>
<HTML>
<TITLE>Open Purchase Orders</TITLE>
  <script type="text/javascript" src="js/common.js"></script>
</HEAD>

<BODY>

<%if strPrint <> "true" then%>

<table width="100%" border=0 cellspacing = 0 cellpadding=4 ID="Table1">
	<tr><td class="internal_report_label" align="center">Id</td>
		<td class="internal_report_label">PO #</td>
    	<td class="internal_report_label" width=380>PO Description</td>
    	<td class="internal_report_label">Ordered</td>
    	<td class="internal_report_label">Required</td>
    	<td class="internal_report_label">Supplier</td>
	</tr>

<%end if%>
          
        <%if( not rsReportOpenPurchaseOrders is nothing ) then
						if( rsReportOpenPurchaseOrders.State = adStateOpen ) then
							while( not rsReportOpenPurchaseOrders.EOF )%>
								<tr><td class="poid" align="center" valign="baseline"><a href="po_details.asp?intPurchaseOrderId=<%=rsReportOpenPurchaseOrders( "PurchaseOrderId" )%>"><%=rsReportOpenPurchaseOrders( "PurchaseOrderId" )%></a></td>
								  <td class="ponum" valign="baseline"><%=rsReportOpenPurchaseOrders( "PurchaseOrderNumber" )%></td>
								  <td class="podesc" width=380 valign="baseline"><%=rsReportOpenPurchaseOrders( "PurchaseOrderDescription" )%></td>
								  <td class="ordate" valign="baseline"><%=rsReportOpenPurchaseOrders( "OrderDate" )%></td>
								  <%if( IsNull( rsReportOpenPurchaseOrders( "DateRequired" ) ) ) then%>
								  <td valign="baseline">&nbsp;</td>
								  <%else%>
								  <td class="required" valign="baseline"><%=rsReportOpenPurchaseOrders( "DateRequired" )%></td>
								  <%end if%>
								  <td class="supplier" valign="baseline"><%=rsReportOpenPurchaseOrders( "SupplierName" )%></td>
								<%rsReportOpenPurchaseOrders.MoveNext
							wend
					end if
				end if%>  
				
			<%if strPrint <> "true" then%>
	</tr>
			<%end if%> 
  </table>
</BODY>
</HTML>
<%
if not( rsReportOpenPurchaseOrders is nothing ) then
	if( rsReportOpenPurchaseOrders.State = adStateOpen ) then
		rsReportOpenPurchaseOrders.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsReportOpenPurchaseOrders		= nothing
set cmdReportOpenPurchaseOrders		= nothing
set cnnSQLConnection							= nothing
%>
