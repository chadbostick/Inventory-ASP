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
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/reports.css" TITLE="tables_css">
  <script type="text/javascript" src="js/common.js"></script>
  <script language="Javascript">
  function OpenPurchaseOrder( strPurchaseOrderId )
	{
	  OpenWindow( 'po_details.asp?intPurchaseOrderId='+strPurchaseOrderId,'CenterlinePO'+strPurchaseOrderId,'width=657,height=440,left='+GetCenterX( 657 )+',top='+GetCenterY( 440 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 440 )+',scrollbars=yes,dependent=yes' );
	}
	</script>
</HEAD>

<BODY<%if strPrint = "true" then%> onload="javascript: if (window.print) window.print();"<%end if%>>
  <table width=600 align=left>
  
    <tr><td><table width=600 class="form_header">
      <tr><td valign=top><img src="images/centerline_logo_s.jpg" width=149 height=73></td>
       <td align=right><table>
        <tr><td class="internal_report_title">Open Purchase Orders</td></tr>
        </table></td></tr>
    </table></td></tr>
    
    <tr><td class="internal_report_info"><table cellspacing=0 cellpadding=4 width=800>
      
      <tr><td class="internal_report_label" align="center" width=50>Id</td>
        <td class="internal_report_label" width=80>PO Number</td>
        <td class="internal_report_label" width=380>PO Description</td>
        <td class="internal_report_label" width=70>Ordered</td>
        <td class="internal_report_label" width=70>Required</td>
        <td class="internal_report_label" width=150>Supplier</td>
      </tr>
      
      <%if strPrint <> "true" then%>
      <tr><td colspan=6>
				<div><!--  style="height:420px; overflow: auto;"  //-->
				<table width="100%" border=0 cellspacing = 0 cellpadding=4 ID="Table1">
			<%end if%>
          
        <%if( not rsReportOpenPurchaseOrders is nothing ) then
						if( rsReportOpenPurchaseOrders.State = adStateOpen ) then
							while( not rsReportOpenPurchaseOrders.EOF )%>
								<tr><td class="internal_report_group_name" width=50 align="center" valign="baseline"><a href="javascript:OpenPurchaseOrder( '<%=rsReportOpenPurchaseOrders( "PurchaseOrderId" )%>' );"><%=rsReportOpenPurchaseOrders( "PurchaseOrderId" )%></a></td>
								  <td class="internal_report_group_name" width=80 valign="baseline"><%=rsReportOpenPurchaseOrders( "PurchaseOrderNumber" )%></td>
								  <td class="internal_report_group_name" width=380 valign="baseline"><%=rsReportOpenPurchaseOrders( "PurchaseOrderDescription" )%></td>
								  <td class="internal_report_group_name" width=70 valign="baseline"><%=rsReportOpenPurchaseOrders( "OrderDate" )%></td>
								  <%if( IsNull( rsReportOpenPurchaseOrders( "DateRequired" ) ) ) then%>
								  <td class="internal_report_group_name" width=70 valign="baseline">&nbsp;</td>
								  <%else%>
								  <td class="internal_report_group_name" width=70 valign="baseline"><%=rsReportOpenPurchaseOrders( "DateRequired" )%></td>
								  <%end if%>
								  <td class="internal_report_group_name" width=150 valign="baseline"><%=rsReportOpenPurchaseOrders( "SupplierName" )%></td>
								<%rsReportOpenPurchaseOrders.MoveNext
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
