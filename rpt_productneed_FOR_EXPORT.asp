<%option explicit

dim cnnSQLConnection
dim rsReportProductsNeeded, cmdReportProductsNeeded
dim strPrint

set cnnSQLConnection				= nothing
set rsReportProductsNeeded	= nothing
set cmdReportProductsNeeded	= nothing
set strPrint = nothing

strPrint = Request.QueryString( "Print" )

set cnnSQLConnection				= Server.CreateObject( "ADODB.Connection" )
set cmdReportProductsNeeded	= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

cmdReportProductsNeeded.ActiveConnection	= cnnSQLConnection

cmdReportProductsNeeded.CommandText = "{call spReportProductsNeeded }"

on error resume next
set rsReportProductsNeeded = cmdReportProductsNeeded.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if

on error goto 0%>
<HTML>
<TITLE>Products</TITLE>
  <script type="text/javascript" src="js/common.js"></script>
  <script language="Javascript">
  function OpenProduct( strProductId )
	{
	  OpenWindow( 'product_details.asp?intProductId='+strProductId,'Centerline','width=660,height=480,left='+GetCenterX( 660 )+',top='+GetCenterY( 480 )+',screenX='+GetCenterX( 660 )+',screenY='+GetCenterY( 480 )+',scrollbars=yes,dependent=yes' );
	}
  </script>
</HEAD>

<BODY<%if strPrint = "true" then%> onload="javascript: if (window.print) window.print();"<%end if%>>
  <table align=left>
    
    <tr><td class="internal_report_info"><table cellspacing=0 cellpadding=4 width=800>
      
      <tr><td class="internal_report_label" align="center" width=50>Id</td>
        <td class="internal_report_label" width=130>Part Number</td>
        <td class="internal_report_label" width=500>Product Description</td>
        <td class="internal_report_label" align="right" width=40>Reorder Level</td>
        <td class="internal_report_label" align="right" width=40>On Hand</td>
        <td class="internal_report_label" align="right" width=40>On Order</td>
      </tr>
      
      <%if strPrint <> "true" then%>
      <tr><td colspan=6>

				<table width="100%" border=0 cellspacing = 0 cellpadding=4>
			<%end if%>
        
        <%if( not rsReportProductsNeeded is nothing ) then
						if( rsReportProductsNeeded.State = adStateOpen ) then
							while( not rsReportProductsNeeded.EOF )%>
								<tr><td class="internal_report_group_name" align="center" valign="baseline" width=50><a href="product_details.asp?intProductId=<%=rsReportProductsNeeded( "ProductId" )%>"><%=rsReportProductsNeeded( "ProductId" )%></a></td>
								  <td class="internal_report_group_name" valign="baseline" width=130><%=rsReportProductsNeeded( "PartNumber" )%></td>
								  <td class="internal_report_group_name" valign="baseline" width=500><%=rsReportProductsNeeded( "ProductShortDescription" )%></td>
								  <td class="internal_report_group_name" valign="baseline" align="right" width=40><%=rsReportProductsNeeded( "ReorderLevel" )%></td>
								  <td class="internal_report_group_name" valign="baseline" align="right" width=40><%=rsReportProductsNeeded( "OnHand" )%></td>
								  <td class="internal_report_group_name" valign="baseline" align="right" width=40><%=rsReportProductsNeeded( "OnOrder" )%></td>
								</tr>
								<%rsReportProductsNeeded.MoveNext
							wend
					end if
				end if%> 
			<%if strPrint <> "true" then%>
				</table>
			</td></tr>
			<%end if%>
    </table></td></tr>
  </table>
</BODY>
</HTML>
<%
if not( rsReportProductsNeeded is nothing ) then
	if( rsReportProductsNeeded.State = adStateOpen ) then
		rsReportProductsNeeded.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsReportProductsNeeded		= nothing
set cmdReportProductsNeeded		= nothing
set cnnSQLConnection					= nothing
%>
