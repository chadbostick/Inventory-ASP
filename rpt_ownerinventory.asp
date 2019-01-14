<%option explicit

dim cnnSQLConnection
dim rsReportInventoryTransactions, cmdReportInventoryTransactions
dim strPrint, strOwner, intProductOwnerId, strOrderBy

set cnnSQLConnection				= nothing
set rsReportInventoryTransactions	= nothing
set cmdReportInventoryTransactions	= nothing
set strPrint = nothing
set strOwner = nothing
set strOrderBy = nothing

strPrint = Request.QueryString( "Print" )
intProductOwnerId = Request.QueryString( "id" )
strOwner = Request.QueryString( "owner" )
strOrderBy = Request.QueryString( "orderby" )

set cnnSQLConnection				= Server.CreateObject( "ADODB.Connection" )
set cmdReportInventoryTransactions	= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

cmdReportInventoryTransactions.ActiveConnection = cnnSQLConnection
cmdReportInventoryTransactions.CommandText = "{call spReportInventoryTransactions( ?,? ) }"
cmdReportInventoryTransactions.Parameters.Append cmdReportInventoryTransactions.CreateParameter( "@ProductOwnerId", adInteger, adParamInput, 4 )
cmdReportInventoryTransactions.Parameters.Append cmdReportInventoryTransactions.CreateParameter( "@OrderBy", adVarChar, adParamInput, 50 )
cmdReportInventoryTransactions.Parameters( "@ProductOwnerId" )	= intProductOwnerId
cmdReportInventoryTransactions.Parameters( "@OrderBy" )	= strOrderBy

on error resume next
set rsReportInventoryTransactions = cmdReportInventoryTransactions.Execute
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
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/reports.css" TITLE="tables_css">
  <script type="text/javascript" src="js/common.js"></script>
  <script language="Javascript">
  function OpenProduct( strProductId )
	{
	  OpenWindow( 'product_details.asp?intProductId='+strProductId,'Centerline','width=660,height=480,left='+GetCenterX( 660 )+',top='+GetCenterY( 480 )+',screenX='+GetCenterX( 660 )+',screenY='+GetCenterY( 480 )+',scrollbars=yes,dependent=yes' );
	}
  </script>
</HEAD>

<BODY<%if strPrint = "true" then%> onload="javascript: if (window.print) window.print();"<%end if%>>
  <table align=left width=800>
  
    <tr><td><table width=800 class="form_header">
      <tr><td valign=top><img src="images/centerline_logo_s.jpg" width=149 height=73></td>
       <td align=right><table>
        <tr><td class="internal_report_title"><%= strOwner %> Inventory Report</td></tr>
        </table></td></tr>
    </table></td></tr>
    
    <tr><td class="internal_report_info"><table cellspacing=0 cellpadding=4 width=800>
      
      <tr><td class="internal_report_label" align="center" width=50>Id</td>
        <td class="internal_report_label">Name</td>
        <td class="internal_report_label">Category</td>
        <td class="internal_report_label">Description</td>
        <td class="internal_report_label" align="right">Reorder Lvl</td>
        <td class="internal_report_label" align="right" width=40>Units On Hand</td>
        <td class="internal_report_label" align="right" width=40>Units On Order</td>
        <td class="internal_report_label" align="right" width=60>Cost/Value</td>
      </tr>
      
      <%if strPrint <> "true" then%>
      <tr><td colspan=9>
				<div style="height:420px; overflow: auto;" >
				<table width="100%" border=0 cellspacing = 0 cellpadding=4 ID="Table1">
			<%end if%>   
              
                
        <%if( not rsReportInventoryTransactions is nothing ) then
						if( rsReportInventoryTransactions.State = adStateOpen ) then
							while( not rsReportInventoryTransactions.EOF )%>
								<tr><td class="internal_report_group_name" align="center"><a href="javascript:OpenProduct('<%=rsReportInventoryTransactions( "ProductId" )%>');"><%=rsReportInventoryTransactions( "ProductId" )%></a></td>
								  <td class="internal_report_group_name"><%=rsReportInventoryTransactions( "PartNumber" )%></td>
								  <td class="internal_report_group_name"><%=rsReportInventoryTransactions( "CategoryName" )%></td>
								  <td class="internal_report_group_name"><%=rsReportInventoryTransactions( "ProductShortDescription" )%></td>
								  <td class="internal_report_group_name" align="right"><%=rsReportInventoryTransactions( "ReorderLevel" )%></td>
								  <td class="internal_report_group_name" align="right"><%=rsReportInventoryTransactions( "OnHand" )%></td>
								  <td class="internal_report_group_name" align="right"><%=rsReportInventoryTransactions( "OnOrder" )%></td>
								  <td class="internal_report_group_name" align="right"><%=FormatCurrency(rsReportInventoryTransactions( "UnitPrice" ))%></td>
								<%rsReportInventoryTransactions.MoveNext
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
if not( rsReportInventoryTransactions is nothing ) then
	if( rsReportInventoryTransactions.State = adStateOpen ) then
		rsReportInventoryTransactions.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsReportInventoryTransactions		= nothing
set cmdReportInventoryTransactions		= nothing
set cnnSQLConnection					= nothing
%>
