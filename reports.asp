<%option explicit

dim cnnSQLConnection
dim rsGetProductOwners, cmdGetProductOwners

set cnnSQLConnection						= nothing
set rsGetProductOwners					= nothing
set cmdGetProductOwners					= nothing

set cnnSQLConnection						= Server.CreateObject( "ADODB.Connection" )
set cmdGetProductOwners					= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient
'**********************************************************************************
cmdGetProductOwners.ActiveConnection = cnnSQLConnection
cmdGetProductOwners.CommandText = "{call spGetProductOwnerList}"

on error resume next
set rsGetProductOwners = cmdGetProductOwners.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
						
on error goto 0
'**********************************************************************************
%>
<html>
	<head>
		<title>Centerline - Reports</title>
		<LINK REL="STYLESHEET" TYPE="text/css" HREF="css/default.css" TITLE="default_css">
			<LINK REL="STYLESHEET" TYPE="text/css" HREF="css/titles.css" TITLE="titles_css">
				<LINK REL="STYLESHEET" TYPE="text/css" HREF="css/forms.css" TITLE="forms_css">
					<SCRIPT language="Javascript" src="js/main_menu.js"></SCRIPT>
					<script type="text/javascript" src="js/common.js"></script>
					<script language="javascript">
			function ViewReport()
			{
				var nReportCount;

				for( nReportCount = 0; nReportCount < document.frmReports.rdoMRPReports.length; nReportCount++ )
				{
					if( document.frmReports.rdoMRPReports[nReportCount].checked == true )
					{
						OpenWindow( document.frmReports.rdoMRPReports[nReportCount].value,'ViewReport','width=845,height=600,left='+GetCenterX( 845 )+',top='+GetCenterY( 600 )+',screenX='+GetCenterX( 845 )+',screenY='+GetCenterY( 600 )+',scrollbars=yes,dependent=yes' );
					}
				}
			}

			function PrintReport()
			{
				var nReportCount;

				for( nReportCount = 0; nReportCount < document.frmReports.rdoMRPReports.length; nReportCount++ )
				{
					if( document.frmReports.rdoMRPReports[nReportCount].checked == true )
					{
						OpenPrintBox( document.frmReports.rdoMRPReports[nReportCount].value + '?print=true'  );
					}
				}
			}
			
			function ViewInventory()
			{
				var nReportCount;
				var strOrderBy;				

				for( var i = 0; i < document.frmInventory.selOrderBy.length; i++ )
				{
					if( document.frmInventory.selOrderBy[i].selected == true )
					{
						strOrderBy = document.frmInventory.selOrderBy[i].value;
					}
				}

				for( nReportCount = 0; nReportCount < document.frmInventory.selProductOwner.length; nReportCount++ )
				{
					if( document.frmInventory.selProductOwner[nReportCount].selected == true )
					{
						OpenWindow( 'rpt_ownerinventory.asp?id=' + document.frmInventory.selProductOwner[nReportCount].value + '&owner=' + document.frmInventory.selProductOwner[nReportCount].text + '&orderby=' + strOrderBy,'ViewReport','width=845,height=600,left='+GetCenterX( 845 )+',top='+GetCenterY( 600 )+',screenX='+GetCenterX( 845 )+',screenY='+GetCenterY( 600 )+',scrollbars=yes,dependent=yes' );
					}
				}
			}
			
			function PrintInventory()
			{
				var nReportCount;
				var strOrderBy;				

				for( var i = 0; i < document.frmInventory.selOrderBy.length; i++ )
				{
					if( document.frmInventory.selOrderBy[i].selected == true )
					{
						strOrderBy = document.frmInventory.selOrderBy[i].value;
					}
				}

				for( nReportCount = 0; nReportCount < document.frmInventory.selProductOwner.length; nReportCount++ )
				{
					if( document.frmInventory.selProductOwner[nReportCount].selected == true )
					{
						OpenWindow( 'rpt_ownerinventory.asp?print=true&id=' + document.frmInventory.selProductOwner[nReportCount].value + '&owner=' + document.frmInventory.selProductOwner[nReportCount].text + '&orderby=' + strOrderBy,'ViewReport','width=845,height=600,left='+GetCenterX( 845 )+',top='+GetCenterY( 600 )+',screenX='+GetCenterX( 845 )+',screenY='+GetCenterY( 600 )+',scrollbars=yes,dependent=yes' );
					}
				}
			}
	</script>
	</head>
	<body topmargin="2" leftmargin="2" marginwidth="2" marginheight="2">
		<table border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td><table border="0" cellpadding="0" cellspacing="0" width="770">
						<tr>
							<td><a href="index.html"><IMG SRC="images/menu/menu_top_left.gif" width="202" height="54" border="0"></a></td>
							<td><IMG SRC="images/menu/menu_top_middle.gif" width="460" height="54" border="0"></td>
							<td><IMG SRC="images/menu/menu_icon_reports.gif" width="108" height="54" border="0"></td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td><table border="0" cellpadding="0" cellspacing="0" width="770">
						<tr>
							<td><a href="index.html"><IMG SRC="images/menu/menu_left.gif" width="202" height="25" border="0"></a></td>
							<td><a href="customer.asp" onmouseOver="imgOn('customer')" onMouseOut="imgOff('customer')">
									<IMG name="customer" SRC="images/menu/menu_customer_off.gif" width="94" height="25" border="0"></a></td>
							<td><a href="project.asp" onmouseOver="imgOn('project')" onMouseOut="imgOff('project')">
									<IMG name="project" SRC="images/menu/menu_project_off.gif" width="76" height="25" border="0"></a></td>
							<td><a href="spindle.asp" onmouseOver="imgOn('spindle')" onMouseOut="imgOff('spindle')">
									<IMG name="spindle" SRC="images/menu/menu_spindle_off.gif" width="80" height="25" border="0"></a></td>
							<td><a href="workorder.asp" onmouseOver="imgOn('work_order')" onMouseOut="imgOff('work_order')">
									<IMG name="work_order" SRC="images/menu/menu_work_order_off.gif" width="103" height="25" border="0"></a></td>
							<td><a href="quote.asp" onmouseOver="imgOn('quote')" onMouseOut="imgOff('quote')"> <IMG name="quote" SRC="images/menu/menu_quote_off.gif" width="70" height="25" border="0"></a></td>
							<td><a href="po.asp" onmouseOver="imgOn('po')" onMouseOut="imgOff('po')"> <IMG name="po" SRC="images/menu/menu_po_off.gif" width="64" height="25" border="0"></a></td>
							<td><IMG name="reports" SRC="images/menu/menu_reports_on.gif" width="81" height="25" border="0"></td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td><table border="0" cellpadding="0" cellspacing="0" width="770">
						<tr>
							<td><a href="index.html"><IMG SRC="images/menu/menu_bottom_left.gif" width="202" height="39" border="0"></a></td>
							<td><IMG SRC="images/menu/menu_bottom_right.gif" width="568" height="39" border="0"></td>
					</table>
					<!-- a little spacer betwen header and content -->
					&nbsp; 
					<!-- Page Contents -->
					<table width="700" border="0" cellpadding="0" cellspacing="0">
						<!--Reports -->
						<tr>
							<td align="right">
								<table width="600" border="0" cellpadding="0" cellspacing="0" align="right">
									<form id="frmReports" name="frmReports" action="" method="post">
										<!--Top Title Bar -->
										<tr>
											<td class="pagetitle">
												<table class="sectiontitle" width="100%" cellpadding="0" cellspacing="0">
													<tr>
														<td colspan="3"><IMG SRC="images/title_top.gif" width="100%" height="6"></td>
													</tr>
													<tr>
														<td class="sectiontitle" width="160">MRP Reports</td>
														<td align="left" width="300"><IMG SRC="images/title_right.gif" width="23" height="25"></td>
														<td align="right" width="140"><table width="140" cellspacing="0" cellpadding="0">
																<tr>
																	<td width="76" align="right"><a href="javascript:ViewReport();"><IMG SRC="images/view_button.gif" width="60" height="20" border="0"></a></td>
																	<td width="64" align="right"><a href="javascript:PrintReport();"><IMG SRC="images/print_button.gif" width="60" height="20" border="0"></a></td>
																</tr>
															</table>
														</td>
													</tr>
												</table>
											</td>
										</tr>
										<!-- End Top Title Bar -->
										<tr>
											<td align="center"><table class="form" width="538" border="0" cellpadding="2" cellspacing="4" align="middle">
													<tr>
														<td class="generic_input"><input type="radio" name="rdoMRPReports" value="rpt_openprojects.asp">Open 
															Projects</td>
													</tr>
													<tr>
														<td class="generic_input"><input type="radio" name="rdoMRPReports" value="rpt_productneed.asp">Products 
															Needed</td>
													</tr>
													<tr>
														<td class="generic_input"><input type="radio" name="rdoMRPReports" value="rpt_openpo.asp">Open 
															Purchase Orders</td>
													</tr>
													<tr>
														<td class="generic_input"><input type="radio" name="rdoMRPReports" value="rpt_openworkorders.asp">Open 
															Work Orders</td>
													</tr>
													<tr>
														<td class="generic_input"><input type="radio" name="rdoMRPReports" value="rpt_openquotes.asp">Open 
															Quotes</td>
													</tr>
												</table>
											</td>
										</tr>
								</table>
							</td>
						</tr>
						<!-- End Reports -->
						<tr>
							<td>&nbsp;</td>
						</tr>
						<!--Manual Forms -->
						<tr>
							<td align="right">
								<table width="600" border="0" cellpadding="0" cellspacing="0" align="right" ID="Table1">
									<!--Top Title Bar -->
									<tr>
										<td class="pagetitle">
											<table class="sectiontitle" width="100%" cellpadding="0" cellspacing="0" ID="Table2">
												<tr>
													<td colspan="3"><IMG SRC="images/title_top.gif" width="100%" height="6"></td>
												</tr>
												<tr>
													<td class="sectiontitle" width="150">Manual Forms</td>
													<td align="left" width="390"><IMG SRC="images/title_right.gif" width="23" height="25"></td>
													<td width="60" align="right"><a href="javascript:ViewReport();"><IMG SRC="images/view_button.gif" width="60" height="20" border="0"></a></td>
												</tr>
											</table>
										</td>
									</tr>
									<!-- End Top Title Bar -->
									<tr>
										<td align="center"><table class="form" width="538" border="0" cellpadding="2" cellspacing="4" align="middle" ID="Table3">
												<tr>
													<td class="generic_input"><input type="radio" name="rdoMRPReports" value="rpt_receivinginspectionsheet.asp" ID="Radio1">Receiving 
														Inspection Sheet</td>
												</tr>
												<tr>
													<td class="generic_input"><input type="radio" name="rdoMRPReports" value="rpt_printspindlerepair.asp" ID="Radio2">Spindle 
														Repair Information Sheet</td>
												</tr>
											</table>
										</td>
									</tr>
									</form>
								</table>
							</td>
						</tr>
						<!-- End Manual Forms -->
						<tr>
							<td>&nbsp;</td>
						</tr>
						<!--Manual Forms -->
						<tr>
							<td align="right">
								<form id="Form1" name="frmInventory" action="" method="post">
									<table width="600" border="0" cellpadding="0" cellspacing="0" align="right" ID="Table4">
										<!--Top Title Bar -->
										<tr>
											<td class="pagetitle">
												<table class="sectiontitle" width="100%" cellpadding="0" cellspacing="0" ID="Table5">
													<tr>
														<td colspan="3"><IMG SRC="images/title_top.gif" width="100%" height="6"></td>
													</tr>
													<tr>
														<td class="sectiontitle" width="160">Inventory Report</td>
														<td align="left" width="300"><IMG SRC="images/title_right.gif" width="23" height="25"></td>
														<td align="right" width="140"><table width="140" cellspacing="0" cellpadding="0" ID="Table7">
																<tr>
																	<td width="76" align="right"><a href="javascript:ViewInventory();"><IMG SRC="images/view_button.gif" width="60" height="20" border="0"></a></td>
																	<td width="64" align="right"><a href="javascript:PrintInventory();"><IMG SRC="images/print_button.gif" width="60" height="20" border="0"></a></td>
																</tr>
															</table>
														</td>
													</tr>
												</table>
											</td>
										</tr>
										<!-- End Top Title Bar -->
										<tr>
											<td align="center"><table class="form" width="538" border="0" cellpadding="2" cellspacing="4" align="middle" ID="Table6">
													<tr>
														<td class="label_1">Owner</td>
														<td class="input_1">
   														<select class="forminput" name="selProductOwner" size=1>
   														<% if not( rsGetProductOwners is nothing ) then
   															if( rsGetProductOwners.State = adStateOpen ) then
   																while( not rsGetProductOwners.EOF )%>
   																	<option value="<%=rsGetProductOwners( "ProductOwnerId" )%>"><%=rsGetProductOwners( "ProductOwner" )%></option>
   																	<%rsGetProductOwners.MoveNext
   																wend
   															end if
   														end if
   														%>
   														</select>
   													</td>
   												</tr>
   												<tr>
   													<td class="label_1">Order By</td>
														<td class="input_1">															
   														<select class="forminput" name="selOrderBy" size=1>
   															<option value="ProductID">ID</option>
   															<option value="ProductName">Name</option>
   															<option value="CategoryName">Category</option>
   															<option value="ReorderLevel">Reorder Level</option>
   															<option value="OnHand">On Hand</option>
   															<option value="OnOrder">On Order</option>
   														</select>
														</td>
													</tr>
												</table>
											</td>
										</tr>
								</form>
					</table>
				</td>
			</tr>
			<!-- End Manual Forms -->
		</table>
		</td> </tr> </table>
	</body>
</html>
