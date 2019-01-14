if (document.images) {
  map_customer_norm = new Image();
  map_customer_norm.src = "images/map/customer.gif";
  map_customer_add = new Image();
  map_customer_add.src = "images/map/customer_add.gif";
  map_customer_view = new Image();
  map_customer_view.src = "images/map/customer_view.gif";
  mid_mid_map_customer_norm = new Image();
  mid_mid_map_customer_norm.src = "images/map/mid_mid_customer.gif";
  mid_mid_map_customer_add = new Image();
  mid_mid_map_customer_add.src = "images/map/mid_mid_customer_add.gif";
  mid_mid_map_customer_view = new Image();
  mid_mid_map_customer_view.src = "images/map/mid_mid_customer_view.gif";
  
  map_project_norm = new Image();
  map_project_norm.src = "images/map/project.gif";
  map_project_add = new Image();
  map_project_add.src = "images/map/project_add.gif";
  map_project_view = new Image();
  map_project_view.src = "images/map/project_view.gif";
  mid_mid_map_project_norm = new Image();
  mid_mid_map_project_norm.src = "images/map/mid_mid_project.gif";
  mid_mid_map_project_add = new Image();
  mid_mid_map_project_add.src = "images/map/mid_mid_project_add.gif";
  mid_mid_map_project_view = new Image();
  mid_mid_map_project_view.src = "images/map/mid_mid_project_view.gif";
  
  map_spindle_norm = new Image();
  map_spindle_norm.src = "images/map/spindle.gif";
  map_spindle_add = new Image();
  map_spindle_add.src = "images/map/spindle_add.gif";
  map_spindle_view = new Image();
  map_spindle_view.src = "images/map/spindle_view.gif";
  mid_mid_map_spindle_norm = new Image();
  mid_mid_map_spindle_norm.src = "images/map/mid_mid_spindle.gif";
  mid_mid_map_spindle_add = new Image();
  mid_mid_map_spindle_add.src = "images/map/mid_mid_spindle_add.gif";
  mid_mid_map_spindle_view = new Image();
  mid_mid_map_spindle_view.src = "images/map/mid_mid_spindle_view.gif";
  
  map_po_norm = new Image();
  map_po_norm.src = "images/map/po.gif";
  map_po_add = new Image();
  map_po_add.src = "images/map/po_add.gif";
  map_po_view = new Image();
  map_po_view.src = "images/map/po_view.gif";
  mid_mid_map_po_norm = new Image();
  mid_mid_map_po_norm.src = "images/map/mid_mid_po.gif";
  mid_mid_map_po_add = new Image();
  mid_mid_map_po_add.src = "images/map/mid_mid_po_add.gif";
  mid_mid_map_po_view = new Image();
  mid_mid_map_po_view.src = "images/map/mid_mid_po_view.gif";
  
  map_quote_norm = new Image();
  map_quote_norm.src = "images/map/quote.gif";
  map_quote_add = new Image();
  map_quote_add.src = "images/map/quote_add.gif";
  map_quote_view = new Image();
  map_quote_view.src = "images/map/quote_view.gif";
  mid_mid_map_quote_norm = new Image();
  mid_mid_map_quote_norm.src = "images/map/mid_mid_quote.gif";
  mid_mid_map_quote_add = new Image();
  mid_mid_map_quote_add.src = "images/map/mid_mid_quote_add.gif";
  mid_mid_map_quote_view = new Image();
  mid_mid_map_quote_view.src = "images/map/mid_mid_quote_view.gif";
  
  map_work_order_norm = new Image();
  map_work_order_norm.src = "images/map/work_order.gif";
  map_work_order_add = new Image();
  map_work_order_add.src = "images/map/work_order_add.gif";
  map_work_order_view = new Image();
  map_work_order_view.src = "images/map/work_order_view.gif";
  mid_mid_map_work_order_norm = new Image();
  mid_mid_map_work_order_norm.src = "images/map/mid_mid_work_order.gif";
  mid_mid_map_work_order_add = new Image();
  mid_mid_map_work_order_add.src = "images/map/mid_mid_work_order_add.gif";
  mid_mid_map_work_order_view = new Image();
  mid_mid_map_work_order_view.src = "images/map/mid_mid_work_order_view.gif";
  
  map_reports_norm = new Image();
  map_reports_norm.src = "images/map/reports.gif";
  map_reports_add = new Image();
  map_reports_add.src = "images/map/reports_add.gif";
  map_reports_view = new Image();
  map_reports_view.src = "images/map/reports_view.gif";
  mid_mid_map_reports_norm = new Image();
  mid_mid_map_reports_norm.src = "images/map/mid_mid_reports.gif";
  mid_mid_map_reports_add = new Image();
  mid_mid_map_reports_add.src = "images/map/mid_mid_reports_add.gif";
  mid_mid_map_reports_view = new Image();
  mid_mid_map_reports_view.src = "images/map/mid_mid_reports_view.gif";
  
  map_admin_norm = new Image();
  map_admin_norm.src = "images/map/admin.gif";
  map_admin_add = new Image();
  map_admin_add.src = "images/map/admin_add.gif";
  map_admin_view = new Image();
  map_admin_view.src = "images/map/admin_view.gif";
  mid_mid_map_admin_norm = new Image();
  mid_mid_map_admin_norm.src = "images/map/mid_mid_admin.gif";
  mid_mid_map_admin_add = new Image();
  mid_mid_map_admin_add.src = "images/map/mid_mid_admin_add.gif";
  mid_mid_map_admin_view = new Image();
  mid_mid_map_admin_view.src = "images/map/mid_mid_admin_view.gif";
  
  
  mid_mid_norm = new Image();
  mid_mid_norm.src = "images/map/mid_mid.gif";
}


function rollItem(whichItem, itemState) {
  if (document.images) {
    if (itemState == "none") {
      document[whichItem].src = eval(whichItem + "_norm.src");
      document["mid_mid"].src = mid_mid_norm.src;
    } else {
      document[whichItem].src = eval(whichItem + "_" + itemState + ".src");
      document["mid_mid"].src = eval("mid_mid_" + whichItem + "_" + itemState + ".src");
    }
  }
}