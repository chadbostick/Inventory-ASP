if (document.images) {
 //main menu
  // Active images
  customer_on = new Image();
  customer_on.src = "images/menu/menu_customer_on.gif";
  project_on = new Image();
  project_on.src = "images/menu/menu_project_on.gif";
  spindle_on = new Image();
  spindle_on.src = "images/menu/menu_spindle_on.gif";
  work_order_on = new Image();
  work_order_on.src = "images/menu/menu_work_order_on.gif";
  quote_on = new Image();
  quote_on.src = "images/menu/menu_quote_on.gif";
  po_on = new Image();
  po_on.src = "images/menu/menu_po_on.gif";
  reports_on = new Image();
  reports_on.src = "images/menu/menu_reports_on.gif";
  
  // Inactive images
  customer_off = new Image();
  customer_off.src = "images/menu/menu_customer_off.gif";
  project_off = new Image();
  project_off.src = "images/menu/menu_project_off.gif";
  spindle_off = new Image();
  spindle_off.src = "images/menu/menu_spindle_off.gif";
  work_order_off = new Image();
  work_order_off.src = "images/menu/menu_work_order_off.gif";
  quote_off = new Image();
  quote_off.src = "images/menu/menu_quote_off.gif";
  po_off = new Image();
  po_off.src = "images/menu/menu_po_off.gif";
  reports_off = new Image();
  reports_off.src = "images/menu/menu_reports_off.gif";
}


function imgOn(imgName) {
  if (document.images) {
    document[imgName].src = eval(imgName + "_on.src");
  }
}


function imgOff(imgName) {
  if (document.images) {
    document[imgName].src = eval(imgName + "_off.src");
  }
}