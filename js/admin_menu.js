if (document.images) {
 //main menu
  // Active images
  admin_menu_on = new Image();
  admin_menu_on.src = "images/menu/menu_admin_menu_on.gif";
  main_menu_on = new Image();
  main_menu_on.src = "images/menu/menu_main_menu_on.gif";
  
  // Inactive images
  admin_menu_off = new Image();
  admin_menu_off.src = "images/menu/menu_admin_menu_off.gif";
  main_menu_off = new Image();
  main_menu_off.src = "images/menu/menu_main_menu_off.gif";
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