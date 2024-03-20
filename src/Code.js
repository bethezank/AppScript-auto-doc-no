// กำนหนดค่าคงที่ใช้ทุก Function

// Sheet Id
const sheetId = "12FJygBVlkse52W5HZF58vIC0pXtaFXNlAhOEW9A8XHQ"

// Sheet Name
const sheetName = "DocumentNo"

// เข้าถึง sheet
const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName)


// ========================================================================================

// สร้างเมนู
function getMenu() {

  //  กำหนดเมนู และ function
  const menuItems = [
    { name: '🆕 สร้างชื่อเอกสารใหม่', functionName: 'newDocNo' },
  ];

  //  สร้าง object ui
  const ui = SpreadsheetApp.getUi();

  //  สร้าง menu
  const menu = ui.createMenu('⭐ Smart Structure')

  // วน Loop สร้าง menu item
  menuItems.forEach(item => {
    menu.addItem(item.name, item.functionName);
  });

  // เพิ่ม menu ลง toolbar
  menu.addToUi();
}

// ========================================================================================

// function หลัก
const home = () => {
  getMenu()
}

