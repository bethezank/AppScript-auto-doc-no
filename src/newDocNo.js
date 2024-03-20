const newDocNo = () => {
  // กำหนด Row เริ่มต้น
  const startRow = 2;

  // กำหนด Column ของแต่ละข้อมูล
  const idCol = 1;
  const docNoCol = 2;
  const createdAtCol = 3;

  // รูปแบบของชื่อเอกสาร
  // รูปแบบที่ต้องการ "SSB-INVOICE-2024/0001"
  const docName = "SSB-INVOICE";

  // เลขที่เริ่มต้นเอกสาร
  let docNo = 0;
  let docYear = 0;

  // ========================================================================================

  // รับค่า Row สุดท้ายที่มีข้อมูล
  const lastRow = SpreadsheetApp.getActiveSpreadsheet().getLastRow();

  // ตรวจสอบว่า Row สุดท้ายเท่ากับ 1 หรือไม่
  // ถ้าเท่ากับ 1 ให้ใช้ค่าจาก startRow (เพื่อไม่เอา Header ของ Column)
  let currentRow = 0;
  if (lastRow == 1) {
    currentRow = startRow;
  } else {
    currentRow = lastRow;
  }

  // รับค่าจาก Column Id
  let lastId = sheet.getRange(currentRow, idCol).getValue();

  // ตรวจสอบค่าที่รับมาว่าเป็นค่าว่างหรือไม่
  // ถ้าเป็นค่าว่าง ให้กำหนดเป็นค่า 0
  if (lastId == "") {
    lastId = 0;
  }

  // รับค่าจาก Column DocNo
  // รูปแบบ "SSB-INVOICE-0000/0000"
  let lastDocNo = sheet.getRange(currentRow, docNoCol).getValue();

  // ตรวจสอบค่าที่รับมาว่าเป็นค่าว่างหรือไม่
  // ถ้าเป็นค่าว่าง ให้กำหนดเป็นค่าเริ่มต้นของ รูปแบบของชื่อเอกสาร
  if (lastDocNo != "") {
    // ถ้ารับค่าสำเร็จ ให้ทำการแยกส่วน String ออกจากกัน
    // 1.) แยกเลขที่เอกสารออกมาก่อน
    // Destucturing ค่าที่แยกได้ออกมา
    // เช่น
    // docNameYear คือ "SSB-INVOICE-0000"
    // docNo คือ "0000"
    let [getDocNameYear, getDocNo] = lastDocNo.split("/");

    // 2. แยกปีออกจากชื่อเอกสาร "docNameYear"
    // เช่น
    // ssb = "SSB"
    // docName = "INVOICE"
    // docYear = "2024"
    let [getSsb, getDocName, getDocYear] = getDocNameYear.split("-");

    // update ค่าตัวแปรให้ docNo และ docYear
    // เพื่อให้เป็นค่าล่าสุดที่อ่านมาจากตาราง
    // ค่าที่รับมาจากเอกสาร เป็น Type = String
    // ต้องแปลง Type => Integer ก่อน
    docNo = parseInt(getDocNo);
    docYear = parseInt(getDocYear);
  }

  // แยกรูปแบบของชื่อเอกสาร ด้วยการดักจับ "/"
  // Destucturing ค่าที่แยกได้ออกมา
  // เช่น
  // docNameYear คือ "SSB-INVOICE-2024"
  // docNo คือ "000"
  // let [docNameYear, docNo] = lastDocNo.split("/");
  // let [ssb, docName, docYear] = docNameYear.split("-");

  // ========================================================================================

  // รับค่าวันที่และเวลาที่สร้างเอกสาร
  const dateNow = new Date();

  // รับค่าปี
  const year = dateNow.getFullYear();

  // รับค่าเดือน
  const month = String(dateNow.getMonth()).padStart(2, "0");

  // รับค่าวัน
  const day = String(dateNow.getDate()).padStart(2, "0");

  // รับค่าชั่วโมง
  const hour = String(dateNow.getHours()).padStart(2, "0");

  // รับค่านาที
  const minutes = String(dateNow.getMinutes()).padStart(2, "0");

  // สร้างรูปแบบวันที่ด้วยวิธี template literal โดยการใช้ Backtick
  // รูปแบบที่ต้องการ "2024-2-19 17:00"
  const currentDate = `${year}-${month}-${day} ${hour}:${minutes}`;

  // ========================================================================================

  // ตรวจสอบปีที่รับมาจากเอกสาร เทียบกับปีปัจจุบัน
  // ถ้าเป็นปีเดียวกัน ให้บวกเลขที่เอกสาร = +1
  if (docYear == year) {
    docNo = parseInt(docNo) + 1;
  } else {
    // ถ้าปีปีปัจจุบันมากกว่า ให้เริ่มต้นใหม่ = 1
    docNo = 1;
  }

  // ชี้ตำแหน่งไปที่ Row ถัดไป โดยเอา Row สุดท้ายมา + 1
  currentRow = lastRow + 1;

  // สร้างตัวเลข Id ถัดไป โดยเอา Id สุดท้ายมา + 1
  const currentId = lastId + 1;

  // ========================================================================================

  // สร้างเลขที่เอกสารถัดไป โดยการรวม String ด้วยวิธี template literal โดยการใช้ Backtick
  // รูปแบบที่ต้องการ "SSB-INVOICE-2024/0001"
  // ใช้ Method padStart(4, '0') เพื่อสร้างตัวเลข 4 หลัก เช่น 0001, 0010, 0100 เป็นต้น
  const currentDocName = `${docName}-${year}/${String(docNo).padStart(4, "0")}`;

  // ========================================================================================

  // บันทึกค่าลง sheet
  sheet.getRange(currentRow, idCol).setValue(currentId);
  sheet.getRange(currentRow, docNoCol).setValue(currentDocName);
  sheet.getRange(currentRow, createdAtCol).setValue(currentDate);
};
