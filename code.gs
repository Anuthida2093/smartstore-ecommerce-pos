/**
 * ============================================================================
 * SMART STORE SYSTEM - FULL BACKEND (E-Commerce, POS Multi-items, CRM, Admin)
 * ============================================================================
 */

// ฟังก์ชันเข้าถึง Spreadsheet ปัจจุบัน
const getDb = () => SpreadsheetApp.getActiveSpreadsheet();

// ชื่อแท็บ (Sheet) ในฐานข้อมูล (ต้องสร้างใน Google Sheets ให้ตรงตามนี้)
const SHEETS = {
  PRODUCTS: "Products",
  ORDERS: "Orders",
  SALES: "Sales",
  USERS: "Users",
  CUSTOMERS: "Customers"
};

// ==========================================
// 0. แสดงผลหน้าเว็บ (Index.html / index.html)
// ==========================================
function doGet(e) {
  try {
    // ลองหาไฟล์ชื่อ Index (ตัว I ใหญ่)
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('SmartStore E-Commerce & POS')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (error) {
    // ถ้าหาไม่เจอ ให้ลองหาไฟล์ชื่อ index (ตัว i เล็ก)
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('SmartStore E-Commerce & POS')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}

// ==========================================
// 1. DATA HELPER (ฟังก์ชันช่วยดึงและแปลงข้อมูล)
// ==========================================
function getSheetData(sheetName) {
  try {
    const sheet = getDb().getSheetByName(sheetName);
    if (!sheet) return [];
    
    const values = sheet.getDataRange().getDisplayValues();
    if (values.length <= 1) return []; 
    
    const headers = values[0].map(h => h.trim());
    return values.slice(1).map(row => {
      let obj = {};
      headers.forEach((header, i) => {
        let key = header.charAt(0).toLowerCase() + header.slice(1);
        
        // จัดการกรณีชื่อคอลัมน์พิเศษให้ตรงกับที่ Frontend ต้องการ
        if (key.toLowerCase() === "imgurl" || key.toLowerCase() === "imageurl") key = "imgUrl";
        if (key.toLowerCase() === "floorprice") key = "floorPrice";
        if (key.toLowerCase() === "customername") key = "customerName";
        if (key.toLowerCase() === "customerphone") key = "customerPhone";
        if (key.toLowerCase() === "shippingfee") key = "shippingFee";
        if (key.toLowerCase() === "grandtotal") key = "grandTotal";
        if (key.toLowerCase() === "totalspent") key = "totalSpent";
        if (key.toLowerCase() === "slipurl") key = "slipUrl";
        if (key.toLowerCase() === "trackingno") key = "trackingNo";
        if (key.toLowerCase() === "paymentmethod") key = "paymentMethod";
        if (key.toLowerCase() === "idcard") key = "idCard";
        if (key.toLowerCase() === "startdate") key = "startDate";
        
        // แปลงค่าตัวเลขให้เป็น Number เพื่อให้ Frontend คำนวณได้ถูกต้อง
        let val = row[i];
        const numFields = ['cost', 'price', 'floorprice', 'stock', 'qty', 'cartqty', 'discount', 'cartdiscount', 'total', 'shippingfee', 'grandtotal', 'profit', 'points', 'totalspent'];
        if (numFields.includes(key.toLowerCase())) {
            val = Number(val) || 0;
        }
        obj[key] = val;
      });
      return obj;
    });
  } catch (e) {
    console.error("Error reading sheet: " + sheetName, e);
    return [];
  }
}

// ไดนามิกฟังก์ชันสำหรับบันทึกข้อมูลลงแถวใหม่
function appendToSheet(sheetName, dataObj) {
  const sheet = getDb().getSheetByName(sheetName);
  if(!sheet) throw new Error("ไม่พบแท็บชีต " + sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const newRow = headers.map(header => {
    const key = header.toLowerCase().trim();
    const foundKey = Object.keys(dataObj).find(k => k.toLowerCase() === key);
    return foundKey ? dataObj[foundKey] : "";
  });
  
  sheet.appendRow(newRow);
  return true;
}

// ไดนามิกฟังก์ชันสำหรับอัปเดตข้อมูลตาม ID (หรือ Username)
function updateInSheet(sheetName, idField, idValue, dataObj) {
  const sheet = getDb().getSheetByName(sheetName);
  if(!sheet) throw new Error("ไม่พบแท็บชีต " + sheetName);
  const data = sheet.getDataRange().getValues();
  if(data.length <= 1) return false;
  
  const headers = data[0].map(h => h.toString().toLowerCase().trim());
  const idIdx = headers.indexOf(idField.toLowerCase());
  if(idIdx === -1) throw new Error("ไม่พบคอลัมน์ระบุตัวตน (ID/Username)");

  for(let i=1; i<data.length; i++) {
      if(String(data[i][idIdx]).trim() === String(idValue).trim()) {
          headers.forEach((header, colIdx) => {
              const key = Object.keys(dataObj).find(k => k.toLowerCase() === header);
              if(key !== undefined) {
                  sheet.getRange(i+1, colIdx+1).setValue(dataObj[key]);
              }
          });
          return true;
      }
  }
  throw new Error("ไม่พบข้อมูลที่ต้องการอัปเดต");
}

// ==========================================
// 2. MAIN API ENDPOINTS (เรียกจาก Frontend)
// ==========================================

// โหลดข้อมูลเริ่มต้นทั้งหมด
function getInitialData() {
  try {
    return {
      products: getSheetData(SHEETS.PRODUCTS).reverse(),
      orders: getSheetData(SHEETS.ORDERS).reverse(), 
      sales: getSheetData(SHEETS.SALES).reverse(),
      users: getSheetData(SHEETS.USERS),
      customers: getSheetData(SHEETS.CUSTOMERS).reverse()
    };
  } catch (e) {
    return { error: true, message: e.toString() };
  }
}

// ระบบเข้าสู่ระบบ
function loginUser(username, password) {
  try {
    const users = getSheetData(SHEETS.USERS);
    const user = users.find(u => String(u.username).trim() === String(username).trim() && String(u.password).trim() === String(password).trim());
    if (user) {
      if(user.status === 'inactive') return { success: false, message: "บัญชีของคุณถูกระงับการใช้งาน" };
      return { success: true, user: user };
    }
    return { success: false, message: "Username หรือ Password ไม่ถูกต้อง" };
  } catch (e) { return { success: false, message: e.toString() }; }
}

// ==========================================
// 3. PRODUCTS (สินค้าคงคลัง)
// ==========================================
function addProduct(data) {
  try { appendToSheet(SHEETS.PRODUCTS, data); return { success: true }; } 
  catch (e) { return { success: false, message: e.toString() }; }
}

function updateProduct(data) {
  try { updateInSheet(SHEETS.PRODUCTS, 'id', data.id, data); return { success: true }; } 
  catch (e) { return { success: false, message: e.toString() }; }
}

// ==========================================
// 4. ORDERS (ลูกค้าสั่งซื้อผ่านหน้าเว็บออนไลน์)
// ==========================================
function placeOrder(data) {
  try {
    appendToSheet(SHEETS.ORDERS, data);
    // ตัดสต็อกเมื่อลูกค้าสั่งออนไลน์สำเร็จ
    if(data.items) {
      deductStockFromItems(JSON.parse(data.items));
    }
    return { success: true, id: data.id };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function updateOrderStatus(data) {
  try { updateInSheet(SHEETS.ORDERS, 'id', data.id, { status: data.status }); return { success: true }; } 
  catch (e) { return { success: false, message: e.toString() }; }
}

// ==========================================
// 5. SALES & POS (พนักงานขายหน้าร้าน)
// ==========================================
function addSale(data) {
  try {
    // 1. บันทึกบิลการขาย
    appendToSheet(SHEETS.SALES, data);
    
    // 2. ตัดสต็อกสินค้าจากตะกร้าแบบหลายชิ้น
    if(data.items) {
      deductStockFromItems(JSON.parse(data.items));
    }
    
    return { success: true };
  } catch (e) { return { success: false, message: e.toString() }; }
}

// ฟังก์ชันช่วยตัดสต็อกสินค้าอัตโนมัติ (รองรับทั้งการขาย POS แบบตะกร้า และ ออเดอร์เว็บ)
function deductStockFromItems(itemsArray) {
  const sheet = getDb().getSheetByName(SHEETS.PRODUCTS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => h.toString().toLowerCase().trim());
  const idIdx = headers.indexOf('id');
  const stockIdx = headers.indexOf('stock');

  if(idIdx > -1 && stockIdx > -1) {
    itemsArray.forEach(item => {
       // รองรับตัวแปร qty (จากฝั่งเว็บ) หรือ cartQty (จากฝั่ง POS)
       const qtyToDeduct = Number(item.qty) || Number(item.cartQty) || 1; 
       
       for(let i=1; i<data.length; i++) {
          if(String(data[i][idIdx]).trim() === String(item.id).trim()) {
             let currentStock = Number(data[i][stockIdx]) || 0;
             let newStock = Math.max(0, currentStock - qtyToDeduct);
             sheet.getRange(i+1, stockIdx+1).setValue(newStock);
             break; 
          }
       }
    });
  }
}

// ==========================================
// 6. USERS (พนักงาน และ แอดมิน)
// ==========================================
function registerUser(data) {
  try { appendToSheet(SHEETS.USERS, data); return { success: true }; } 
  catch (e) { return { success: false, message: e.toString() }; }
}

function updateUser(data) {
  try {
    updateInSheet(SHEETS.USERS, 'username', data.username, data);
    return { success: true };
  } catch (e) { return { success: false, message: e.toString() }; }
}

// ==========================================
// 7. CUSTOMERS (ลูกค้า และ ระบบสมาชิก CRM)
// ==========================================
function addCustomer(data) {
  try { appendToSheet(SHEETS.CUSTOMERS, data); return { success: true }; } 
  catch (e) { return { success: false, message: e.toString() }; }
}

function updateCustomer(data) {
  try { updateInSheet(SHEETS.CUSTOMERS, 'id', data.id, data); return { success: true }; } 
  catch (e) { return { success: false, message: e.toString() }; }
}

// ==========================================
// 8. GENERIC DELETE (ลบหลายรายการพร้อมกัน)
// ==========================================
function deleteItem(type, idsArray) {
  try {
    let sheetName = "";
    let idColumn = "id"; 
    
    if (type.toUpperCase() === 'PRODUCTS') sheetName = SHEETS.PRODUCTS;
    else if (type.toUpperCase() === 'USERS') { sheetName = SHEETS.USERS; idColumn = "username"; }
    else if (type.toUpperCase() === 'CUSTOMERS') sheetName = SHEETS.CUSTOMERS;
    else return { success: false, message: "Invalid type" };

    const sheet = getDb().getSheetByName(sheetName);
    if(!sheet) return { success: false, message: "ไม่พบชีตฐานข้อมูล" };

    const data = sheet.getDataRange().getValues();
    if(data.length <= 1) return { success: true };
    
    const headers = data[0].map(h => h.toString().toLowerCase().trim());
    const idIdx = headers.indexOf(idColumn.toLowerCase());

    if(idIdx === -1) return { success: false, message: "ไม่พบคอลัมน์ระบุตัวตน" };

    let deleted = 0;
    // วนลูปจากล่างขึ้นบน เพื่อป้องกัน Index เคลื่อนเมื่อลบแถว
    for(let i = data.length - 1; i > 0; i--) {
        if(idsArray.includes(String(data[i][idIdx]).trim())) {
            sheet.deleteRow(i + 1);
            deleted++;
        }
    }
    return { success: true, message: `ลบสำเร็จ ${deleted} รายการ` };
  } catch(e) { return { success: false, message: e.toString() }; }
}
