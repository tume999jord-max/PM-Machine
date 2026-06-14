// ==========================================
// PM Machine — Google Apps Script Backend
// ==========================================

var SHEET_ID      = "1c4mPNZ4BbMZ4GbfBA-oErjctkfsR9swUXoExqaWTPgs";
var SHEET_PM      = "ข้อมูลทั้งหมด";
var SHEET_MACHINE = "เครื่องจักรทั้งหมด";
var SHEET_TECH    = "ช่างเทคนิค";
var FOLDER_BEFORE = "14SqYwe8gidxEIaDMWWK0y6ylJWxUz09D";
var FOLDER_AFTER  = "1i1u8bIyAWlHZJWXO2PZ1pnIuRSEcyiyU";

// ─── CORS helper ───────────────────────────────────────────────
// ContentService ไม่รองรับ addHeader จึงใช้ setHeader แบบนี้แทน
function cors(jsonString) {
      var output = ContentService.createTextOutput(jsonString);
        output.setMimeType(ContentService.MimeType.JSON);
          return output;
}

function doOptions(e) {
      var output = ContentService.createTextOutput("");
      output.setMimeType(ContentService.MimeType.TEXT);
      return output;
}

// ─── GET: ดึงข้อมูลเครื่องจักร + ช่าง ──────────────────────────
function doGet(e) {
      try {
            var ss = SpreadsheetApp.openById(SHEET_ID);
                var action = e && e.parameter && e.parameter.action;

                    if (action === "getMachines") {
                              var sheet = ss.getSheetByName(SHEET_MACHINE);
                                    var data  = sheet.getDataRange().getValues();
                                          var machines = [];
                                                for (var i = 1; i < data.length; i++) {
                                                            if (data[i][0]) {
                                                                          machines.push({
                                                                                        name:    String(data[i][0]).trim(),
                                                                                                    regNo:   String(data[i][1]).trim(),
                                                                                                                dept:    String(data[i][2] || "").trim()
                                                                          });
                                                            }
                                                }
                                                      return cors(JSON.stringify({ ok: true, data: machines }));
                    }

                        if (action === "getMachine") {
                                  var regNo = e.parameter.regNo || "";
                                        var sheet = ss.getSheetByName(SHEET_MACHINE);
                                              var data  = sheet.getDataRange().getValues();
                                                    for (var i = 1; i < data.length; i++) {
                                                                if (String(data[i][1]).trim() === regNo) {
                                                                              return cors(JSON.stringify({
                                                                                            ok: true,
                                                                                                        data: {
                                                                                                                          name:  String(data[i][0]).trim(),
                                                                                                                                        regNo: String(data[i][1]).trim(),
                                                                                                                                                      dept:  String(data[i][2] || "").trim()
                                                                                                        }
                                                                              }));
                                                                }
                                                    }
                                                          return cors(JSON.stringify({ ok: false, msg: "ไม่พบเครื่องที่ตรงกัน" }));
                        }

                            if (action === "getTechs") {
                                      // สร้างชีตช่างถ้ายังไม่มี
                                            var techSheet = ss.getSheetByName(SHEET_TECH);
                                                  if (!techSheet) {
                                                            techSheet = ss.insertSheet(SHEET_TECH);
                                                                    techSheet.appendRow(["ชื่อ-สกุล"]);
                                                                            techSheet.appendRow(["ช่างเทคนิค 1"]);
                                                                                    techSheet.appendRow(["ช่างเทคนิค 2"]);
                                                  }
                                                        var techData = techSheet.getDataRange().getValues();
                                                              var techs = [];
                                                                    for (var j = 1; j < techData.length; j++) {
                                                                                if (techData[j][0]) techs.push(String(techData[j][0]).trim());
                                                                    }
                                                                          return cors(JSON.stringify({ ok: true, data: techs }));
                            }

                                return cors(JSON.stringify({ ok: false, msg: "unknown action" }));
      } catch (err) {
            return cors(JSON.stringify({ ok: false, msg: err.message }));
      }
}

// ─── POST: บันทึก PM record ─────────────────────────────────────
function doPost(e) {
      try {
            Logger.log("📥 doPost called");
            Logger.log("Request content type: " + (e.contentLength || 'unknown'));
            
            if (!e || !e.postData || !e.postData.contents) {
                      Logger.log("❌ No POST data received");
                      throw new Error("ไม่พบข้อมูล POST payload");
            }

                var payload = JSON.parse(e.postData.contents);
                    Logger.log("✅ Payload parsed successfully");
                    Logger.log("Data: regNo=" + payload.regNo + ", techName=" + payload.techName);
                    
                    var ss      = SpreadsheetApp.openById(SHEET_ID);
                        var sheet   = ss.getSheetByName(SHEET_PM);
                        
                        if (!sheet) {
                                  Logger.log("❌ Sheet not found: " + SHEET_PM);
                                  throw new Error("ไม่พบชีต " + SHEET_PM);
                        }

                            // อัปโหลดรูปก่อน
                                var urlBefore = "";
                                    if (payload.photoBefore && String(payload.photoBefore).trim()) {
                                              try {
                                                        Logger.log("📸 Uploading photo BEFORE...");
                                                        urlBefore = uploadPhoto(payload.photoBefore, FOLDER_BEFORE,
                                                                  "before_" + payload.regNo + "_" + Date.now());
                                                        Logger.log("✅ Photo BEFORE uploaded: " + urlBefore);
                                              } catch (photoErr) {
                                                        Logger.log("❌ uploadPhoto BEFORE error: " + photoErr.message);
                                                                throw new Error("ไม่สามารถอัปโหลดรูปก่อนได้: " + photoErr.message);
                                              }
                                    }

                                        // อัปโหลดรูปหลัง
                                            var urlAfter = "";
                                                if (payload.photoAfter && String(payload.photoAfter).trim()) {
                                                          try {
                                                                    Logger.log("📸 Uploading photo AFTER...");
                                                                    urlAfter = uploadPhoto(payload.photoAfter, FOLDER_AFTER,
                                                                              "after_" + payload.regNo + "_" + Date.now());
                                                                    Logger.log("✅ Photo AFTER uploaded: " + urlAfter);
                                                          } catch (photoErr) {
                                                                    Logger.log("❌ uploadPhoto AFTER error: " + photoErr.message);
                                                                            throw new Error("ไม่สามารถอัปโหลดรูปหลังได้: " + photoErr.message);
                                                          }
                                                }

                                                    // บันทึกลงชีต
                                                        // คอลัมน์: วันที่ | เครื่องจักร | ทะเบียนเครื่อง | ช่างเทคนิค | ผลการตรวจ | รูปก่อนทำ | รูปหลังทำ
                                                            Logger.log("💾 Appending row to sheet...");
                                                            sheet.appendRow([
                                                                      new Date(payload.date || Date.now()),
                                                                            payload.machineName  || "",
                                                                                  payload.regNo        || "",
                                                                                        payload.techName     || "",
                                                                                              payload.result       || "",
                                                                                                    urlBefore,
                                                                                                          urlAfter
                                                            ]);
                                                            Logger.log("✅ Row appended successfully");

                                                                return cors(JSON.stringify({ 
                                                                          ok: true, 
                                                                                urlBefore: urlBefore, 
                                                                                      urlAfter: urlAfter,
                                                                                            msg: "บันทึกสำเร็จ"
                                                                }));
      } catch (err) {
            Logger.log("❌ doPost error: " + err.message);
            Logger.log("Stack: " + err.stack);
                return cors(JSON.stringify({ ok: false, msg: err.message }));
      }
}

// ─── helper: base64 → Drive ─────────────────────────────────────
function uploadPhoto(base64String, folderId, filename) {
      if (base64String === undefined || base64String === null) {
            throw new Error("❌ base64String ว่าง");
      }

        if (typeof base64String !== "string") {
                try {
                          base64String = String(base64String);
                } catch (err) {
                          throw new Error("❌ base64String ไม่สามารถแปลงเป็น string ได้");
                }
        }

          base64String = base64String.trim();
            if (!base64String) {
                    throw new Error("❌ base64String เป็น string เปล่า");
            }

              var match = base64String.match(/^data:(image\/\w+);base64,(.+)$/);
                var mime, payload;
                  if (match) {
                        mime = match[1];
                            payload = match[2];
                  } else {
                        // หากได้ base64 ธรรมดา ไม่มี data: prefix
                            var raw = base64String.replace(/\s+/g, "");
                                if (/^[A-Za-z0-9+/]+={0,2}$/.test(raw)) {
                                          mime = "image/jpeg";
                                                payload = raw;
                                } else {
                                          Logger.log("❌ uploadPhoto invalid format: " + base64String.substring(0, 100));
                                                throw new Error("❌ รูปไม่ใช่ base64 ที่ถูกต้อง (ตรวจสอบ frontend)");
                                }
                  }

                    try {
                            var extension = mime.split("/")[1] || "jpg";
                                var data = Utilities.base64Decode(payload);
                                    var blob = Utilities.newBlob(data, mime, filename + "." + extension);
                                        var folder = DriveApp.getFolderById(folderId);
                                            var file = folder.createFile(blob);
                                                file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
                                                    var url = file.getUrl();
                                                        Logger.log("✅ uploadPhoto สำเร็จ: " + filename + " → " + url);
                                                            return url;
                    } catch (err) {
                            Logger.log("❌ uploadPhoto exception: " + err.message);
                                throw new Error("❌ uploadPhoto ล้มเหลว: " + err.message);
                    }
}

function testDoPost() {
      var event = {
            postData: {
                      contents: JSON.stringify({
                                date: "2026-06-13",
                                        machineName: "TEST เครื่อง",
                                                regNo: "TEST-001",
                                                        techName: "ทดสอบ ช่าง",
                                                                result: "ทดสอบบันทึก",
                                                                        photoBefore: "",
                                                                                photoAfter: ""
                      })
            }
      };

        var result = doPost(event);
          try {
                var content = result && typeof result.getContent === 'function' ? result.getContent() : JSON.stringify(result);
                    Logger.log("testDoPost result: " + content);
                        return content;
          } catch (err) {
                Logger.log("testDoPost error: " + err.message);
                    return JSON.stringify({ ok: false, msg: err.message });
          }
}