function sistemOtomatis(e) {
  // 1. PENGAMAN DASAR
  if (!e || !e.source) return;

  // --- ANTRIAN (LOCK SYSTEM) ---
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000); 
  } catch (e) {
    SpreadsheetApp.getUi().alert('Sistem sibuk, coba lagi.');
    return;
  }

  try {
    var range = e.range;
    var sheet = range.getSheet();
    var sheetName = sheet.getName();
    var rowAwal = range.getRow();
    var colAwal = range.getColumn();
    var numRows = range.getNumRows(); // Deteksi berapa baris yang diedit

    if (rowAwal === 1) return; // Abaikan header

    // --- CONFIG GLOBAL ---
    var listOutlet = ['kwb', 'kwu', 'kwm', 'mut', 'mus', 'mms', 'klinik mut'];
    var listTerjualDist = ['sudah', 'belum', 'promo', 'Retur', 'EXP']; 
    var listTerjualPromo = ['sudah', 'belum', 'Retur', 'EXP'];
    
    var listTujuanInput = ['Distribusi', 'Promo', 'Retur', 'EXP'];
    var listTujuanPromo = ['Retur', 'EXP'];
    var listKetRetur = ['Kongsi', 'Tidak Kongsi'];

    // --- KONFIGURASI SHEET ---
    var config = {};

    if (sheetName === "Input_Data") {
      config = {
        sheetType: "input",

        // B - G (Nama s/d Harga)
        colStart: 2,
        colEnd: 7,

        colTanggal: 1,   // A
        colJml: 5,        // E
        colHrg: 7,        // G
        colTotal: 8,      // H

        colAlihkan: 9,    // I
        colKurangi: 10,   // J  << KURANGI JUMLAH
        colDone: 11,      // K  << CHECKBOX KONFIRMASI

        opsiDropdown: listTujuanInput
      };
    }

    else if (sheetName === "Distribusi") {
      config = {
        sheetType: "distribusi",
        colStart: 2, colEnd: 8,
        colTanggal: 1,

        colJml: 5,         // E stok
        colHrg: 7,         // G harga
        colTotal: 8,       // H total

        colOutlet: 9,      // I
        colTerjual: 10,    // J
        colUpdateTgl: 11,  // K
        colJmlTerjual: 12, // L
        colDone: 13,       // M (checkbox tunggal)
        colKonfJual: 13    // samakan saja biar fungsi parsial pakai ini
      };
    }
    else if (sheetName === "Promo") {
      config = {
        sheetType: "promo",
        colStart: 2, colEnd: 12,
        colTanggal: 1,

        colProgram: 9,     // I
        colOutlet: 10,     // J
        colTerjual: 11,    // K
        colUpdateTgl: 12,  // L

        colJml: 5,         // E  ✅ WAJIB untuk prosesPenjualanParsial
        colHrg: 7,         // G  (kalau mau konsisten)
        colTotal: 8,       // H  (kalau mau konsisten)

        colJmlTerjual: 13, // M
        colDone: 14,       // N
        colKonfJual: 14    // (biar prosesPenjualanParsial pakai ini)
      };
    }
    else if (sheetName === "Retur") {
      config = {
        sheetType: "retur",
        colStart: 6,    colEnd: 9,    
        colKet: 14,     colPantau: 15, colDone: 15,
        opsiDropdown: listKetRetur
      };
    }
    else {
      return; 
    }
    // ======================================================
    // BAGIAN 1: EXECUTOR CHECKBOX
    // ======================================================
    if (colAwal === config.colDone || (config.colKonfJual && colAwal === config.colKonfJual)) {

      var values = range.getValues();
      var rowsToDelete = []; // ✅ antrian delete

      for (var i = numRows - 1; i >= 0; i--) {
        var isChecked = values[i][0];
        var currentRow = rowAwal + i;
        if (isChecked !== true) continue;

        var ok = false;

        // ===== INPUT_DATA (1 checkbox: Done atau Kurangi) =====
        if (config.sheetType === "input" && colAwal === config.colDone) {
          var kurangiRaw = sheet.getRange(currentRow, config.colKurangi).getValue();
          var kurangi = (typeof kurangiRaw === "number") ? kurangiRaw : Number(String(kurangiRaw).replace(",", "."));

          if (kurangiRaw !== "" && kurangiRaw != null && !isNaN(kurangi) && kurangi > 0) {
            ok = prosesKurangiDanAlihkanInput(sheet, currentRow, config, e.source, listOutlet, listTerjualDist);
            if (ok) updateAnalisisRealtime(e.source);
          } else {
            ok = prosesPindahData(sheet, currentRow, config, e.source, listOutlet, listTerjualDist, listTerjualPromo, listTujuanPromo);
          }

          if (ok) rowsToDelete.push(currentRow);
          continue;
        }
        // ===== DISTRIBUSI (1 checkbox: pindah atau parsial) =====
        if (config.sheetType === "distribusi" && colAwal === config.colDone) {

          var status = String(sheet.getRange(currentRow, config.colTerjual).getValue() || "")
            .trim().toLowerCase();

          var cellKonf = sheet.getRange(currentRow, config.colDone);

          var lRaw = sheet.getRange(currentRow, config.colJmlTerjual).getValue();
          var lVal = (typeof lRaw === "number") ? lRaw : Number(String(lRaw).replace(",", "."));

          // 1) PRIORITAS pindah: promo/retur/exp => L harus kosong
          if (status === "promo" || status === "retur" || status === "exp") {

            if (lRaw !== "" && lRaw != null && !isNaN(lVal) && lVal > 0) {
              SpreadsheetApp.getUi().alert("Kalau Terjual = Promo/Retur/EXP, kolom Jumlah Terjual (L) harus kosong.");
              cellKonf.uncheck();
              continue;
            }

            ok = prosesPindahData(sheet, currentRow, config, e.source, listOutlet, listTerjualDist, listTerjualPromo, listTujuanPromo);
            if (ok) rowsToDelete.push(currentRow);
            continue;
          }

          // 2) Selain itu => parsial hanya kalau L>0
          if (lRaw !== "" && lRaw != null && !isNaN(lVal) && lVal > 0) {
            ok = prosesPenjualanParsial(sheet, currentRow, config, e.source);
            if (ok) updateAnalisisRealtime(e.source);

            // parsial biasanya tidak delete baris (stok masih ada)
            continue;
          }

          // 3) kalau L kosong dan status bukan promo/retur/exp => tidak boleh eksekusi
          SpreadsheetApp.getUi().alert("Isi Jumlah Terjual (L) untuk parsial, atau pilih Terjual = Promo/Retur/EXP untuk pindah.");
          cellKonf.uncheck();
          continue;
        }

        // ===== PROMO (1 checkbox) =====
        if (config.sheetType === "promo" && colAwal === config.colDone) {

          var statusK = String(sheet.getRange(currentRow, config.colTerjual).getValue() || "")
            .trim()
            .toLowerCase(); // k: "Retur" => "retur"

          var cellKonf = sheet.getRange(currentRow, config.colDone); // N

          var mRaw = sheet.getRange(currentRow, config.colJmlTerjual).getValue(); // M
          var mVal = (typeof mRaw === "number") ? mRaw : Number(String(mRaw).replace(",", "."));

          ok = false;

          // 1) PRIORITAS: Retur/EXP => pindah (M harus kosong)
          if (statusK === "retur" || statusK === "exp") {

            if (mRaw !== "" && mRaw != null && !isNaN(mVal) && mVal > 0) {
              SpreadsheetApp.getUi().alert("Kalau Terjual = Retur/EXP, kolom Jumlah Terjual (M) harus kosong.");
              cellKonf.uncheck();
              continue;
            }

            ok = prosesPindahData(sheet, currentRow, config, e.source, listOutlet, listTerjualDist, listTerjualPromo, listTujuanPromo);
            if (ok) rowsToDelete.push(currentRow);
            continue;
          }

          // 2) MODE PARSIAL: selama M > 0, BOLEH walau K = belum/kosong/sudah
          if (mRaw !== "" && mRaw != null && !isNaN(mVal) && mVal > 0) {
            ok = prosesPenjualanParsial(sheet, currentRow, config, e.source);
            if (ok) {
              updateAnalisisRealtime(e.source);

              // sesuai request kamu: setelah parsial, kolom K dikosongkan biar bisa dipilih lagi
              sheet.getRange(currentRow, config.colTerjual).clearContent();
            }
            continue; // parsial tidak delete baris
          }

          // 3) Kalau bukan retur/exp dan M kosong => tidak boleh eksekusi
          SpreadsheetApp.getUi().alert("Isi Jumlah Terjual (M) untuk proses parsial, atau pilih Terjual = Retur/EXP untuk pindah.");
          cellKonf.uncheck();
          continue;
        }

        // ===== KONFIRMASI PARSIAL (kalau dipakai) =====
        if (config.colKonfJual && colAwal === config.colKonfJual) {
          ok = prosesPenjualanParsial(sheet, currentRow, config, e.source);
          if (ok) updateAnalisisRealtime(e.source);
          continue;
        }

        // ===== DEFAULT pindah data =====
        if (colAwal === config.colDone) {
          ok = prosesPindahData(sheet, currentRow, config, e.source, listOutlet, listTerjualDist, listTerjualPromo, listTujuanPromo);
          if (ok) rowsToDelete.push(currentRow);
          continue;
        }
      }

      // ✅ Hapus setelah semua proses selesai (baris besar -> kecil)
      rowsToDelete.sort(function(a, b){ return b - a; });
      rowsToDelete.forEach(function(r){
        sheet.deleteRow(r);
      });

      return;
    }

    // ======================================================
    // BAGIAN 2: LOGIKA EDIT UTAMA (FORMULA & DROPDOWN)
    // ======================================================
    // Bagian ini hanya berjalan jika kamu mengedit cell biasa (bukan centang eksekusi)
    // Kita gunakan loop juga supaya aman jika kamu copy-paste data banyak baris
    
    for (var i = 0; i < numRows; i++) {
        var row = rowAwal + i;
        var col = colAwal; // Kolom tetap sama karena edit biasanya vertikal

        // --- A. INPUT DATA (Hitung Total) ---
        if (config.sheetType === "input") {
          if (col === config.colJml || col === config.colHrg) {
            var valJml = sheet.getRange(row, config.colJml).getValue();
            var valHrg = sheet.getRange(row, config.colHrg).getValue();
            sheet.getRange(row, config.colTotal).setValue((typeof valJml === 'number' && typeof valHrg === 'number') ? valJml * valHrg : "");
          }
        }

        // --- B. DISTRIBUSI ---
        if (config.sheetType === "distribusi") {
          // Logika Saat Kolom Terjual Diedit (Visual Checkbox)
          if (col === config.colTerjual) {
            var valTerjual = sheet.getRange(row, config.colTerjual).getValue();
            var cellUpdate = sheet.getRange(row, config.colUpdateTgl);
            var cellDone = sheet.getRange(row, config.colDone);

            if (valTerjual === "promo" || valTerjual === "Retur" || valTerjual === "EXP") {
              cellUpdate.setValue(new Date()).setNumberFormat("dd/MM/yyyy HH:mm:ss");
              cellDone.clearDataValidations(); 
              cellDone.insertCheckboxes();
              cellDone.uncheck(); 
            } 
            else if (valTerjual === "sudah") {
              cellUpdate.setValue(new Date()).setNumberFormat("dd/MM/yyyy HH:mm:ss");
              cellDone.removeCheckboxes().clearContent();
              updateAnalisisRealtime(e.source);
            }
            else if (valTerjual === "belum") {
              cellUpdate.setValue(new Date()).setNumberFormat("dd/MM/yyyy HH:mm:ss");
              cellDone.removeCheckboxes(); 
              updateAnalisisRealtime(e.source);
            } else {
              cellDone.removeCheckboxes().clearContent();
            }
          }
          if (col === config.colOutlet || col === config.colTerjual || col === config.colJmlTerjual) {
            cekMunculkanKonfirmasiDistribusi_(sheet, row, config);
          }
          
          // Auto Dropdown
          var rangeProduk = sheet.getRange(row, 2, 1, 7).getValues()[0]; 
          var isProdukAda = rangeProduk.every(function(c) { return c !== "" && c !== null; });
          var cellOutlet = sheet.getRange(row, config.colOutlet);
          var cellTerjual = sheet.getRange(row, config.colTerjual);
          var cellDone = sheet.getRange(row, config.colDone); 
          
          if (isProdukAda) {
            if (cellOutlet.getDataValidation() == null) {
                cellOutlet.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(listOutlet).setAllowInvalid(false).build());
            }
            cellTerjual.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(listTerjualDist).setAllowInvalid(false).build());
          } else {
            cellOutlet.clearDataValidations().clearContent();
            cellTerjual.clearDataValidations().clearContent();
            cellDone.removeCheckboxes().clearContent();
          }
        }

        // --- C. PROMO ---
        if (config.sheetType === "promo") {
          var valProgram = sheet.getRange(row, config.colProgram).getValue();
          var cellOutlet  = sheet.getRange(row, config.colOutlet);
          var cellTerjual = sheet.getRange(row, config.colTerjual);
          var cellDone    = sheet.getRange(row, config.colDone);
          var cellUpdate  = sheet.getRange(row, config.colUpdateTgl);

          var produk = sheet.getRange(row, 2).getValue();
          var isProdukAda = (produk !== "" && produk !== null);

          if (isProdukAda) {

            // Outlet dropdown
            if (cellOutlet.getDataValidation() == null) {
              cellOutlet.setDataValidation(
                SpreadsheetApp.newDataValidation().requireValueInList(listOutlet).setAllowInvalid(false).build()
              );
            }

            var valOutlet = String(cellOutlet.getValue() || "").trim();

            // Terjual dropdown aktif hanya jika Program + Outlet terisi
            if (valProgram !== "" && valProgram != null && valOutlet !== "") {
              cellTerjual.setDataValidation(
                SpreadsheetApp.newDataValidation().requireValueInList(listTerjualPromo).setAllowInvalid(false).build()
              );
            } else {
              cellTerjual.clearContent().clearDataValidations();
              cellUpdate.clearContent();
              cellDone.removeCheckboxes().clearContent();
            }

            // Reaksi saat user memilih Terjual (K)
            if (col === config.colTerjual) {
              var valTerjualBaru = String(cellTerjual.getValue() || "").trim();

              // reset tombol dulu
              cellDone.removeCheckboxes().clearContent();

              if (valTerjualBaru === "Retur" || valTerjualBaru === "EXP") {
                // wajib kosongkan M agar tidak konflik
                sheet.getRange(row, config.colJmlTerjual).clearContent();

                cellUpdate.setValue(new Date()).setNumberFormat("dd/MM/yyyy HH:mm:ss");
                cellDone.insertCheckboxes();
                cellDone.uncheck();
              }
              else if (valTerjualBaru === "sudah") {
                // belum munculkan checkbox sampai M diisi (biar disiplin)
                cellUpdate.setValue(new Date()).setNumberFormat("dd/MM/yyyy HH:mm:ss");
                // checkbox akan dimunculkan saat M diisi (lihat tambahan C)
              }
              else if (valTerjualBaru === "belum") {
                cellUpdate.setValue(new Date()).setNumberFormat("dd/MM/yyyy HH:mm:ss");
                // tidak ada checkbox
              } else {
                cellUpdate.clearContent();
              }
            }
            // =====================================================
            // MUNCULKAN CHECKBOX KONFIRMASI (N) berdasar K & M
            // =====================================================
            if (col === config.colJmlTerjual || col === config.colTerjual || col === config.colProgram || col === config.colOutlet) {

              var statusK = String(sheet.getRange(row, config.colTerjual).getValue() || "").trim().toLowerCase();
              var mRaw = sheet.getRange(row, config.colJmlTerjual).getValue();                                   
              var mVal = (typeof mRaw === "number") ? mRaw : Number(String(mRaw).replace(",", "."));

              var outletVal = String(cellOutlet.getValue() || "").trim();
              var programVal = String(valProgram || "").trim();

              // reset dulu checkbox N
              cellDone.removeCheckboxes().clearContent();

              // syarat dasar: produk ada + program + outlet harus terisi
              var syaratDasar = (programVal !== "" && outletVal !== "");

              if (!syaratDasar) return;
              if (statusK === "retur" || statusK === "exp") {
                if (mRaw === "" || mRaw == null) {
                  cellDone.insertCheckboxes();
                  cellDone.uncheck();
                } else {
                }
                return;
              }

              // 2) K=sudah => checkbox muncul hanya jika M>0
              if (statusK === "sudah") {
                if (mRaw !== "" && mRaw != null && !isNaN(mVal) && mVal > 0) {
                  cellDone.insertCheckboxes();
                  cellDone.uncheck();
                }
                return;
              }
              if (statusK === "belum" || statusK === "") {
                if (mRaw !== "" && mRaw != null && !isNaN(mVal) && mVal > 0) {
                  cellDone.insertCheckboxes();
                  cellDone.uncheck();
                }
                return;
              }
            }
          } else {
            cellOutlet.clearDataValidations().clearContent();
            cellTerjual.clearDataValidations().clearContent();
            cellDone.removeCheckboxes().clearContent();
            cellUpdate.clearContent();
          }
        }

        // --- D. CEK INPUT DATA (Alihkan & Checkbox) ---
        if (config.sheetType === "input") {
             var jumlahKolomCek = config.colEnd - config.colStart + 1;
             var inputValues = sheet.getRange(row, config.colStart, 1, jumlahKolomCek).getValues()[0];
             var isLengkap = inputValues.every(function(cell) { return cell !== "" && cell !== null; });
             var cellAlihkan = sheet.getRange(row, config.colAlihkan);
             var cellDone = sheet.getRange(row, config.colDone);
             var cellTanggal = sheet.getRange(row, config.colTanggal);

             if (isLengkap) {
               if (cellAlihkan.getDataValidation() == null) cellAlihkan.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(config.opsiDropdown).setAllowInvalid(false).build());
               if (cellTanggal.getValue() === "") cellTanggal.setValue(new Date()).setNumberFormat("dd/MM/yyyy HH:mm:ss");
             } else {
               cellAlihkan.clearDataValidations().clearContent();
               cellDone.removeCheckboxes().clearContent();
             }
             if (col === config.colAlihkan) {
               var valAlihkan = sheet.getRange(row, config.colAlihkan).getValue();
               if (valAlihkan !== "" && valAlihkan != null) {
                  if (!cellDone.isChecked()) cellDone.insertCheckboxes();
               } else {
                  cellDone.removeCheckboxes();
               }
             }
             cekMunculkanKonfirmasiKurangi_(sheet, row, config);
        }
    } // End Loop Biasa

    // Update Dashboard di akhir
    if (["distribusi", "promo", "retur", "exp"].indexOf(config.sheetType) > -1) {
          updateAnalisisRealtime(e.source);
    }
  } catch (err) {
    console.error(err);
  } finally {
    lock.releaseLock();
  }
}

// ========================================================
// FUNGSI PROSES PINDAH DATA
// ========================================================
function prosesPindahData(sheet, row, config, spreadsheet, listOutlet, listTerjualDist, listTerjualPromo, listTujuanPromo) {

  // ===== 1. AMBIL STATUS / TUJUAN PINDAH =====
  var status = "";

  if (config.sheetType === "distribusi") {
    status = sheet.getRange(row, config.colTerjual).getValue();

  } else if (config.sheetType === "promo") {
    // ✅ PROMO: status dari kolom TERJUAL (K)
    status = sheet.getRange(row, config.colTerjual).getValue();

  } else if (config.sheetType === "input") {
    // ✅ INPUT: status dari kolom ALIHKAN (I)
    status = sheet.getRange(row, config.colAlihkan).getValue();
  }

  var statusLower = String(status || "").trim().toLowerCase();

  var targetSheetName = "";
  if (statusLower === "promo") targetSheetName = "Promo";
  else if (statusLower === "retur") targetSheetName = "Retur";
  else if (statusLower === "exp") targetSheetName = "EXP";
  else if (statusLower === "distribusi") targetSheetName = "Distribusi";

  if (!targetSheetName) {
    // gagal menentukan tujuan -> batal
    sheet.getRange(row, config.colDone).uncheck();
    spreadsheet.toast("Gagal pindah: status/tujuan kosong atau tidak valid", "INFO", 5);
    return;
  }

  var targetSheet = spreadsheet.getSheetByName(targetSheetName);

  if (!targetSheet) {
    sheet.getRange(row, config.colDone).uncheck();
    spreadsheet.toast("Sheet tujuan tidak ditemukan: " + targetSheetName, "ERROR", 5);
    return;
  }

  // ===== 2. AMBIL DATA SUMBER (A-H) =====
  var valTanggal = sheet.getRange(row, 1).getValue();
  var valNama    = sheet.getRange(row, 2).getValue();
  var valKand    = sheet.getRange(row, 3).getValue();
  var valED      = sheet.getRange(row, 4).getValue();
  var valJml     = sheet.getRange(row, 5).getValue();
  var valSat     = sheet.getRange(row, 6).getValue();
  var valHrg     = sheet.getRange(row, 7).getValue();
  var valTot     = sheet.getRange(row, 8).getValue();

  var asalOutlet = "";
  if (config.sheetType === "distribusi" || config.sheetType === "promo") {
    asalOutlet = sheet.getRange(row, config.colOutlet).getValue();
  }

  // ===== 3. SIAPKAN ROW BARU SESUAI TARGET =====
  var newRowData = [];

  if (targetSheetName === "Promo") {
    // [Tgl, Nama, Kand, ED, Jml, Sat, Hrg, Tot, Program, Outlet, Status, Update]
    newRowData = [valTanggal, valNama, valKand, valED, valJml, valSat, valHrg, valTot, "", asalOutlet, "", ""];

  } else if (targetSheetName === "Distribusi") {
    // [Tgl, Nama, Kand, ED, Jml, Sat, Hrg, Tot, Outlet, Terjual, Update]
    newRowData = [valTanggal, valNama, valKand, valED, valJml, valSat, valHrg, valTot, "", "", ""];

  } else if (targetSheetName === "EXP") {
    // EXP ambil lengkap A-H
    newRowData = [valTanggal, valNama, valKand, valED, valJml, valSat, valHrg, valTot];

  } else if (targetSheetName === "Retur") {
    // Retur tetap A-E (kalau memang maunya begitu)
    newRowData = [valTanggal, valNama, valED, valJml, valSat];
  }

  // ===== 4. APPEND + FORMAT MINIMAL =====
  try {
    targetSheet.appendRow(newRowData);
  } catch (err) {
    sheet.getRange(row, config.colDone).uncheck();
    spreadsheet.toast("Append gagal: " + err.message, "ERROR", 8);
    throw err;
  }

  var lastRow = targetSheet.getLastRow();
  targetSheet.getRange(lastRow, 1).setNumberFormat("dd/MM/yyyy");

  // dropdown minimal
  if (targetSheetName === "Distribusi") {
    targetSheet.getRange(lastRow, 9).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(listOutlet).setAllowInvalid(false).build()
    );
    targetSheet.getRange(lastRow, 10).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(listTerjualDist).setAllowInvalid(false).build()
    );
  } else if (targetSheetName === "Promo") {
    targetSheet.getRange(lastRow, 10).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(listOutlet).setAllowInvalid(false).build()
    );
  }
  // ===== 5. HAPUS BARIS SUMBER =====
  return true;
}


function prosesKurangiDanAlihkanInput(sheet, row, config, spreadsheet, listOutlet, listTerjualDist) {
  var ui = SpreadsheetApp.getUi();

  var cellJml = sheet.getRange(row, config.colJml);          // E
  var cellKurangi = sheet.getRange(row, config.colKurangi);  // K
  var cellKonf = sheet.getRange(row, config.colDone); // L
  var cellAlihkan = sheet.getRange(row, config.colAlihkan);  // I

  var jumlahAwal = cellJml.getValue();
  var kurangi = cellKurangi.getValue();
  var alihkan = cellAlihkan.getValue();

  var hargaSatuan = sheet.getRange(row, config.colHrg).getValue();   // G
  var cellTotal = sheet.getRange(row, config.colTotal);             // H

  // Validasi dasar
  if (typeof jumlahAwal !== 'number' || typeof kurangi !== 'number' || typeof hargaSatuan !== 'number') {
    ui.alert("Pastikan Jumlah (E), Kurangi (K), dan Harga Satuan (G) berupa angka.");
    cellKonf.uncheck();
    return;
  }
  if (!alihkan) {
    ui.alert("Pilih tujuan Alihkan (kolom I) dulu.");
    cellKonf.uncheck();
    return;
  }
  if (kurangi <= 0) {
    cellKonf.uncheck();
    return;
  }
  if (kurangi > jumlahAwal) {
    ui.alert("Kurangi Jumlah tidak boleh lebih besar dari Jumlah.");
    cellKonf.uncheck();
    return;
  }

  // Hitung bagian yang dipindah & sisa
  var sisa = jumlahAwal - kurangi;
  var totalPindah = kurangi * hargaSatuan;
  var totalSisa = sisa * hargaSatuan;

  // Tentukan sheet tujuan
  var tujuan = String(alihkan).toLowerCase();
  var targetSheetName = "";
  if (tujuan === "promo") targetSheetName = "Promo";
  else if (tujuan === "retur") targetSheetName = "Retur";
  else if (tujuan === "exp") targetSheetName = "EXP";
  else if (tujuan === "distribusi") targetSheetName = "Distribusi";

  if (!targetSheetName) {
    sheet.getRange(row, config.colDone).uncheck();
    spreadsheet.toast("Gagal pindah: status/tujuan kosong atau tidak valid", "INFO", 5);
    return false;
  }

  var targetSheet = spreadsheet.getSheetByName(targetSheetName);

  if (!targetSheet) {
    sheet.getRange(row, config.colDone).uncheck();
    spreadsheet.toast("Sheet tujuan tidak ditemukan: " + targetSheetName, "ERROR", 5);
    return false;
  }

  // Ambil data sumber (A-H)
  var valTanggal = sheet.getRange(row, 1).getValue();
  var valNama    = sheet.getRange(row, 2).getValue();
  var valKand    = sheet.getRange(row, 3).getValue();
  var valED      = sheet.getRange(row, 4).getValue();
  var valSat     = sheet.getRange(row, 6).getValue();

  // Data yang dipindah: pakai jumlah = kurangi, total = totalPindah
  var newRowData = [];

  if (targetSheetName === "Promo") {
    // [Tgl,Nama,Kand,ED,Jml,Sat,Hrg,Tot,ProgramKosong,OutletKosong,Status,Update]
    newRowData = [valTanggal, valNama, valKand, valED, kurangi, valSat, hargaSatuan, totalPindah, "", "", "", ""];
  } 
  else if (targetSheetName === "Distribusi") {
    // [Tgl,Nama,Kand,ED,Jml,Sat,Hrg,Tot,OutletKosong,Status,Update]
    newRowData = [valTanggal, valNama, valKand, valED, kurangi, valSat, hargaSatuan, totalPindah, "", "", ""];
  }
  else if (targetSheetName === "EXP") {
    newRowData = [valTanggal, valNama, valKand, valED, kurangi, valSat, hargaSatuan, totalPindah];
  }
  else if (targetSheetName === "Retur") {
    newRowData = [valTanggal, valNama, valED, kurangi, valSat];
  }

  // Append
  targetSheet.appendRow(newRowData);

  // Injeksi format & dropdown minimal
  var lastRow = targetSheet.getLastRow();
  targetSheet.getRange(lastRow, 1).setNumberFormat("dd/MM/yyyy");

  if (targetSheetName === "Distribusi") {
    targetSheet.getRange(lastRow, 9).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(listOutlet).setAllowInvalid(false).build()
    );
    targetSheet.getRange(lastRow, 10).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(listTerjualDist).setAllowInvalid(false).build()
    );
  } 
  else if (targetSheetName === "Promo") {
    targetSheet.getRange(lastRow, 10).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(listOutlet).setAllowInvalid(false).build()
    );
  }
  else if (targetSheetName === "Retur") {
    // Pastikan keterangan di kolom 14 (N) dan Done di 15 (O)
    var listKet = ['Barang Rusak', 'Hampir ED', 'Salah Pesan', 'Recall', 'Lainnya'];
    targetSheet.getRange(lastRow, 14).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(listKet).setAllowInvalid(false).build()
    );
    targetSheet.getRange(lastRow, 15).insertCheckboxes();
  }
  // Update baris asal (sisa)
  if (sisa > 0) {
    cellJml.setValue(sisa);
    cellTotal.setValue(totalSisa);

    // ✅ Bersih-bersih: hanya reset aksi, data utama tetap
    sheet.getRange(row, config.colAlihkan).clearContent();
    sheet.getRange(row, config.colKurangi).clearContent();
    sheet.getRange(row, config.colDone).removeCheckboxes().clearContent();
  } else {
    // kalau sisa 0, baris sumber dihapus (karena sudah habis)
    sheet.deleteRow(row);
  }

}

function prosesPenjualanParsial(sheet, row, config, spreadsheet) {
  var ui = SpreadsheetApp.getUi();
  
  // 1. Ambil Data
  var cellJmlStok = sheet.getRange(row, config.colJml);        
  var cellJmlJual = sheet.getRange(row, config.colJmlTerjual); 
  var cellKonf    = sheet.getRange(row, config.colKonfJual);   
  
  var colHargaSatuan = 7; 
  var colTotalHarga  = 8; 

  var stokAwal = cellJmlStok.getValue();
  var jmlJual  = cellJmlJual.getValue();
  var hargaSatuan = sheet.getRange(row, colHargaSatuan).getValue(); 

  // 2. Validasi
  if (typeof stokAwal !== 'number' || typeof jmlJual !== 'number') {
    ui.alert("Error: Pastikan input angka benar."); cellKonf.uncheck(); return;
  }
  if (jmlJual > stokAwal) {
    ui.alert("Stok kurang!"); cellKonf.uncheck(); return;
  }
  if (jmlJual <= 0) { cellKonf.uncheck(); return; }

  // 3. Hitung
  var sisaStok = stokAwal - jmlJual;
  var totalHargaSisa = sisaStok * hargaSatuan;    
  var totalHargaTerjual = jmlJual * hargaSatuan;  

  // 4. Siapkan Data Arsip
  var rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  rowData[4] = jmlJual;           
  rowData[7] = totalHargaTerjual; 
  rowData[config.colTerjual - 1] = "sudah"; 

  var tglSekarang = new Date();
  rowData[config.colUpdateTgl - 1] = tglSekarang; 

  var cutOffIndex = config.colUpdateTgl; 
  var dataToArchive = rowData.slice(0, cutOffIndex);

  // 5. KIRIM KE ARSIP (DENGAN LOGIKA "LANTAI DASAR" LABEL)
  var sheetArsip = spreadsheet.getSheetByName("Arsip");
  if (sheetArsip) {
    var targetCol = 1; // Default Distribusi (Kolom A)

    if (config.sheetType === "promo") {
      targetCol = 37; // Promo (Kolom AL)
    }

    // A. Cari baris terakhir data di kolom tersebut
    var lastRowSpecific = cekBarisTerakhirPerKolom(sheetArsip, targetCol);
    
    // B. Cari baris Label Bulan Terakhir (Sebagai Batas Lantai)
    var lastRowLabel = cariPosisiLabelTerakhir(sheetArsip);
    
    var targetRow = Math.max(lastRowSpecific, lastRowLabel) + 1;

    // Tulis data
    sheetArsip.getRange(targetRow, targetCol, 1, dataToArchive.length).setValues([dataToArchive]);
  }

  // 6. UPDATE SHEET ASAL
  if (sisaStok > 0) {
    cellJmlStok.setValue(sisaStok); 
    sheet.getRange(row, colTotalHarga).setValue(totalHargaSisa); 
    sheet.getRange(row, config.colUpdateTgl).setValue(tglSekarang).setNumberFormat("dd/MM/yyyy HH:mm:ss"); 
    sheet.getRange(row, config.colTerjual).clearContent();

  } else {
    cellJmlStok.setValue(0);
    sheet.getRange(row, colTotalHarga).setValue(0);
  }
  
  // 7. Bersihkan
  cellJmlJual.clearContent();
  cellKonf.removeCheckboxes().clearContent(); 
}

// ========================================================
// FUNGSI BANTUAN: CARI BARIS TERAKHIR DI KOLOM TERTENTU
// ========================================================
function cekBarisTerakhirPerKolom(sheet, colIndex) {
  var lastRow = sheet.getLastRow();
  if (lastRow == 0) return 0;
  
  var data = sheet.getRange(1, colIndex, lastRow, 1).getValues();
  for (var i = data.length - 1; i >= 0; i--) {
    if (data[i][0] !== "" && data[i][0] != null) {
      return i + 1;
    }
  }
  return 0;
}

// ========================================================
// 3. UPDATE ANALISIS REALTIME (RETUR & EXP SAJA)
// ========================================================
function updateAnalisisRealtime(spreadsheet) {
  var ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();

  var sheetHistory = ss.getSheetByName("History");
  var sheetRetur   = ss.getSheetByName("Retur");
  var sheetExp     = ss.getSheetByName("EXP");

  if (!sheetHistory || !sheetRetur || !sheetExp) return;

  // --- SETUP HEADER ---
  var headerStyle = sheetHistory.getRange("K1:M3");
  headerStyle.setHorizontalAlignment("center")
             .setVerticalAlignment("middle")
             .setFontWeight("bold");

  // =========================================================
  // HITUNG RETUR (DATA MULAI BARIS 3)
  // kolom checkbox Retur kamu: kolom 15 (O) => index 14
  // =========================================================
  var returSudah = 0, returBelum = 0;
  var lastRowRetur = sheetRetur.getLastRow();

  if (lastRowRetur >= 3) {
    var dataRetur = sheetRetur.getRange(3, 1, lastRowRetur - 2, 15).getValues();

    for (var i = 0; i < dataRetur.length; i++) {
      var namaProd  = dataRetur[i][1];   // kolom B
      var isChecked = dataRetur[i][14];  // kolom O

      if (namaProd !== "" && namaProd != null) {
        if (isChecked === true) returSudah++;
        else returBelum++;
      }
    }
  }

  // =========================================================
  // HITUNG EXP
  // hitung baris yang kolom B (Nama Produk) tidak kosong
  // =========================================================
  var countExp = 0;
  var lastRowExp = sheetExp.getLastRow();

  if (lastRowExp >= 2) {
    var dataExp = sheetExp.getRange(2, 2, lastRowExp - 1, 1).getValues(); // kolom B
    countExp = dataExp.filter(function(r){ return r[0] !== "" && r[0] != null; }).length;
  }

  // =========================================================
  // TULIS HASIL
  // =========================================================
  var resultRange = sheetHistory.getRange("K4:M4");
  resultRange.setFontSize(12).setHorizontalAlignment("center");

  sheetHistory.getRange("K4").setValue(returSudah);
  sheetHistory.getRange("L4").setValue(returBelum);
  sheetHistory.getRange("M4").setValue(countExp);
}