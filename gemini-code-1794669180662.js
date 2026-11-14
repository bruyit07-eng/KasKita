const SHEET_TRANSAKSI = 'Transaksi';
const SHEET_DOMPET = 'Dompet';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('KasKita — Pencatatan Keuangan Pribadi')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Setup Sheet Transaksi
  let txSheet = ss.getSheetByName(SHEET_TRANSAKSI);
  if (!txSheet) {
    txSheet = ss.insertSheet(SHEET_TRANSAKSI);
    const txHeaders = ['ID', 'Tanggal', 'Tipe', 'Dompet', 'Kategori', 'Deskripsi', 'Jumlah'];
    txSheet.getRange(1, 1, 1, txHeaders.length).setValues([txHeaders]);
    txSheet.getRange(1, 1, 1, txHeaders.length)
      .setFontWeight('bold')
      .setBackground('#0284c7')
      .setFontColor('#ffffff');
    txSheet.setFrozenRows(1);
  }

  // 2. Setup Sheet Dompet + Data Bawaan
  let walletSheet = ss.getSheetByName(SHEET_DOMPET);
  if (!walletSheet) {
    walletSheet = ss.insertSheet(SHEET_DOMPET);
    const walletHeaders = ['ID', 'Nama', 'Jenis', 'SaldoAwal'];
    walletSheet.getRange(1, 1, 1, walletHeaders.length).setValues([walletHeaders]);
    walletSheet.getRange(1, 1, 1, walletHeaders.length)
      .setFontWeight('bold')
      .setBackground('#0284c7')
      .setFontColor('#ffffff');
    walletSheet.setFrozenRows(1);

    const defaultWallets = [
      ['WLT-1', 'Cash', 'Tunai', 0],
      ['WLT-2', 'SPay', 'E-Wallet', 0],
      ['WLT-3', 'OVO', 'E-Wallet', 0],
      ['WLT-4', 'BCA', 'Bank', 0]
    ];
    walletSheet.getRange(2, 1, defaultWallets.length, 4).setValues(defaultWallets);
  }

  return { txSheet, walletSheet };
}

function getInitialData() {
  const { txSheet, walletSheet } = initSheets();
  const tz = Session.getScriptTimeZone();

  // Ambil Daftar Dompet
  const wLastRow = walletSheet.getLastRow();
  let wallets = [];
  if (wLastRow > 1) {
    const wData = walletSheet.getRange(2, 1, wLastRow - 1, 4).getValues();
    wallets = wData
      .filter(row => row[0] !== '' && row[1] !== '')
      .map(row => ({
        id: String(row[0]),
        nama: String(row[1]),
        jenis: String(row[2] || 'E-Wallet'),
        saldoAwal: Number(row[3]) || 0
      }));
  }

  // Ambil Daftar Transaksi
  const tLastRow = txSheet.getLastRow();
  const tLastCol = txSheet.getLastColumn();
  let transactions = [];
  if (tLastRow > 1) {
    const tData = txSheet.getRange(2, 1, tLastRow - 1, Math.max(tLastCol, 7)).getValues();
    transactions = tData
      .filter(row => row[0] !== '')
      .map(row => {
        let formattedDate = row[1];
        if (row[1] instanceof Date) {
          formattedDate = Utilities.formatDate(row[1], tz, 'yyyy-MM-dd');
        }
        // Kompatibilitas jika sheet lama belum memiliki kolom Dompet
        if (tLastCol === 6) {
          return {
            id: String(row[0]),
            tanggal: String(formattedDate),
            tipe: String(row[2]),
            dompet: 'Cash',
            kategori: String(row[3]),
            deskripsi: String(row[4]),
            jumlah: Number(row[5]) || 0
          };
        }
        return {
          id: String(row[0]),
          tanggal: String(formattedDate),
          tipe: String(row[2]),
          dompet: String(row[3] || 'Cash'),
          kategori: String(row[4]),
          deskripsi: String(row[5]),
          jumlah: Number(row[6]) || 0
        };
      })
      .reverse();
  }

  return { wallets, transactions };
}

function addTransaction(payload) {
  const { txSheet } = initSheets();
  const id = 'KK-' + new Date().getTime();
  const newRow = [
    id,
    payload.tanggal,
    payload.tipe,
    payload.dompet,
    payload.kategori,
    payload.deskripsi,
    Number(payload.jumlah)
  ];
  txSheet.appendRow(newRow);

  return {
    success: true,
    transaction: {
      id: id,
      tanggal: payload.tanggal,
      tipe: payload.tipe,
      dompet: payload.dompet,
      kategori: payload.kategori,
      deskripsi: payload.deskripsi,
      jumlah: Number(payload.jumlah)
    }
  };
}

function deleteTransaction(id) {
  const { txSheet } = initSheets();
  const lastRow = txSheet.getLastRow();
  if (lastRow <= 1) return { success: false };

  const ids = txSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      txSheet.deleteRow(i + 2);
      return { success: true };
    }
  }
  return { success: false };
}

function addWallet(payload) {
  const { walletSheet } = initSheets();
  const nama = String(payload.nama || '').trim();
  if (!nama) return { success: false, message: 'Nama dompet wajib diisi.' };

  const lastRow = walletSheet.getLastRow();
  if (lastRow > 1) {
    const existingNames = walletSheet.getRange(2, 2, lastRow - 1, 1).getValues();
    const isDuplicate = existingNames.some(r => String(r[0]).toLowerCase() === nama.toLowerCase());
    if (isDuplicate) {
      return { success: false, message: 'Nama dompet sudah terdaftar.' };
    }
  }

  const id = 'WLT-' + new Date().getTime();
  const jenis = String(payload.jenis || 'E-Wallet');
  const saldoAwal = Number(payload.saldoAwal) || 0;

  walletSheet.appendRow([id, nama, jenis, saldoAwal]);

  return {
    success: true,
    wallet: {
      id: id,
      nama: nama,
      jenis: jenis,
      saldoAwal: saldoAwal
    }
  };
}

function deleteWallet(id) {
  const { walletSheet } = initSheets();
  const lastRow = walletSheet.getLastRow();
  if (lastRow <= 1) return { success: false };

  const ids = walletSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      walletSheet.deleteRow(i + 2);
      return { success: true };
    }
  }
  return { success: false };
}