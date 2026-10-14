/**
 * DETEKTIF BAHASA TAHUN 5 — SISTEM REKOD MARKAH
 * 1. Buka Google Sheets baharu > Extensions > Apps Script.
 * 2. Ganti SPREADSHEET_ID di bawah dengan ID Sheets anda.
 * 3. Jalankan fungsi setupSheet sekali dan benarkan akses.
 * 4. Deploy > New deployment > Web app > Execute as Me > Access Anyone.
 * 5. Tampal URL berakhir /exec dalam index.html.
 *
 * Jangan kongsi Google Sheet secara umum. Apps Script berjalan sebagai pemilik.
 */
const SPREADSHEET_ID = 'MASUKKAN_ID_GOOGLE_SHEETS_DI_SINI';
const SHEET_NAME = 'MARKAH_BM5';
const ANSWER_KEY = ['B','S','B','S','B','S','B','B','S','S','B','S','B','S','B'];
const HEADERS = ['Tarikh & Masa','ID Cubaan','Nama Murid','Kelas','Markah','Jumlah','Peratus (%)','Status','Pilihan Murid','Semakan'];

function openSheet_(){
  if(!SPREADSHEET_ID || SPREADSHEET_ID.indexOf('MASUKKAN_') === 0){
    throw new Error('Sila isi SPREADSHEET_ID dahulu.');
  }
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet=ss.getSheetByName(SHEET_NAME);
  if(!sheet){sheet=ss.insertSheet(SHEET_NAME);}
  if(sheet.getLastRow()===0){
    sheet.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1,1,1,HEADERS.length).setBackground('#12324a').setFontColor('white').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm:ss');
    sheet.getRange('I:J').setWrap(true);
  }
  return sheet;
}

function setupSheet(){
  const sh=openSheet_();
  sh.autoResizeColumns(1,HEADERS.length);
  Logger.log('Sheet markah sedia: '+sh.getParent().getUrl());
}

function safeText_(value,max){
  let text=String(value||'').replace(/[\x00-\x1f\x7f]/g,' ').trim().replace(/\s+/g,' ').slice(0,max);
  // Halang kandungan nama/kelas menjadi formula Google Sheets.
  if(/^[=+\-@]/.test(text)) text="'"+text;
  return text;
}

function json_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e){
  try{
    if(!e||!e.parameter||!e.parameter.payload)throw new Error('Tiada data diterima.');
    const data=JSON.parse(e.parameter.payload);
    const id=String(data.id||'').toLowerCase().trim();
    if(!/^[a-z0-9-]{12,80}$/.test(id)) throw new Error('ID cubaan tidak sah.');
    const nama=safeText_(data.nama,60);
    const kelas=safeText_(data.kelas,40);
    if(nama.length<2 || !kelas)throw new Error('Nama atau kelas tidak lengkap.');
    const answers=data.jawapan;
    if(!Array.isArray(answers)||answers.length!==ANSWER_KEY.length||!answers.every(x=>x==='B'||x==='S')){
      throw new Error('Jawapan tidak lengkap atau tidak sah.');
    }
    // Markah disemak semula di server — bukannya percaya markah dihantar pelayar.
    const markah=answers.reduce((sum,ans,i)=>sum+(ans===ANSWER_KEY[i]?1:0),0);
    const peratus=Math.round(markah/ANSWER_KEY.length*100);
    const semakan=answers.map((a,i)=>`${i+1}:${a===ANSWER_KEY[i]?'✓':'✗'}`).join(' | ');
    const lock=LockService.getScriptLock();
    lock.waitLock(20000);
    try{
      const sheet=openSheet_();
      // Permintaan berulang yang memiliki ID sama tidak mencipta rekod pendua.
      const existing=findAttempt_(sheet,id);
      if(existing){return json_({ok:true,duplicate:true});}
      sheet.appendRow([new Date(),id,nama,kelas,markah,ANSWER_KEY.length,peratus,'Selesai',answers.join(', '),semakan]);
      SpreadsheetApp.flush();
      return json_({ok:true});
    }finally{lock.releaseLock();}
  }catch(err){
    console.error('Ralat simpan markah:',err);
    return json_({ok:false,error:String(err && err.message||err)});
  }
}

function findAttempt_(sheet,id){
  const last=sheet.getLastRow();
  if(last<2)return null;
  return sheet.getRange(2,2,last-1,1).createTextFinder(id).matchEntireCell(true).findNext();
}

/** Semakan RESIT sahaja — tidak memaparkan nama murid atau data Sheet lain. */
function doGet(e){
  const p=(e&&e.parameter)||{};
  const callback=String(p.callback||'');
  if(!/^[$A-Za-z_][$A-Za-z_0-9]{0,80}$/.test(callback)){
    return json_({found:false,error:'Callback tidak sah.'});
  }
  let result={found:false};
  try{
    if(p.mode==='check'){
      const id=String(p.id||'').toLowerCase().trim();
      if(/^[a-z0-9-]{12,80}$/.test(id)){
        const sh=openSheet_();
        const match=findAttempt_(sh,id);
        if(match){
          const row=match.getRow();
          const [markah,jumlah]=sh.getRange(row,5,1,2).getValues()[0];
          result={found:true,markah:Number(markah),jumlah:Number(jumlah)};
        }
      }
    }
  }catch(err){console.error('Ralat semakan resit:',err);}
  // JSONP hanya untuk semakan resit tidak sensitif, bukan senarai nama/markah kelas.
  return ContentService.createTextOutput(callback+'('+JSON.stringify(result)+');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
