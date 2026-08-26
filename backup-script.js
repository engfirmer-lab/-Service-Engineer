/**
 * ===================================================================
 * TRBL Daily Backup Script (รันโดย GitHub Actions อัตโนมัติทุกวัน)
 * ===================================================================
 * ดึงข้อมูลทั้งหมดจาก Firestore -> สร้างไฟล์ JSON + Excel -> อัปโหลดเข้า
 * โฟลเดอร์ "Service Center/Backups" ใน Google Drive
 *
 * ไม่ต้องรันไฟล์นี้เองครับ — GitHub Actions จะรันให้อัตโนมัติทุกวัน
 * ตามที่ตั้งเวลาไว้ใน .github/workflows/daily-backup.yml
 * ===================================================================
 */
const admin = require('firebase-admin');
const { google } = require('googleapis');
const XLSX = require('xlsx');
const fs = require('fs');

const serviceAccountKey = JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT_KEY);
const DRIVE_FOLDER_ID = process.env.DRIVE_FOLDER_ID;

admin.initializeApp({ credential: admin.credential.cert(serviceAccountKey) });
const db = admin.firestore();

function docTypeLabel(key){
  const map = { service_manual:'Service Manual', spare_part:'Spare Part', install_checklist:'Install Checklist', pm_checklist:'PM Checklist', video:'Video' };
  return map[key] || key;
}

async function main(){
  console.log('เริ่มดึงข้อมูลจาก Firestore...');

  const [errorsSnap, solutionsSnap, docsSnap, propSnap, pendSnap, catDoc, supDoc] = await Promise.all([
    db.collection('errors').get(),
    db.collectionGroup('solutions').get(),
    db.collection('documents').get(),
    db.collection('proposals').get(),
    db.collection('pendingErrors').get(),
    db.doc('meta/categories').get(),
    db.doc('meta/supProducts').get(),
  ]);

  const errors = [];
  const errorMap = {};
  errorsSnap.forEach(d => {
    const e = { id: d.id, ...d.data(), solutions: [] };
    errors.push(e);
    errorMap[d.id] = e;
  });
  solutionsSnap.forEach(d => {
    const errorId = d.ref.parent.parent.id;
    const parentError = errorMap[errorId];
    if (parentError) parentError.solutions.push({ id: d.id, ...d.data() });
  });
  errors.forEach(e => e.solutions.sort((a, b) => (a.order ?? 0) - (b.order ?? 0)));

  const documents = []; docsSnap.forEach(d => documents.push({ id: d.id, ...d.data() }));
  const proposals = []; propSnap.forEach(d => proposals.push({ id: d.id, ...d.data() }));
  const pendingErrors = []; pendSnap.forEach(d => pendingErrors.push({ id: d.id, ...d.data() }));
  const categoriesData = catDoc.exists ? catDoc.data() : { categories: {}, owners: {} };
  const supProductsData = supDoc.exists ? supDoc.data() : { names: [] };

  const backup = {
    exportedAt: new Date().toISOString(),
    errors, documents, proposals, pendingErrors,
    categories: categoriesData.categories || {},
    owners: categoriesData.owners || {},
    supProducts: supProductsData.names || [],
  };

  console.log(`ดึงข้อมูลเสร็จ: ${errors.length} Error, ${documents.length} เอกสาร, ${proposals.length} ข้อเสนอ, ${pendingErrors.length} รายการรออนุมัติ`);

  const dateStamp = new Date().toISOString().slice(0, 10);

  // ----- ไฟล์ JSON -----
  const jsonFilename = `trbl-backup-${dateStamp}.json`;
  fs.writeFileSync(jsonFilename, JSON.stringify(backup, null, 2));

  // ----- ไฟล์ Excel (โครงสร้างเดียวกับปุ่ม Backup ในเว็บ) -----
  const wb = XLSX.utils.book_new();

  const errorRows = errors.map(e => ({ id: e.id, ประเภท: e.category, รุ่น: e.model, ErrorCode: e.errorCode || '', อาการ: e.symptom, เปิดดู: e.viewCount || 0, ยอดโหวตแก้ได้จริง: e.totalUps || 0 }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(errorRows), 'Errors');

  const solRows = [];
  errors.forEach(e => e.solutions.forEach((s, idx) => {
    solRows.push({ ประเภท: e.category, รุ่น: e.model, ErrorCode: e.errorCode || '', อาการ: e.symptom, ลำดับวิธี: idx + 1, คำอธิบาย: s.desc, อุปกรณ์: s.tools || '', อะไหล่: s.parts || '', เวลาโดยประมาณ: s.time || '', จำนวนรูป: (s.images || []).length, จำนวนโหวตแก้ได้จริง: (s.feedback || []).filter(f => f.type === 'up').length, ErrorId: e.id });
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(solRows), 'Solutions');

  const docRows = documents.map(d => ({ id: d.id, ประเภทเอกสาร: docTypeLabel(d.docType) + (d.videoSubtype ? ' - ' + d.videoSubtype : ''), ประเภทเครื่องจักร: d.category, รุ่น: d.model, ชื่อเอกสาร: d.title, ลิงก์: d.url || '', อัปโหลดเมื่อ: d.uploadedAt ? new Date(d.uploadedAt).toISOString() : '' }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(docRows), 'Documents');

  const propRows = proposals.map(p => ({ id: p.id, ErrorId: p.errorId, คำอธิบาย: p.desc, สถานะ: p.status, เสนอโดย: p.submitterName || '', วันที่: p.createdAt ? new Date(p.createdAt).toISOString() : '' }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(propRows), 'Proposals');

  const pendRows = pendingErrors.map(p => ({ id: p.id, ประเภท: p.category, รุ่น: p.model, ErrorCode: p.errorCode || '', อาการ: p.symptom, ส่งโดย: p.submitterName || '', สถานะ: p.status, วันที่: p.createdAt ? new Date(p.createdAt).toISOString() : '' }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(pendRows), 'PendingErrors');

  const catRows = [];
  Object.entries(backup.categories).forEach(([cat, models]) => {
    (models || []).forEach(m => {
      const key = cat + '|' + m;
      const owner = (backup.owners.modelOverrides || {})[key] || (backup.owners.categoryOwners || {})[cat] || '';
      catRows.push({ ประเภท: cat, รุ่น: m, ผู้ดูแล: owner });
    });
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(catRows), 'Categories');

  const excelFilename = `trbl-backup-${dateStamp}.xlsx`;
  XLSX.writeFile(wb, excelFilename);

  console.log('สร้างไฟล์ backup เสร็จแล้ว กำลังอัปโหลดขึ้น Google Drive...');

  // ----- อัปโหลดเข้า Google Drive -----
  const auth = new google.auth.GoogleAuth({
    credentials: serviceAccountKey,
    scopes: ['https://www.googleapis.com/auth/drive'],
  });
  const drive = google.drive({ version: 'v3', auth });

  // หาหรือสร้างโฟลเดอร์ "Backups" ใต้ Service Center
  let backupsFolderId;
  const searchRes = await drive.files.list({
    q: `'${DRIVE_FOLDER_ID}' in parents and name='Backups' and mimeType='application/vnd.google-apps.folder' and trashed=false`,
    fields: 'files(id, name)',
  });
  if (searchRes.data.files.length > 0) {
    backupsFolderId = searchRes.data.files[0].id;
  } else {
    const folderRes = await drive.files.create({
      resource: { name: 'Backups', mimeType: 'application/vnd.google-apps.folder', parents: [DRIVE_FOLDER_ID] },
      fields: 'id',
    });
    backupsFolderId = folderRes.data.id;
    console.log('สร้างโฟลเดอร์ "Backups" ใหม่ใน Service Center แล้ว');
  }

  await drive.files.create({
    resource: { name: jsonFilename, parents: [backupsFolderId] },
    media: { mimeType: 'application/json', body: fs.createReadStream(jsonFilename) },
  });
  await drive.files.create({
    resource: { name: excelFilename, parents: [backupsFolderId] },
    media: { mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', body: fs.createReadStream(excelFilename) },
  });

  console.log(`✓ อัปโหลด Backup วันที่ ${dateStamp} เข้า Service Center/Backups เสร็จสมบูรณ์`);
}

main().catch(err => {
  console.error('เกิดข้อผิดพลาด:', err);
  process.exit(1);
});
