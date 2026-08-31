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
const { initializeApp, cert } = require('firebase-admin/app');
const { getFirestore } = require('firebase-admin/firestore');
const { google } = require('googleapis');
const XLSX = require('xlsx');
const fs = require('fs');

const serviceAccountKey = JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT_KEY);
const DRIVE_FOLDER_ID = process.env.DRIVE_FOLDER_ID;

// Firestore ยังใช้ Service Account เหมือนเดิม (จุดนี้ไม่มีปัญหา) — เปลี่ยนเฉพาะฝั่ง Drive
const firebaseApp = initializeApp({ credential: cert(serviceAccountKey) });
const db = getFirestore(firebaseApp);

// Drive ต้องใช้สิทธิ์ในนามบัญชี Gmail จริง (ไม่ใช่ Service Account) เพราะ Service Account
// ไม่มีโควต้าพื้นที่เก็บข้อมูลเป็นของตัวเอง อัปโหลดไฟล์ใหม่ไม่ได้ (403)
const oauth2Client = new google.auth.OAuth2(
  process.env.GOOGLE_OAUTH_CLIENT_ID,
  process.env.GOOGLE_OAUTH_CLIENT_SECRET,
  'http://localhost:53682'
);
oauth2Client.setCredentials({ refresh_token: process.env.GOOGLE_OAUTH_REFRESH_TOKEN });
const drive = google.drive({ version: 'v3', auth: oauth2Client });

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

  const errorRows = [];
  errors.forEach(e => {
    if ((e.solutions || []).length === 0) {
      errorRows.push({ ประเภท: e.category, รุ่น: e.model, ErrorCode: e.errorCode || '', อาการ: e.symptom, เปิดดู: e.viewCount || 0, ยอดโหวตแก้ได้จริง: e.totalUps || 0, ลำดับวิธี: '', คำอธิบาย: '', อุปกรณ์: '', อะไหล่: '', เวลาโดยประมาณ: '', จำนวนรูป: '', จำนวนโหวตวิธีนี้: '' });
    } else {
      e.solutions.forEach((s, idx) => {
        errorRows.push({ ประเภท: e.category, รุ่น: e.model, ErrorCode: e.errorCode || '', อาการ: e.symptom, เปิดดู: e.viewCount || 0, ยอดโหวตแก้ได้จริง: e.totalUps || 0, ลำดับวิธี: idx + 1, คำอธิบาย: s.desc, อุปกรณ์: s.tools || '', อะไหล่: s.parts || '', เวลาโดยประมาณ: s.time || '', จำนวนรูป: (s.images || []).length, จำนวนโหวตวิธีนี้: (s.feedback || []).filter(f => f.type === 'up').length });
      });
    }
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(errorRows), 'Errors');

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

  console.log('สร้างไฟล์ backup เสร็จแล้ว กำลังอัปโหลดขึ้น Google Drive (ในนามบัญชี Gmail จริง)...');

  // ----- อัปโหลดเข้า Google Drive (ใช้ drive client ที่สร้างไว้ด้านบนสุดของไฟล์แล้ว) -----

  async function findOrCreateFolder(name, parentId){
    const searchRes = await drive.files.list({
      q: `'${parentId}' in parents and name='${name}' and mimeType='application/vnd.google-apps.folder' and trashed=false`,
      fields: 'files(id, name)',
    });
    if (searchRes.data.files.length > 0) return searchRes.data.files[0].id;
    const folderRes = await drive.files.create({
      resource: { name, mimeType: 'application/vnd.google-apps.folder', parents: [parentId] },
      fields: 'id',
    });
    console.log(`สร้างโฟลเดอร์ "${name}" ใหม่แล้ว`);
    return folderRes.data.id;
  }

  const backupsFolderId = await findOrCreateFolder('Backups', DRIVE_FOLDER_ID);
  const dateFolderId = await findOrCreateFolder(dateStamp, backupsFolderId); // แยกเป็นโฟลเดอร์ตามวันที่ก่อน

  await drive.files.create({
    resource: { name: jsonFilename, parents: [dateFolderId] },
    media: { mimeType: 'application/json', body: fs.createReadStream(jsonFilename) },
  });
  await drive.files.create({
    resource: { name: excelFilename, parents: [dateFolderId] },
    media: { mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', body: fs.createReadStream(excelFilename) },
  });

  console.log(`✓ อัปโหลด Backup วันที่ ${dateStamp} เข้า Service Center/Backups/${dateStamp} เสร็จสมบูรณ์`);

  await cleanupOldBackups(backupsFolderId, BACKUP_RETENTION_DAYS);
}

// ----- ลบไฟล์ Backup เก่าที่เก็บไว้เกิน N วัน กันไฟล์สะสมเยอะเกินไป -----
const BACKUP_RETENTION_DAYS = 7;
async function cleanupOldBackups(folderId, daysToKeep){
  console.log(`กำลังตรวจสอบไฟล์ Backup เก่าที่เก็บไว้เกิน ${daysToKeep} วัน...`);
  const cutoff = new Date(Date.now() - daysToKeep * 24 * 60 * 60 * 1000).toISOString();
  const res = await drive.files.list({
    q: `'${folderId}' in parents and trashed=false and createdTime < '${cutoff}'`,
    fields: 'files(id, name, createdTime)',
    pageSize: 1000,
  });
  const oldFiles = res.data.files || [];
  if (oldFiles.length === 0) { console.log('ไม่มีไฟล์เก่าที่ต้องลบ'); return; }
  for (const f of oldFiles) {
    try {
      await drive.files.delete({ fileId: f.id });
      console.log(`ลบไฟล์เก่าแล้ว: ${f.name} (สร้างเมื่อ ${f.createdTime})`);
    } catch (e) {
      console.error(`ลบไฟล์ไม่สำเร็จ: ${f.name}`, e.message);
    }
  }
  console.log(`✓ ลบไฟล์ Backup เก่าไปทั้งหมด ${oldFiles.length} ไฟล์`);
}

// ----- บันทึกสถานะการ backup ล่าสุด (สำเร็จ/ไม่สำเร็จ + เวลา) ไว้ในไฟล์เล็กๆ ให้หน้าเว็บอ่านไปแสดงผลได้ -----
// ไม่ใช้ Firestore เพราะ Service Account ตั้งใจให้เป็น "อ่านอย่างเดียว" เพื่อความปลอดภัย (ถ้า Key หลุดจะได้เขียน/ลบอะไรไม่ได้)
async function writeStatusFile(success, errorMessage){
  try{
    const searchRes = await drive.files.list({
      q: `'${DRIVE_FOLDER_ID}' in parents and name='Backups' and mimeType='application/vnd.google-apps.folder' and trashed=false`,
      fields: 'files(id, name)',
    });
    let backupsFolderId;
    if (searchRes.data.files.length > 0) {
      backupsFolderId = searchRes.data.files[0].id;
    } else {
      const folderRes = await drive.files.create({
        resource: { name: 'Backups', mimeType: 'application/vnd.google-apps.folder', parents: [DRIVE_FOLDER_ID] },
        fields: 'id',
      });
      backupsFolderId = folderRes.data.id;
    }

    const statusContent = JSON.stringify({
      success,
      timestamp: Date.now(),
      dateStamp: new Date().toISOString().slice(0, 10),
      error: errorMessage || null,
    }, null, 2);

    const statusSearch = await drive.files.list({
      q: `'${backupsFolderId}' in parents and name='backup-status.json' and trashed=false`,
      fields: 'files(id)',
    });
    const media = { mimeType: 'application/json', body: statusContent };
    if (statusSearch.data.files.length > 0) {
      await drive.files.update({ fileId: statusSearch.data.files[0].id, media });
    } else {
      await drive.files.create({ resource: { name: 'backup-status.json', parents: [backupsFolderId] }, media });
    }
    console.log(`บันทึกสถานะ backup-status.json แล้ว (success=${success})`);
  } catch (e) {
    console.error('บันทึกไฟล์สถานะไม่สำเร็จ (ไม่กระทบผลลัพธ์หลักของการ backup):', e.message);
  }
}

main()
  .then(() => writeStatusFile(true, null))
  .catch(async err => {
    console.error('เกิดข้อผิดพลาด:', err);
    await writeStatusFile(false, err.message);
    process.exit(1);
  });
