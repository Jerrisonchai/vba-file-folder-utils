# Test Cases
This document defines test cases for validating all utilities in the repo. Each case ensures both normal and edge conditions are covered.

---

## 1. CodeCounter/
**Test Cases**
- [ ] Count modules with 5 lines of code → expect 5 total / 5 executable.
- [ ] Count module with comments → exclude `'` lines.
- [ ] Empty module → expect 0 lines.
- [ ] Large project (500+ modules) → expect no crash.

---

## 2. ConvertToTxtCsv/
**Test Cases**
- [ ] Export range `A1:C10` → valid `.csv` with 10 rows.
- [ ] Export large range (100k rows) → file opens in Notepad++.
- [ ] Invalid path → script errors gracefully.
- [ ] Special chars (€, 中文, 😊) → preserved in UTF-8.

---

## 3. CountAndRenameFolders/
**Test Cases**
- [ ] Count 5 subfolders in test dir → expect result = 5.
- [ ] Rename folders by prefix rule → names updated.
- [ ] No permission folder → skipped gracefully.
- [ ] Deep nested structure → still counts accurately.

---

## 4. CreateFoldersEmailing/
**Test Cases**
- [ ] Input list of 3 users → 3 folders created.
- [ ] Duplicate user → folder not recreated.
- [ ] Invalid path → handled with error msg.
- [ ] 1,000+ folders → created within 10 sec.

---

## 5. LoopListFolderItems/
**Test Cases**
- [ ] Directory with 10 files → metadata listed in "Data".
- [ ] Directory with subfolders → still lists files.
- [ ] Empty folder → "No items found".
- [ ] 50,000+ files → completes without crash.

---

## 6. RenameAllFiles/
**Test Cases**
- [ ] Rename `test1.pdf` → `invoice1.pdf` works.
- [ ] Rename multiple files with mapping in `Data` sheet.
- [ ] File in use (locked) → skipped gracefully.
- [ ] 5,000+ files → processed in < 5 sec.

---

## 7. SplitSheetsToFiles/
**Test Cases**
- [ ] Workbook with 3 sheets → 3 new files created.
- [ ] Sheet with special chars → filename sanitised.
- [ ] Destination folder missing → created automatically.
- [ ] Large workbook (200 sheets) → splits successfully.

---

## 🔑 Notes
- Use **dummy test data** (not production files).  
- Validate outputs with **manual check + automated checksum**.  
- Edge cases (permissions, locked files, invalid chars) must be tested.  
