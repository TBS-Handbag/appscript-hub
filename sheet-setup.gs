/**
 * Định dạng Google Sheet "Apps Script Hub - Danh sách link" cho dễ nhập.
 * Chạy 1 lần trong Tiện ích → Apps Script. Chạy lại cũng được, không mất dữ liệu.
 * Trang tính đầu tiên (gid=0) là trang được xuất bản ra CSV cho web đọc, nên giữ nó ở vị trí đầu.
 */
function dinhDangSheet() {
  const ss = SpreadsheetApp.getActive();
  const list = ss.getSheets()[0];
  list.setName('Danh sách');

  // Trang "Nhóm": danh sách nhóm cho dropdown, muốn thêm nhóm mới thì thêm dòng ở đây
  let groups = ss.getSheetByName('Nhóm');
  if (!groups) {
    groups = ss.insertSheet('Nhóm', 1);
    const existing = list.getLastRow() > 1
      ? [...new Set(list.getRange(2, 1, list.getLastRow() - 1, 1).getValues().flat().filter(String))]
      : [];
    const seed = existing.length ? existing : ['Sản xuất'];
    groups.getRange(1, 1).setValue('Tên nhóm');
    groups.getRange(2, 1, seed.length, 1).setValues(seed.map(g => [g]));
  }
  styleHeader_(groups.getRange(1, 1));
  groups.setFrozenRows(1);
  groups.setColumnWidth(1, 220);
  trimColumns_(groups, 1);

  // Trang "Danh sách"
  const HEADERS = ['Nhóm', 'Tên hiển thị', 'Mô tả ngắn', 'Link'];
  trimColumns_(list, HEADERS.length);
  list.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  list.setFrozenRows(1);
  list.setRowHeight(1, 36);
  [150, 220, 320, 460].forEach((w, i) => list.setColumnWidth(i + 1, w));

  const rows = list.getMaxRows() - 1;
  const body = list.getRange(2, 1, rows, HEADERS.length);
  body.setVerticalAlignment('middle').setFontSize(11);
  list.getRange(2, 1, rows, 3).setWrap(true);
  list.getRange(2, 4, rows, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setFontColor('#2F5D8A');
  list.getRange(2, 2, rows, 1).setFontWeight('bold');

  list.getBandings().forEach(b => b.remove());
  list.getRange(1, 1, list.getMaxRows(), HEADERS.length)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
    .setHeaderRowColor('#1C2B4A').setFirstRowColor('#FFFFFF').setSecondRowColor('#F3F5F1');
  styleHeader_(list.getRange(1, 1, 1, HEADERS.length));

  // Dropdown nhóm lấy từ trang "Nhóm" để không gõ sai chính tả thành 2 nhóm
  list.getRange(2, 1, rows, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(groups.getRange('A2:A'), true)
    .setAllowInvalid(false)
    .setHelpText('Chọn nhóm có sẵn. Muốn thêm nhóm mới thì thêm vào trang "Nhóm".')
    .build());
  // Link phải là URL, web chỉ hiện link bắt đầu bằng http(s)
  list.getRange(2, 4, rows, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireTextIsUrl()
    .setAllowInvalid(false)
    .setHelpText('Dán link đầy đủ, bắt đầu bằng https://')
    .build());

  list.getRange('A1').setNote('Các script cùng nhóm sẽ nằm chung một khu trên web. Thêm nhóm mới ở trang "Nhóm".');
  list.getRange('B1').setNote('Tên hiện trên ô. Viết ngắn, dễ nhận ra. Bắt buộc.');
  list.getRange('C1').setNote('Một câu mô tả script làm gì. Có thể để trống.');
  list.getRange('D1').setNote('Link web app (…/exec) hoặc link mở editor. Bắt buộc.');

  ss.setActiveSheet(list);
}

function styleHeader_(range) {
  range.setBackground('#1C2B4A').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(11).setVerticalAlignment('middle');
}

function trimColumns_(sheet, keep) {
  const extra = sheet.getMaxColumns() - keep;
  if (extra > 0) sheet.deleteColumns(keep + 1, extra);
}
