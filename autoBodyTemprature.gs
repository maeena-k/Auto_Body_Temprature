function autoBodyTemprature() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Body Temprature');

  let max_row = sheet.getMaxRows();
  let max_col = sheet.getMaxColumns();
  
  // clear the exisinting content
  if (max_row > 5) {
    sheet.getRange(5, 1, max_row - 4, max_col).clearContent();
  };
  
  sheet.getRange(4,2,1,max_col-1).clearContent();
  
  // calculate the number of days in the required period
  let start_date = new Date(sheet.getRange('B1').getValue());
  let end_date = new Date(sheet.getRange('B2').getValue());
  
  let diff_day;
  
  if (end_date >= start_date){
    diff_day = (end_date - start_date) / (24*60*60*1000);
  } else {
    diff_day = (start_date - end_date) / (24*60*60*1000);
    start_date = end_date;
  };
  
  // get the number of members
  const member_num = parseInt(sheet.getRange('B3').getValue());
  
  // delete or add rows and columns if needed
  if (diff_day + 5 < max_row){
    sheet.deleteRows(diff_day + 5, max_row - diff_day - 5);
  } else if (diff_day + 5 > max_row){
    sheet.insertRowsAfter(max_row, diff_day + 5 - max_row);
  };
  
  max_row = sheet.getMaxRows();
  
  if (member_num + 1 < max_col){
    sheet.deleteColumns(member_num + 1, max_col - member_num - 1);
  } else if (member_num + 1 > max_col){
    sheet.insertColumnsAfter(max_col, member_num + 1 - max_col);
  };
  
  max_col = sheet.getMaxColumns();
  
  // fix the column titles
  let person_list = []
  
  for (let j = 2; j <= max_col; j++) {
    person_list.push('PERSON' + (j - 1));
  };
  
  sheet.getRange(4, 2, 1, person_list.length).setValues([person_list]);
  
  // create random body temprature for each day and person
  let date_val, body_temprature;
  let date_list = [], temprature_list_l = [], temprature_list_s = [];
  
  for (let i = 0; i <= diff_day ; i++){
    date_val = new Date(start_date.getFullYear(), start_date.getMonth(), start_date.getDate() + i + 1);
    date_val = Utilities.formatDate(date_val, 'JST', 'yyyy-MM-dd');
    date_list.push([date_val]);

    for (let j = 2; j <= max_col; j++){
      body_temprature = Math.random() * 0.7 + 35.8;
      body_temprature = Math.round(body_temprature*10)/10;     
      temprature_list_s.push(body_temprature);
      if (j == max_col){
        temprature_list_l.push(temprature_list_s);
        temprature_list_s = [];
      }
    };
  };
  
  // paste values in thr range and fix format
  sheet.getRange(5, 1, date_list.length, 1).setValues(date_list);
  sheet.getRange(5, 2, temprature_list_l.length, temprature_list_l[0].length).setValues(temprature_list_l);
  sheet.getRange(5, 2, temprature_list_l.length, temprature_list_l[0].length).setNumberFormat('00.0');
}
