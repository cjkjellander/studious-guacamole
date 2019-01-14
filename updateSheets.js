// -*- mode: js; js-indent-level: 2; -*-
function updateSheets() {
  var ss = SpreadsheetApp.getActive();
  var inputSheet = findSheet(ss, 'DropTable');

  range_step = '10';
  range_max = '1250';

  var range = inputSheet.getRange(1, 1, 6, 17);

  var allInputData = range.getValues();

  hpa = parseInt(allInputData[0][14]);
  hgt_sgt = allInputData[1][16];
  hum = allInputData[0][16];
  temps = allInputData[4].slice(1, 10);
  mvs = allInputData[5].slice(1, 10);

  for (var i = 0; i < temps.length; i++) {
    temp = parseInt(temps[i]);
    var sheet = findSheet(ss, 'Temp'+temp);

    var all_data = queryJBM(hgt_sgt, range_step, range_max, hpa, hum, mvs[i], temp);
    var write_range = sheet.getRange(1,1, all_data.length, all_data[0].length);
    write_range.setValues(all_data);
  }
}

function findSheet(sheets, name) {
  var sheet_list = sheets.getSheets();
  for (var i = 0; i < sheet_list.length; i++) {
    if (sheet_list[i].getName().match(name)) {
      return sheet_list[i];
    }
  }
  return null;
}

function queryJBM(hgt_sgt, range_step, range_max, hpa, hum, mv, temp) {

  data = getData(range_step, range_max, mv, temp, hpa, hum, hgt_sgt);
  rows = parseData(data);

  ret = [];
  ret.push(
      [ 'range'
      , 'drop'
      , 'wind'
      , 'lead'
      , 'danger'
      , 'mach'
      ]);
  var printz = true;
  for (var i = 0; i < rows.length; i++) {
    if (!printz) {
      ret.push(['', '', '', '', '', '']);
    } else {
      ret.push(
        [ i*(range_step)
        , rows[i]['drop_angle_cell']
        , rows[i]['wind_angle_cell']
        , rows[i]['lead_angle_cell']
        , rows[i]['range_cell']
        , rows[i]['mach_cell']
        ]);
    }
    if (rows[i]['mach_cell'] < 1 && temp != 0) {
      printz = false;
    }
  }

  return ret;
}

function parseData(data) {
  row = /<TR CLASS="(data|zero)_row">(.+?)<\/TR>/gi;

  cell = /<TD CLASS=\"([^"]+)\">([^<]+)<\/TD>/gi;

  output = [];

  while (rows = row.exec(data)) {
    e = {};
    while (cells = cell.exec(rows[2])) {
      e[cells[1]] = cells[2];
    }
    output.push(e);
  }
  return output;
}

function getData(range_step, range_max, mv, temp, hpa, hum, hgt_sgt) {
 var formData =
{ "b_id.v":"-1"
, "bc.v":"0.274"
, "d_f.v":"4"
, "bt_wgt.v":"130"
, "bt_wgt.u":"23"
, "cal.v":"6.5"
, "cal.u":"9"
, "m_vel.v":mv
, "m_vel.u":"17"
, "ch_dst.v":"3"
, "ch_dst.u":"12"
, "hgt_sgt.v":hgt_sgt
, "hgt_sgt.u":"11"
, "ofs_sgt.v":"0.0"
, "ofs_sgt.u":"11"
, "hgt_zer.v":"0.0"
, "hgt_zer.u":"11"
, "ofs_zer.v":"0.0"
, "ofs_zer.u":"11"
, "azm.v":"0.0"
, "azm.u":"2"
, "ele.v":"0.0"
, "ele.u":"2"
, "los.v":"0.0"
, "cnt.v":"0.0"
, "spd_wnd.v":"5"
, "spd_wnd.u":"17"
, "ang_wnd.v":"90.0"
, "spd_tgt.v":"5"
, "spd_tgt.u":"17"
, "ang_tgt.v":"90.0"
, "siz_tgt.v":"30"
, "siz_tgt.u":"11"
, "rng_min.v":"0"
, "rng_max.v":range_max
, "rng_inc.v":range_step
, "rng_zer.v":"100"
, "tmp.v":temp
, "tmp.u":"18"
, "prs.v":hpa
, "prs.u":"21"
, "hum.v":hum
, "alt.v":"0.0"
, "alt.u":"12"
, "rad_vz.v":"12.5"
, "rad_vz.u":"11"
, "col_eng.v":"1"
, "col1_un.v":"1.00"
, "col1_un.u":"11"
, "col2_un.v":"1.00"
, "col2_un.u":"2"
, "cor_ele.v": "on"
, "rng_un.v":  "on"
, "def_cnt.v": "on"
//, "mrk_trs.v": "on"
, "inc_ds.v": "on"
 };

  var options = {
    'method': 'post',
    'payload': formData,
    'headers': {
      'Referer': 'http://www.jbmballistics.com/cgi-bin/jbmtraj-5.1.cgi',
      'Content-Type': 'application/x-www-form-urlencoded',
      'Origin': 'http://www.jbmballistics.com',
      'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
      'Accept-Encoding': 'gzip, deflate',
      'Accept-Language': 'en-US,en;q=0.8,sv;q=0.6',
      'Upgrade-Insecure-Requests': '1',
      'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3005.2 Safari/537.36',
      'Cache-Control': 'max-age=0',
    }
  };
  var url = 'http://www.jbmballistics.com/cgi-bin/jbmtraj-5.1.cgi';
  //url = 'https://httpbin.org/post';
  resp = UrlFetchApp.fetch(url, options);

  resp = resp.getContentText();
  resp = resp.replace(/\n/g, "");
  return resp;
}
