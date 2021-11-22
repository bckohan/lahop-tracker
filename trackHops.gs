/**
 * The following field needs to be set to the ID of the spreadsheet to sync to.
 */
var HOP_SHEET_ID = '1vfzkKmBUYTy5fn-L4maZXEfN8hlAX1kQDYsvgblWxQA'; // the google spreadsheet ID

// this is the search string as you would type it into Gmail search
var HOP_QUERY = 'from: donotreply@lahsa.org subject: Outreach Request';
var SUBJ_RE = /\(Request ID: (\d+)\)/;
var SHEET_NAME = 'HOPS';
var LAST_SYNC = null;
var HOP_SHEET = null;
var FWD_SHEET = null;
var VERSION = '0.0.2';
var LAST_VERSION = null;

var columns = {
    "LA-HOP ID": null,
    'Name': null,
    "Origin Date": null,
    'Vol': null,
    "Vol Phone": null,
    "Submit Email": null,
    "#PPL": null,
    "Address": null,
    "Location Description": null,
    "Physical Description": null,
    "Needs Description": null,
    "Status": null,
    "Resolve Date": null
};

var col_map = {
    "LA-HOP ID": 'id',
    'Name': 'name',
    "Origin Date": 'time',
    'Vol': 'vol',
    "Vol Phone": 'phone',
    "Submit Email": 'email',
    "#PPL": 'num',
    "Address": 'addr',
    "Location Description": 'location',
    "Physical Description": 'desc',
    "Needs Description": 'needs',
    "Status": 'status',
    "Last Update": 'last_update'
};

// reverse lookup of col_map - all this is very ugly but whatever
var attr_map = {};
for ( const [col, attr] of Object.entries(col_map) ) {
  attr_map[attr] = col;
}

// component regexes
var attr_re = {
    'addr': /<p><strong>Address: <\/strong>(.*)<\/p>/,
    'location': /<p><strong>Description of location: <\/strong>(.*)<\/p>/,
    'time': /<p><strong>Date last seen: <\/strong>(.*)<\/p>/,
    'num': /<p><strong>Number of people: <\/strong>(\d+)<\/p>/,
    'name': /<p><strong>Name of person\/people requiring outreach: <\/strong>(.*)<\/p>/,
    'desc': /<p><strong>Physical description of person\/people: <\/strong>(.*)<\/p>/,
    'needs': /<p><strong>Description of person\/people's needs: <\/strong>(.*)<\/p>/,
    'vol': /<p><strong>Your name: <\/strong>(.*)<\/p>/,
    'org': /<p><strong>Company\/organization: <\/strong>(.*)<\/p>/,
    'vol_desc': /<p><strong>How would you describe yourself? <\/strong>(.*)<\/p>/,
    'email': /<p><strong>Email: <\/strong>(.*)<\/p>/,
    'phone': /<p><strong>Phone: <\/strong>(.*)<\/p>/
};
function columnToLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function letterToColumn(letter) {
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

//DATE_TIME_RE = /(\d{1,2})\/(\d{1,2})\/(\d{4}) (\d{1,2}):(\d{1,2})(.*) (A|P)M/;

const NotificationType = Object.freeze({"NEW":1, "UPDATE":2, "RECEIVED":3, "UNKNOWN": 4});
const Status = Object.freeze({"UNRESOLVED":1, "NO_CONTACT":2, "CONTACT":3, "REDUNDANT": 4});

function getGlobalMeta() {
    for (const meta of HOP_SHEET.createDeveloperMetadataFinder().withKey('sync_time').find()) {
      LAST_SYNC = Date.parse(meta.getValue());
    }
    for (const meta of HOP_SHEET.createDeveloperMetadataFinder().withKey('version').find()) {
      LAST_VERSION = meta.getValue();
    }
}

// Main function, the one that you must select before run
function findHOPs() {

    console.log(`Searching for: "${HOP_QUERY}"`);

    getSheet();
    getGlobalMeta();
    let lastProcessed = null;
    let newestTime = null;
    let start = 0;
    let max = 500;

    let threads = GmailApp.search(HOP_QUERY, start, max);

    let hops = {};
    while (threads.length>0) {
        for (var i in threads) {
            var thread=threads[i];
            var msgs = threads[i].getMessages();
            newestTime = null;
            for (var j in msgs) {
                if (LAST_VERSION === VERSION && LAST_SYNC && msgs[j].getDate() <= LAST_SYNC) { continue; }
                if (newestTime === null || msgs[j].getDate() > newestTime) {
                    newestTime = msgs[j].getDate();
                }
                let msgAttrs = processMessage(msgs[j]);
                if (!hops.hasOwnProperty(msgAttrs.id)) {
                    hops[msgAttrs.id] = msgAttrs;
                } else {
                    for (const attr in msgAttrs) {
                        if (attr === 'last_update') {
                            if (msgAttrs[attr] > hops[msgAttrs.id][attr]) {
                                hops[msgAttrs.id][attr] = msgAttrs[attr];
                                hops[msgAttrs.id]['status'] = msgAttrs['status'];
                            }
                        }
                        else if (attr === 'status') { continue; }
                        else if (msgAttrs[attr] !== null) {
                            hops[msgAttrs.id][attr] = msgAttrs[attr];
                        }
                    }
                }
            }
            if (newestTime === null) { break; }  // no messages occurred after LAST_SYNC
        }
        if (newestTime === null) { break; }
        start = start + max;
        threads = GmailApp.search(HOP_QUERY, start, max);
    }

    console.info(hops);
    writeHOPs(hops);
}

function setAttribute(re, line, parts, attr) {
    let mtch = re.exec(line);
    if (mtch !== null && mtch.length >= 2) {
        parts[attr] = mtch[1];
        return true;
    }
    return false;
}

function processMessage(msg) {
    let parts = {
        'last_update': msg.getDate(),
        'type': NotificationType.UNKNOWN,
        'status': Status.UNRESOLVED
    };
    for (const key in attr_re) {
        parts[key] = null;
    }
    let sub = SUBJ_RE.exec(msg.getSubject());
    if (sub === null || sub.length < 2) {
        console.log('Unable to interpret subject line!');
        return null;
    }
    if (msg.getSubject().includes('New')) {
        parts.type = NotificationType.NEW;
    } else if (msg.getSubject().includes('Update')) {
        parts.type = NotificationType.UPDATE;
    } else if (msg.getSubject().includes('received')) {
        parts.type = NotificationType.RECEIVED;
    }
    parts.id = parseInt(sub[1]);
    let lines = msg.getBody().split('\n');
    for (let idx=0; idx < lines.length; idx++) {
        for (const attr in attr_re) {
            if (setAttribute(attr_re[attr], lines[idx], parts, attr)) continue;
        }
        if (parts.type !== NotificationType.NEW && parts.status === Status.UNRESOLVED) {
            if (lines[idx].includes('unable to make contact after two attempts')) {
                parts.status = Status.NO_CONTACT;
            } else if (lines[idx].includes('made contact with the individual')) {
                parts.status = Status.CONTACT;
            } else if (lines[idx].includes('already serving the area listed')) {
                parts.status = Status.REDUNDANT;
            }
        }
    }
    if (parts.time !== null && typeof parts.time === 'string') {
        /*let dt = DATE_TIME_RE.exec(parts.time);
        parts.time = new Date(
            parseInt(dt[3]),
            parseInt(dt[1])-1,
            parseInt(dt[2]),
            parseInt(dt[4]) + (dt[6] === 'P' ? 12 : 0),
            parseInt(dt[5])
        );*/
        // timezone might be wrong here, but thats ok we only use this for ordering, TZ not displayed in sheet
        parts.time = new Date(parts.time);
    }
    return parts;
}

// Add contents to sheet
function getSheet() {
    var hop_ss = SpreadsheetApp.openById(HOP_SHEET_ID);
    HOP_SHEET = hop_ss.getSheetByName(SHEET_NAME);
    if (HOP_SHEET === null) {
        HOP_SHEET = hop_ss.insertSheet(SHEET_NAME);
    }
    addHeaders();
}

function addHeader(name, idx) {
  HOP_SHEET.getRange(`${columnToLetter(idx)}1:${columnToLetter(idx)}1`).setValue(name);
  HOP_SHEET.getRange(`${columnToLetter(idx)}:${columnToLetter(idx)}`).addDeveloperMetadata('col_name', name);
  columns[name] = idx;
}

function addHeaders() {
  /**
   * Add in the column headers if they dont exist by searching for their meta. Record column numbers for each header and add any unfound columns onto the end.
   */
    var last_col = HOP_SHEET.getLastColumn();
    HOP_SHEET.setFrozenRows(1);
    var idx = 1;
    for (const val in HOP_SHEET.getRange('1:1').getValues()[0]) {
      var meta = HOP_SHEET.getRange(`${columnToLetter(idx)}:${columnToLetter(idx)}`).createDeveloperMetadataFinder().withKey('col_name').find();
      if (meta.length > 0) {
        var col_name = meta[0].getValue();
        columns[col_name] = idx;
      }
      idx++;
      if (idx > last_col) {
        break;
      }
    }
    for (const [name, col] of Object.entries(columns)) {
      if (col == null) {
        addHeader(name, idx);
        idx++;
      }
    }
}

function writeHOPs(hops) {
    //metadata isnt a simple key value store because google's spreadsheets API suuuuucks
    for (const meta of HOP_SHEET.createDeveloperMetadataFinder().withKey('sync_time').find()) { meta.remove(); }
    for (const meta of HOP_SHEET.createDeveloperMetadataFinder().withKey('version').find()) { meta.remove(); }
    HOP_SHEET.addDeveloperMetadata('sync_time', Utilities.formatDate(new Date(), 'America/Los_Angeles', 'MMMM dd, yyyy HH:mm:ss Z'));
    HOP_SHEET.addDeveloperMetadata('version', VERSION);

    // first pass, search for HOPs by ID in sheet and update the found ones

    // second pass, add remaining HOPs to top of sheet in descending date order of creation
}
