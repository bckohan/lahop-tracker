/**
 * The following field(s) should be configured to your liking.
 */
var HOP_SHEET_ID = '1vfzkKmBUYTy5fn-L4maZXEfN8hlAX1kQDYsvgblWxQA'; // the google spreadsheet ID of the sheet to hold the hop log and forward table
var SHEET_NAME = 'HOPS';     // name to use for the sheet containing the HOP log
var FWD_SHEET = 'Forwards';  // this should be set to the name of the sheet in the spreadsheet that holds the forward table
var HOP_LABEL = 'HOP';       // tag HOP emails with this label in gmail
///////////////////////////////////////////////////////////////////////


// DONT CHANGE ANYTHING BELOW THIS UNLESS YOU KNOW WHAT YOURE DOING
// this is the search string as you would type it into Gmail search
var HOP_QUERY = 'subject: Outreach Request'; //'from: donotreply@lahsa.org subject: Outreach Request';
var SUBJ_RE = /\(Request ID: (\d+)\)/;
var LAST_SYNC = null;
var HOP_SHEET = null;
var VERSION = 2;  // When the version is incremented all HOPs on the account will be reprocessed!
var LAST_VERSION = null;

var accountEmails = [Session.getActiveUser().getEmail()].concat(GmailApp.getAliases());

function validateEmail(email) {
  /**
   * Confirm if a string is valid email address and return it if it is. Otherwise, return null
   */
  const re = /^(([^<>()[\]\\.,;:\s@\"]+(\.[^<>()[\]\\.,;:\s@\"]+)*)|(\".+\"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
  if (re.test(email)) {
    return email;
  }
  return null;
}

var char_lookup = {
  '&amp;': '&',
  '&quot;': '"',
  '&nbsp;': ' ',
  '&apos;': "'",
  '&lt;': '<',
  '&gt;': '<'
}
function entityToChar(ent){
  /**
   * Convert &#; html codes to characters. There's no easy way to do this in google apps script without including other libraries because all of the DOM
   * libs are not present. We do a passable job here that works for all #s and some special cases. If we fail we just punt - its not that big a deal.
   */
  if (ent.includes('#')) {
    return String.fromCharCode(ent.slice(2,-1));
  } else if (char_lookup.hasOwnProperty(ent)) {
    return char_lookup[ent];
  }
  return ent;
}

function validateTextField(text) {
  var entities = text.match(/&[^;]*;/g);
  if (entities) {
    for (const ent of entities) {
      text = text.replace(ent, entityToChar(ent));
    }
  }
  return text;
}

// map column types to column indexes in the spreadsheet log - these are discovered dynamically allowing people to add additional columns wherever
var columns = {
    "LA-HOP ID": null,
    "Status": null,
    'Name': null,
    "Origin Date": null,
    "Last Update": null,
    "#PPL": null,
    "Address": null,
    "Location Description": null,
    "Physical Description": null,
    "Needs Description": null,
    'Vol': null,
    "Vol Phone": null,
    "Submit Email": null,
    "FWD": null
};

// map column header names in the spreadsheet to hop object attributes - these could be the same, but meh lazy
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
    "Last Update": 'last_update',
    "FWD": 'fwd'
};

// reverse lookup of col_map - all this is very ugly but whatever
var attr_map = {};
for ( const [col, attr] of Object.entries(col_map) ) {
  attr_map[attr] = col;
}

// component regexes
// Can be of the format attribute name: [ regex1, regex2 ] and/or attribute name: [ [regex, validate func], ... ]
// if using a validation function, it must return null if validation failed and the final value to use if validation was successful
var attr_re = {
    'addr': [[/<p><strong>Address: <\/strong>(.*)<\/p>/, validateTextField]],
    'location': [[/<p><strong>Description of location: <\/strong>(.*)<\/p>/, validateTextField]],
    'time': [/<p><strong>Date last seen: <\/strong>(.*)<\/p>/],
    'num': [/<p><strong>Number of people: <\/strong>(\d+)<\/p>/],
    'name': [[/<p><strong>Name of person\/people requiring outreach: <\/strong>(.*)<\/p>/, validateTextField]],
    'desc': [[/<p><strong>Physical description of person\/people: <\/strong>(.*)<\/p>/, validateTextField]],
    'needs': [[/<p><strong>Description of person\/people's needs: <\/strong>(.*)<\/p>/, validateTextField]],
    'vol': [[/<p><strong>Your name: <\/strong>(.*)<\/p>/, validateTextField]],
    'org': [[/<p><strong>Company\/organization: <\/strong>(.*)<\/p>/, validateTextField]],
    'vol_desc': [[/<p><strong>How would you describe yourself? <\/strong>(.*)<\/p>/, validateTextField]],
    'email': [
      [/<p><strong>Email: <\/strong>(.*)<\/p>/, validateEmail],
      [/<p><strong>Email: <\/strong><a.*>(.*)<\/a><\/p>/, validateEmail]
    ],
    'phone': [/<p><strong>Phone: <\/strong>(.*)<\/p>/]
};

var fwd_table = {};

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
const Status = Object.freeze({"UNRESOLVED":1, "NO_CONTACT":2, "CONTACT":3, "DISMISSED": 4});

function getGlobalMeta() {
    for (const meta of HOP_SHEET.createDeveloperMetadataFinder().withKey('sync_time').find()) {
      LAST_SYNC = Date.parse(meta.getValue());
    }
    for (const meta of HOP_SHEET.createDeveloperMetadataFinder().withKey('version').find()) {
      LAST_VERSION = parseInt(meta.getValue());
    }
}

function buildForwardTable() {
  /**
   * Build the forwarding table from the spreadsheet. This table is used to make sure HOP submitters still get the messages coming from the HOP system.
   * HOP submitters use their own phone number when submitting a HOP. Messages will be forwarded based on the registered phone number. So the forward
   * table maps phone numbers to email addresses. Anytime a new hop message is directly received from the system this table will be checked and if the
   * submitters phone number is registered, the HOP message will be forwarded to them.
   *
   * Forward table should have a header row with at least one column with the word Phone and one column with the word Email. Ex:
   * Row
   *   0 | Phone      | Email
   *   1 | 5558183333 | myemail@example.com
   *   ...
   *   n | 5553234444 | someoneelse@example.com
   *
   *  Phone numbers should be numeric only with no spaces. Its fine to add other columns to the forward table for book keeping.
   */
    var hop_ss = SpreadsheetApp.openById(HOP_SHEET_ID);
    var fwdSheet = hop_ss.getSheetByName(FWD_SHEET);
    if (fwdSheet != null) {
      var values = fwdSheet.getDataRange().getValues();
      if (values.length > 1) {
        var phoneCol = null;
        var emailCol = null;
        idx = 0;
        for (const hdr of values[0]) {
          if (phoneCol == null && hdr.toLowerCase().includes('phone') || hdr.toLowerCase().includes('number')) {
            phoneCol = idx;
          } else if (emailCol == null && hdr.toLowerCase().includes('email')) {
            emailCol = idx;
          }
          if (emailCol != null && phoneCol != null) {
            for (const fwd of values.slice(1)) {
              var phone = parseInt(fwd[phoneCol]);
              var email = validateEmail(fwd[emailCol]);
              if (phone && email) {
                fwd_table[phone] = email;
              } else {
                console.warn(`Forward table entry: ${fwd} is invalid.`);
              }
            }
            break;
          }
          idx++;
        }
      }
    }
    console.info(`Forward Table: ${fwd_table}`);
}

// Main function, the one that you must select before run
function findHOPs() {

    console.log(`Searching for: "${HOP_QUERY}"`);

    getSheet();
    var hopLabel = GmailApp.createLabel(HOP_LABEL);
    getGlobalMeta();
    buildForwardTable();
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
                if (LAST_VERSION >= VERSION && LAST_SYNC && msgs[j].getDate() <= LAST_SYNC) { continue; }
                if (newestTime === null || msgs[j].getDate() > newestTime) {
                    newestTime = msgs[j].getDate();
                }
                let msgAttrs = processMessage(msgs[j]);
                if (msgAttrs) {
                  if (msgAttrs.id === 31234) {
                    console.log(msgAttrs);
                  }
                  thread.addLabel(hopLabel);
                  msgAttrs.fwd = !accountEmails.includes(msgAttrs.email);
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
                  // only forward if the following conditions are met:
                  // 1) we have a registered forward address for the phone #
                  // 2) the message itself was not forwarded but came straight to our HOP reception account
                  // 3) the spreadsheet has been previously synced (don't want lots of errant forwards initially)
                  // 4) the message is newer than the last synchronization time
                  if (!msgAttrs.fwd && LAST_SYNC && msgs[j].getDate() > LAST_SYNC) {
                    if (fwd_table.hasOwnProperty(msgAttrs.phone)) {
                      msgs[j].forward(fwd_table[msgAttrs.phone]);
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

function setAttribute(regexes, line, parts, attr) {
  /**
   * The core parsing logic. For the line, run through all regexes from the attr_re table and if any match and pass validation set that
   * attribute on the message and return true for success. If the line matches no attributes return false.
   */
    for (var rex of regexes) {
      var validate = null;
      if (Array.isArray(rex)) {
        validate = rex[1];
        rex = rex[0];
      }
      let mtch = rex.exec(line);
      if (mtch !== null && mtch.length >= 2) {
          if (validate) {
            parts[attr] = validate(mtch[1]);
          } else {
            parts[attr] = mtch[1];
          }
          if (parts[attr] == null) continue;
          return true;
      }
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
                parts.status = Status.DISMISSED;
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
      if (idx > last_col) {
        break;
      }
      idx++;
    }
    for (const [name, col] of Object.entries(columns)) {
      if (col == null) {
        addHeader(name, idx);
        idx++;
      }
    }
}

function toRow(hop) {
  /**
   * Convert parsed HOP object to an array suitable (as in, indexed correctly) for upload to the sheet.
   * TODO test if nulls overwrite custom columns
   */
  var row = [];
  for (const [attr, val] of Object.entries(hop)) {
    var idx = columns[attr_map[attr]];
    while (row.length < idx) {
      row.push(null);
    }
    row[idx-1] = val;
  }
  return row;
}

function writeHOPs(hops) {

  if (hops.length == 0) {
    return;
  }

  //metadata isnt a simple key value store because google's spreadsheets API suuuuucks
  for (const meta of HOP_SHEET.createDeveloperMetadataFinder().withKey('sync_time').find()) { meta.remove(); }
  for (const meta of HOP_SHEET.createDeveloperMetadataFinder().withKey('version').find()) { meta.remove(); }
  HOP_SHEET.addDeveloperMetadata('sync_time', Utilities.formatDate(new Date(), 'America/Los_Angeles', 'MMMM dd, yyyy HH:mm:ss Z'));
  HOP_SHEET.addDeveloperMetadata('version', VERSION);

  // first pass, search for HOPs by ID in sheet and update the found ones
  var row = 1;
  for (const hop_id of HOP_SHEET.getRange(`${columnToLetter(columns['LA-HOP ID'])}:${columnToLetter(columns['LA-HOP ID'])}`).getValues()[0]) {
    if (hops.hasOwnProperty(hop_id.toString())) {
      HOP_SHEET.setValues([toRow(hops[hop_id.toString()])]);
      delete hops[hop_id.toString()];
    }
    row++;
  }

  // second pass, add remaining HOPs to top of sheet in descending date order of creation
  var rows = [];
  for (const [hop_id, hop] of Object.entries(hops)) {
    rows.push(toRow(hop));
  }

  function submissionOrderDescending(row1, row2) {
    if (row1[columns['Origin Date']] < row2[columns['Origin Date']]) {
      return 1;
    }
    else if (row1[columns['Origin Date']] > row2[columns['Origin Date']]) {
      return -1;
    }
    return 0;
  }
  rows.sort(submissionOrderDescending);

  // determine if we got any out of order forwarded HOPs that may need to be inserted in order between already logged HOPs


  if (rows.length > 0) {
    var insertRange = HOP_SHEET.getRange(`A2:${columnToLetter(HOP_SHEET.getLastColumn())}${rows.length+1}`);
    insertRange.insertCells(SpreadsheetApp.Dimension.ROWS);
    insertRange.setValues(rows);
  }
}
