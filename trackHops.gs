/**
 * The following field(s) should be configured to your liking.
 */
var HOP_SHEET_ID = '1vfzkKmBUYTy5fn-L4maZXEfN8hlAX1kQDYsvgblWxQA'; // the google spreadsheet ID of the sheet to hold the hop log and forward table
var SHEET_NAME = 'HOPS';     // name to use for the sheet containing the HOP log
var FWD_SHEET = 'Forwards';  // this should be set to the name of the sheet in the spreadsheet that holds the forward table
var HOP_LABEL = GmailApp.createLabel('HOP');       // tag HOP emails with this label in gmail
///////////////////////////////////////////////////////////////////////


// DONT CHANGE ANYTHING BELOW THIS UNLESS YOU KNOW WHAT YOURE DOING - Parsing these emails is *tricky*, especially when trying to account for all the weird formatting issues different email clients introduce when messages are forwarded
// this is the search string as you would type it into Gmail search
// The BEST search for finding all HOPs on an email where that email is the account of record with LAHSA is:
//     'from: donotreply@lahsa.org subject: Outreach Request';
var HOP_QUERY = 'subject:"Outreach Request"';
//var HOP_QUERY = '{subject:"Outreach Request"  subject:hops}';
var SUBJ_RE = /\(Request ID: (\d+)\)/;
var LAST_SYNC = null;
var HOP_SHEET = null;
var VERSION = 1;  // When the version is incremented all HOPs on the account will be reprocessed!
var LAST_VERSION = null;

var accountEmails = [Session.getActiveUser().getEmail()].concat(GmailApp.getAliases());
var hops = {};

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
    "Last Seen": null,
    "Submit Date": null,
    "Last Update": null,
    'Vol': null,
    "Vol Phone": null,
    'Org': null,
    'Vol Desc': null,
    "Submit Email": null,
    "FWD": null,
    "#PPL": null,
    "Address": null,
    "Location Description": null,
    "Physical Description": null,
    "Needs Description": null,
};

// map column header names in the spreadsheet to hop object attributes - these could be the same, but meh lazy
var col_map = {
    "LA-HOP ID": 'id',
    'Name': 'name',
    "Last Seen": 'last_seen',
    "Submit Date": 'submit',
    'Vol': 'vol',
    "Vol Phone": 'phone',
    "Org": 'org',
    'Vol Desc': 'vol_desc',
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

var fwdFromIDX = null;
function openFwdHeader(fromLine, str) {
  fwdFromIDX = str.indexOf(fromLine);
  return fromLine;
}
function fwdID(subjectLine) {
  return parseInt(subjectLine);
}

function dateIfFwd(dateLine, str) {
  if (Math.abs(fwdFromIDX - str.indexOf(dateLine)) < 1000) { // we need to be close-ish to the from forward line
    dateLine = dateLine.replace(' at ', ' ');
    return new Date(dateLine);
  }
  return null;
}

function fwdType(typeStr) {
  if (typeStr.toLowerCase() === 'new') {
    return NotificationType.NEW;
  } else if (typeStr.toLowerCase() === 'update') {
    return NotificationType.UPDATE;
  } else if (typeStr.toLowerCase() === 'received') {
    return NotificationType.RECEIVED;
  }
  return NotificationType.UNKNOWN;
}

function toDate(dateStr) {
  return new Date(dateStr);
}

// component regexes
// Can be of the format attribute name: [ regex1, regex2 ] and/or attribute name: [ [regex, validate func], ... ]
// if using a validation function, it must return null if validation failed and the final value to use if validation was successful
var attr_re = {
    'addr': [[/>Address: <\/strong>([^<>]*)<\/p>/, validateTextField]],
    'location': [[/>Description of location: <\/strong>([^<>]*)<\/p>/, validateTextField]],
    'last_seen': [[/>Date last seen: (?:<(?:\/)*(?:[^<>]*)>)*([ -,.:\/A-Za-z0-9]*)<\/p>/, toDate]],
    'num': [/>Number of people: <\/strong>(\d+)<\/p>/],
    'name': [[/>Name of person\/people requiring outreach: <\/strong>([^<>]*)<\/p>/, validateTextField]],
    'desc': [[/>Physical description of person\/people: <\/strong>([^<>]*)<\/p>/, validateTextField]],
    'needs': [[/>Description of person\/people(?:(?:')|(?:&#39;))s needs: <\/strong>([^<>]*)<\/p>/, validateTextField]],
    'vol': [[/>Your name: <\/strong>([^<>]*)<\/p>/, validateTextField]],
    'org': [[/>Company\/organization: <\/strong>([^<>]*)<\/p>/, validateTextField]],
    'vol_desc': [[/>How would you describe yourself\? <\/strong>([^<>]*)<\/p>/, validateTextField]],
    'email': [
      [/>Email: <\/strong>([^<>]*)<\/p>/, validateEmail],
      [/>Email: <\/strong><a.*>(.*)<\/a><\/p>/, validateEmail]
    ],
    'phone': [/>Phone: <\/strong>([^<>]*)<\/p>/],
    'fwd_frm': [[/From:.*(donotreply@lahsa.org).*/, openFwdHeader]],
    'fwd_orig_date': [[/Date:[ ]*(?:<(?:\/)*(?:[^<>]*)>)*([ -,.:\/A-Za-z0-9]*)/, dateIfFwd]],  // the original date of the message from LAHSA if this email was forwarded to us
    'fwd_id': [[/Subject:[ ]*(?:<(?:\/)*(?:[^<>]*)>)*[ ]*Outreach Request \*[a-zA-z]+\* \(Request ID: (.*)\)/, fwdID]],
    'fwd_type': [[/Subject:[ ]*(?:<(?:\/)*(?:[^<>]*)>)*[ ]*Outreach Request \*([a-zA-z]+)\*/, fwdType]]
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

const NotificationType = Object.freeze({"NEW":1, "UPDATE":2, "RECEIVED":3, "UNKNOWN": 4});
const Status = Object.freeze({"UNRESOLVED":1, "FAILED":2, "SUCCESS":3, "DISMISSED": 4, 'UNCATEGORIZED': 5});
var StatusReverse = {};
for (const [status, code] of Object.entries(Status)) {
  StatusReverse[code] = status;
}

function getGlobalMeta() {
  /**
   * Fetch global parameters attached to the sheet we use to track how much if anything needs to be
   * run or re-run. This includes, the last time the sheet was synchronized to the gmail account and
   * the version number of this script that was last used to synchronize.
   */
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

    console.info(`Forward Table:`);
    for (const [phone, email] of Object.entries(fwd_table)) {
      console.info(`\t${phone} -> ${email}`);
    }
}

// Main function, the one that you must select before run
function logHOPs() {
  /**
   * Main logic tree. Read through all emails in chunks of 500 at a time matching the HOP query. This will return some false
   * positives because the native gmail filtering capability is not very robust. We reject the false positives if we are unable
   * to parse out their attributes as HOPs.
   *
   * General logic is thus:
   *  1) Fetch the HOP sheet, do some (re)initialization if needed and grab information from the last synchronization off of it.
   *  2) Work through emails matching HOP search 500 at a time:
   *      a. Stop searching when we see an email older than our last synchronization time - unless this is a new version of the script
   *      b. Try to parse the email as a HOP - skip it if we fail, its not a HOP
   *      c. Add (or Update) parsed HOP object to our registry in local mem
   *      d. Forward email if:
   *        - we have a registered forward address for the phone #
   *        - the message itself was not forwarded but came straight to our HOP reception account
   *        - the spreadsheet has been previously synced (don't want lots of errant forwards initially)
   *        - the message is newer than the last synchronization time
   *      e. write HOP list to spreadsheet
   */

    console.log(`Searching for: "${HOP_QUERY}"`);

    getSheet();
    getGlobalMeta();
    buildForwardTable();
    let newestTime = null;
    let start = 0;
    let max = 500;

    let threads = GmailApp.search(HOP_QUERY, start, max);
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
                /*let attachments = msgs[j].getAttachments({includeAttachments: false});
                if (attachments.length > 0) {
                  for (const attachment of attachments) {
                    console.log(attachment.getName());
                    console.log(attachment.getDataAsString());
                    return;
                    let hopAttrs = processMessage(attachment.getName(), attachment.getDataAsString(), msgs[j].getDate());
                  }
                } else {*/
                handleHOP(processMessage(msgs[j].getSubject(), msgs[j].getBody(), msgs[j].getDate()), msgs[j], thread);
                //}
            }
            if (newestTime === null) { break; }  // no messages occurred after LAST_SYNC
        }
        if (newestTime === null) { break; }
        start = start + max;
        threads = GmailApp.search(HOP_QUERY, start, max);
    }

    writeHOPs(hops);
}

function handleHOP(hopAttrs, msg, thread) {
  if (hopAttrs == null) {
    return;
  }
  console.info(`Handling: ${msg.getSubject()}`);
  thread.addLabel(HOP_LABEL);
  hopAttrs.fwd = !accountEmails.includes(hopAttrs.email);
  if (hopAttrs.fwd && hopAttrs.fwd_orig_date) {
    if (hopAttrs.fwd_type === NotificationType.NEW) {
      hopAttrs.submit = hopAttrs.fwd_orig_date;
    }
    hopAttrs.last_update = hopAttrs.fwd_orig_date;
  }
  if (!hops.hasOwnProperty(hopAttrs.id)) {
      hops[hopAttrs.id] = hopAttrs;
  } else {
      for (const attr in hopAttrs) {
          if (attr === 'last_update') {
              if (hopAttrs[attr] > hops[hopAttrs.id][attr]) {
                  hops[hopAttrs.id][attr] = hopAttrs[attr];
                  hops[hopAttrs.id]['status'] = hopAttrs['status'];
              }
          }
          else if (attr === 'status') { continue; }
          else if (hopAttrs[attr] !== null) {
              hops[hopAttrs.id][attr] = hopAttrs[attr];
          }
      }
  }

  // forward message if we should
  if (!hopAttrs.fwd && LAST_SYNC && msg.getDate() > LAST_SYNC) {
    if (fwd_table.hasOwnProperty(hopAttrs.phone)) {
      console.log(`Forwarding ${hopAttrs.id} -> ${fwd_table[hopAttrs.phone]}`);
      msg.forward(fwd_table[hopAttrs.phone]);
    }
  }
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
      let mtch = line.match(rex);
      if (mtch != null && mtch.length >= 2) {
          if (validate) {
            parts[attr] = validate(mtch[1], line);
          } else {
            parts[attr] = mtch[1];
          }
          if (parts[attr] == null) continue;

          return true;
      }
    }
    return false;
}

function processSubject(subj, receivedDate) {
  let subAttrs = {};
  let sub = SUBJ_RE.exec(subj);
  if (sub === null || sub.length < 2) {
      console.log('Unable to interpret subject line!');
      return null;
  }
  if (subj.includes('New')) {
      subAttrs.type = NotificationType.NEW;
      subAttrs.submit = receivedDate;
  } else if (subj.includes('Update')) {
      subAttrs.type = NotificationType.UPDATE;
  } else if (subj.includes('received')) {
      subAttrs.type = NotificationType.RECEIVED;
  }

  subAttrs.id = parseInt(sub[1]);
  return subAttrs;
}

function processMessage(subject, message, receivedDate) {
    fwdFromIDX = null;
    let parts = {
      'last_update': receivedDate, // this won't do for forwards!!
      'type': NotificationType.UNKNOWN,
      'status': Status.UNRESOLVED,
      'submit': null
    };
    for (const key in attr_re) {
      parts[key] = null;
    }
    let subjAttrs = processSubject(subject, receivedDate);
    if (subjAttrs == null) {
      return null;
    }
    for (const [key, val] of Object.entries(subjAttrs)) {
      parts[key] = val;
    }

    for (const [attr, regexes] of Object.entries(attr_re)) {
        if (parts[attr] != null) continue;
        setAttribute(regexes, message, parts, attr);
    }
    if (parts.type !== NotificationType.NEW && parts.status === Status.UNRESOLVED) {
        if (message.includes('unable to make contact') || message.includes('unable to locate the individual')) {
            parts.status = Status.FAILED;
        } else if (message.includes('made contact with the individual')) {
            parts.status = Status.SUCCESS;
        } else if (message.includes('already serving the area listed')) {
            parts.status = Status.DISMISSED;
        }
    }
    if (parts.type !== NotificationType.NEW && parts.status === Status.UNRESOLVED) {
      parts.status = Status.UNCATEGORIZED;
    }
    if (parts.time !== null && typeof parts.time === 'string') {
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
  // convert status #
  row[columns['Status']-1] = StatusReverse[row[columns['Status']-1]];
  return row;
}

function needsUpdates(oldestNewHop) {
  /**
   * Determine if some new hops need to be inserted between existing rows on the spreadsheet. This could happen if people forward HOPs later.
   */
  return oldestNewHop < new Date(HOP_SHEET.getRange(`${columnToLetter(columns['Last Seen'])}2:${columnToLetter(columns['Last Seen'])}2`).getValue());
}

function doRandomInserts(rows) {
  /**
   * Insert new out of order HOPs where they belong in the existing list and update pre-existing HOPs.
   *
   * Returns a list of all new HOPs that were not inserted.
   */
  if (rows.length > 0) {
    var mark = new Date(HOP_SHEET.getRange(`${columnToLetter(columns['Last Seen'])}2:${columnToLetter(columns['Last Seen'])}2`).getValue());
    var randomInserts = [];
    var rows2 = [];
    for (const hop of rows.reverse()) {
      if (hop[columns['Last Seen']-1] >= mark) {
        rows2.unshift(hop);
      } else {
        randomInserts.unshift(hop);
      }
    }
    rows = rows2;
    // this is an O(n) operation but thats fine for now. On a big sheet this might get annoying, HOPs are submitted at a rate of about 50/day
    if (randomInserts.length > 0) {
      var times = HOP_SHEET.getRange(`${columnToLetter(columns['Last Seen'])}:${columnToLetter(columns['Last Seen'])}`).getValues().slice(1);
      let randomIdx = 0;
      let timeRow = 2;
      let lastTimeIDX = 1;
      for (let time of times) {
        time = new Date(time);
        if (!isNaN(time.getTime())) {
          while (randomInserts.length > randomIdx && randomInserts[randomIdx][columns['Last Seen']-1] >= time) {
            let insertRange = HOP_SHEET.getRange(`A${timeRow}:${columnToLetter(randomInserts[randomIdx].length)}${timeRow}`);
            insertRange.insertCells(SpreadsheetApp.Dimension.ROWS);
            insertRange.setValues([randomInserts[randomIdx]]);
            randomIdx++;
            timeRow++;
          }
          lastTimeIDX = timeRow;
          if (randomInserts.length <= randomIdx) break;
        }
        timeRow++;
      }
      if (randomIdx < randomInserts.length) {
        let toInsert = randomInserts.slice(randomIdx);
        let insertRange = HOP_SHEET.getRange(`A${lastTimeIDX+1}:${columnToLetter(randomInserts[randomIdx].length)}${lastTimeIDX+(randomInserts.length)}`);
        insertRange.insertCells(SpreadsheetApp.Dimension.ROWS);
        insertRange.setValues(toInsert);
      }
    }
  }
  return rows;
}

function incidentOrderDescending(row1, row2) {
  if (row1[columns['Last Seen']-1] < row2[columns['Last Seen']-1]) {
    return 1;
  }
  else if (row1[columns['Last Seen']-1] > row2[columns['Last Seen']-1]) {
    return -1;
  }
  return 0;
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

  var rows = [];
  for (const [hop_id, hop] of Object.entries(hops)) {
    rows.push(toRow(hop));
  }
  rows.sort(incidentOrderDescending);

  if (rows.length > 0) {
    // first pass, search for HOPs by ID in sheet and update the found ones
    if (needsUpdates(rows[rows.length-1][columns['Last Seen']-1])) {
      var remaining = [];
      var existing = {};
      var hidx = 1;
      for (const hop of HOP_SHEET.getRange(`${columnToLetter(columns['LA-HOP ID'])}:${columnToLetter(columns['LA-HOP ID'])}`).getValues()) {
        if (parseInt(hop[0])) {
          existing[hop[0]] = hidx;
        }
        hidx++;
      }
      for (const hop of rows) {
        var hid = hop[columns['LA-HOP ID']-1].toString();
        if (existing.hasOwnProperty(hid)) {
          var range = HOP_SHEET.getRange(`${existing[hid]}:${existing[hid]}`)
          var current = range.getValues()[0];
          if (hop[columns['Last Update']-1] >= new Date(current[columns['Last Update']-1])) {
            var vidx = 0;
            for (const val of hop) {
              if (val != null) {
                current[vidx] = val;
              }
              vidx++;
            }
            range.setValues([current]);
          }
        } else {
          remaining.push(hop);
        }
      }
      // add any missing HOPS between existing rows in the table
      rows = doRandomInserts(remaining);
    }

    // add all new HOPs to the top of the list
    if (rows.length > 0) {
      var insertRange = HOP_SHEET.getRange(`A2:${columnToLetter(HOP_SHEET.getLastColumn())}${rows.length+1}`);
      insertRange.insertCells(SpreadsheetApp.Dimension.ROWS);
      insertRange.setValues(rows);
    }
  }
}
