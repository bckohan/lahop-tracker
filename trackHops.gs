/**
 * The following two fields need to be updated.
 */
var ACCOUNT = Session.getActiveUser().getEmail();
var HOP_SHEET_ID = '1vfzkKmBUYTy5fn-L4maZXEfN8hlAX1kQDYsvgblWxQA'; // the google spreadsheet ID
var ORG_DOMAIN = 'selahnhc.org'; // The domain of your org, i.e. the domain.org part of email@domain.org


// this is the search string as you would type it into Gmail search
var HOP_QUERY = 'from: donotreply@lahsa.org subject: Outreach Request';
var SUBJ_RE = /\(Request ID: (\d+)\)/;
var SHEET_NAME = ACCOUNT.substr(0, ACCOUNT.indexOf('@')).replace('selah', '');
var LAST_SYNC = null;
var CURRENT_SYNC = null;

var columns = {
    "LA-HOP ID": 1,
    "Origin Date": 2,
    "Vol Phone": 3,
    "#PPL": 4,
    "Address": 5,
    "Location Description": 6,
    "Physical Description": 7,
    "Needs Description": 8,
    "Status": 9,
    "Resolve Date": 10
};

for (var a in GmailApp.getAliases()) {
    let alias = GmailApp.getAliases()[a];
    if (alias.endsWith(ORG_DOMAIN)) {
        SHEET_NAME = alias.substr(0, addy.indexOf('@'));
    }
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

//DATE_TIME_RE = /(\d{1,2})\/(\d{1,2})\/(\d{4}) (\d{1,2}):(\d{1,2})(.*) (A|P)M/;

const NotificationType = Object.freeze({"NEW":1, "UPDATE":2, "RECEIVED":3, "UNKNOWN": 4});
const Status = Object.freeze({"UNRESOLVED":1, "NO_CONTACT":2, "CONTACT":3, "REDUNDANT": 4});

function getGlobalMeta(sheet, key) {
    //todo this probably isn't write - the meta API is very obtuse
    arr = sheet.createDeveloperMetadataFinder().withKey(key).find();
    if (arr && arr.length > 0) {
        return arr[idx];
    }
    return null;
}

// Main function, the one that you must select before run
function findHOPs() {

    console.log(`Searching for: "${HOP_QUERY}"`);

    let sheet = getSheet();
    LAST_SYNC = getGlobalMeta(sheet, 'sync_time');
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
                if (LAST_SYNC && msgs[j].getDate() <= LAST_SYNC) { continue; }
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

    console.info(SHEET_NAME);
    console.info(hops);
    console.log(lastProcessed);
    //writeHOPs(sheet, hops);
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
    var sheet = hop_ss.getSheetByName(SHEET_NAME);
    if (sheet === null) {
        sheet = hop_ss.insertSheet(SHEET_NAME);
        addHeaders(sheet);
    }
    return sheet;
}

function addHeaders(sheet) {
    sheet.setFrozenRows(1);
    var range = sheet.getRange("A1:J1");
    range.setValues([Object.keys(columns)]);
    for (const col in columns) {
        sheet.getRange(1, columns[col]).addDeveloperMetadata('col_name', col);
    }
    var range = sheet.getRange(1, 1);
}

function resolveColNums() {

}

function writeHOPs(sheet, hops) {
    // TODO
    sheet.addDeveloperMetadata('sync_time', CURRENT_SYNC);

    // first pass, search for HOPs by ID in sheet and update the found ones

    // second pass, add remaining HOPs to top of sheet in descending date order of creation
}
