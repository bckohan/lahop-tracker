// this is the search string as you would type it into Gmail search
var HOP_QUERY = 'from: donotreply@lahsa.org subject: Outreach Request'
var SUBJ_RE = /\(Request ID: (\d+)\)/;

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

// Main function, the one that you must select before run
function findHOPs() {

    console.log(`Searching for: "${HOP_QUERY}"`);
    var start = 0;
    var max = 500;

    var last_hop = getMostRecentRecordedHOP();

    var threads = GmailApp.search(HOP_QUERY, start, max);

    var hops = {};
    while (threads.length>0) {
      for (var i in threads) {
          var thread=threads[i];
          var msgs = threads[i].getMessages();
          for (var j in msgs) {
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
      }

      if (threads.length == max){
          console.log("Reading next page...");
      } else {
          console.log("Last page readed 🏁");
      }
      start = start + max;
      threads = GmailApp.search(HOP_QUERY, start, max);
    }

    console.info(hops);
}

function getMostRecentRecordedHOP() {

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
        // timezone might be wrong here, but thats ok we only use this for odering, TZ not displayed in sheet
        parts.time = new Date(parts.time);
    }
    return parts;
}

// Add contents to sheet
//function appendData(line, array2d) {
//  var sheet = SpreadsheetApp.getActiveSheet();
//  sheet.getRange(line, 1, array2d.length, array2d[0].length).setValues(array2d);
//}
