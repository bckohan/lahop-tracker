/**
 * The following field(s) should be configured to your liking.
 */
var HOP_LABEL = GmailApp.createLabel('HOP');       // tag HOP emails with this label in gmail
var FORWARD_TO = 'selahhollywood@gmail.com';
///////////////////////////////////////////////////////////////////////

var HOP_QUERY = 'from: donotreply@lahsa.org subject: Outreach Request';

// Main function, the one that you must select before run
function forwardHOPs() {

    console.info('Please only run this once!');
    console.info(`Searching for: "${HOP_QUERY}"`);

    let start = 0;
    let max = 500;

    let threads = GmailApp.search(HOP_QUERY, start, max);
    while (threads.length>0) {
        for (var i in threads) {
            var thread=threads[i];
            thread.addLabel(HOP_LABEL);
            var msgs = threads[i].getMessages();
            for (const msg of msgs) {
                console.info(`Forwarding ${msg.getSubject()}`);
                msg.forward(FORWARD_TO);
            }
        }
        start = start + max;
        threads = GmailApp.search(HOP_QUERY, start, max);
    }
}
