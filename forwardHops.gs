/**
 * The following field(s) should be configured to your liking.
 */
var HOP_LABEL = GmailApp.createLabel('HOP');       // tag HOP emails with this label in gmail
var FORWARD_TO = 'TODO Put The Hop tracker email here!';
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
              if (!msg.getFrom().includes('donotreply@lahsa.org')) {
                continue;
              }
              let dt = `${msg.getDate().toDateString()} at ${msg.getDate().toLocaleTimeString('en-us', { timeZoneName: 'short' })}`;
              let fwdHeaderPlain = `---------- Forwarded message ---------\nFrom: LAHSA <donotreply@lahsa.org>\nDate: ${dt}\nSubject: ${msg.getSubject()}\nTo: ${msg.getTo()}\n`;
              let fwdHeaderHtml = `---------- Forwarded message ---------<br/><b>From:</b> LAHSA &lt;<a href="mailto:donotreply@lahsa.org" target="_blank">donotreply@lahsa.org</a>&gt;<br/><b>Date:</b> ${dt}<br/><b>Subject:</b> ${msg.getSubject()}<br/><b>To:</b> ${msg.getTo()}<br/><br/>`;
              let fwd = msg.createDraftReply(fwdHeaderPlain + msg.getPlainBody());
              fwd.update(
                FORWARD_TO,
                `Fwd: ${msg.getSubject()}`,
                fwdHeaderPlain + msg.getPlainBody(),
                {
                  htmlBody: fwdHeaderHtml + msg.getBody()
                }
              );
              console.log(`Forwarding: ${msg.getSubject()}`);
              fwd.send();
            }
        }
        start = start + max;
        threads = GmailApp.search(HOP_QUERY, start, max);
    }
}
