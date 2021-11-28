# lahop-tracker

A utility that monitors a gmail address for LAHOP request messages and logs the status
of those HOPs to a Google spreadsheet. This script is in active use and works well, but it does
take some know-how to get setup and running. If you or your organization would like to use
this utility please [fill out this form.](https://docs.google.com/forms/d/1a6rOii5MONQSlbHpghopkAFF-Wb70R29dj9kAUEMkt4)
I will reach out and assist with any setup issues. Even if you don't require help setting it up, fill
out the form anyway to get notified of updates.

## Suggested Operation
If using this tracker as part of an organization where multiple people will be submitting HOPs, its
suggested to devote one gmail address to be the submitter email. All hops should be submitted using
that gmail address and this script will run only on that address. HOP submitters should then use their
own phone number when they submit the HOP. The email forwarding feature of this script lets users
map their phone number to their email address so any communications from LAHSA regarding that user's
HOPs will be automatically forwarded to them when this script runs.

## Setup

The steps to setup this script are broadly:

- Login to your gmail/google account.
- Navigate to https://script.google.com/
- Create a new project called LA HOP Tracker (or whatever)
- Copy and paste the contents of trackHops.gs into code.gs in the script editor
- Change this line: var HOP_SHEET_ID = ''; to var HOP_SHEET_ID = 'ID of your google spreadsheet';
    - The google sheet ID is the long string of random numbers and letters in the spreadsheet's url
- If you want to use the forwarding feature, create a sheet called "Forwards" of the following format in the
same google spreadsheet that logs the HOPs:
   *   0 | Phone      | Email
   *   1 | 5558183333 | myemail@example.com
   *   ...
   *   n | 5553234444 | someoneelse@example.com
   *
- Set the script up to run daily, by navigating to "Triggers" (clock icon) and creating a new time based trigger.

## Capturing old HOPs

This script can log individually forwarded HOPs. If you want to read in old HOPs you will have
to individually forward them to the address running this script OR run this script on the address(es)
holding the old HOPs. Unfortunately using gmail's bulk forward as attachment feature does not work
currently for some annoying technical reasons. There may be some add-ons available that provide a
bulk forwarding feature that works as intended, but I haven't tried them.

In gmail the best way to search for old HOPs submitted by that address is: `from: donotreply@lahsa.org subject: outreach request`
