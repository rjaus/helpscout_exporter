# helpscout_exporter
Helpscout mailboxes exporter (csv list + html files with mails)

This version has been made 20x more efficient, by requesting the threads embedded in the conversation request.

# install
Create a .env file

`touch .env`

Add your client id & secret
```
HELPSCOUT_CLIENT_ID=abc
HELPSCOUT_CLIENT_SECRET=def
```

# usage
List mailboxes

`python3 export.py -l`

Download conversation from a mailbox

`python3 export.py -e 123123`

Set the start and end pages from which to download conversations.  Useful for resuming an existing export that failed.

`python3 export.py -e 123123 --start 2 --end 29`


### Exporting Attachments
There are two methods for exporting attchments

1. Export all attachments for an inbox
This is the suggested method.  It creates an attachments folder, and a sub folder for each email address, in which the attachments from that email address are saved.
This method cycles through all conversations, and the associated threads.  It saves each attachment for the thread.
`python3 export.py -e 123123 --attachments-only`

2. Export attachment while exporting the mailbox conversations
Attachments can be downloaded using the --attachments flag.  I'd suggest re-jigging how the filenames are saved, and adding some logic to avoid re-downloading duplicate attachments.  The first method is more efficient.
`python3 export.py -e 123123 --attachments`
