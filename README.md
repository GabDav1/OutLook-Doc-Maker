# OutLook-Doc-Maker
Word-VBA tool that compiles mails(from Outlook) into a document. (given one or more key-terms).

  The macro crawls through the MAPI namespace (all inboxes, folders and subfolders, everything) and "picks up" 
everything that matches the search criteria. The purpose is to generate documentation based on mail backlogs.

  The search term looks in the Subject field of the mail. The first field recovers only the "thread seeds"
(replies to the thread e.g. mails containing "RE:" are ignored). The second field recovers the full threads 
(replies e.g. mails with "RE:" are also included).

  Sometimes, the macro might generate an error upon opening the first mail in the list.
In that case, just click on "OK for everything" and the script will resume running.

![With RE:](https://user-images.githubusercontent.com/17041548/68474143-57eaa200-022d-11ea-8ec1-9d7eff38c5ce.jpg)
