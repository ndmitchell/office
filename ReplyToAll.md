# Reply To All

Some companies lock down the use of Outlook by disabling the Reply To All button, making it harder to manage email. This document describes an associated Office VBA module, which restores the Reply To All button. The associated VBA also includes features to edit a received email (another feature sometimes disabled) and to remove embedded images from an email.

## Enabling Reply To All

The following steps describe how to recover Reply To All functionality. These steps has been tested on Outlook 2003 and 2007, but should also work for other versions of Outlook.

### Step 1: Reduce your Outlook security settings</h3>

In Outlook, go to `Tools`, `Macro`, `Security...`. In the dialog box, select Medium or Low security.

### Step 2: Download the Macro file

Download the macro file from [here](https://raw.githubusercontent.com/ndmitchell/office/master/ReplyToAll.bas), by right clicking and saving the module. Save it on your Desktop.

### Step 3: Import the Macro file

In Outlook, go to `Tools`, `Macro`, `Visual Basic Editor`.

In the Microsoft Visual Basic editor go to `File`, `Import File...`. Select the macro file you saved previously on your Destop named `ReplyToAll.bas`.

Close the Microsoft Visual Basic editor.

### Step 4: Add the buttons

In Outlook, go to `Tools`, `Macro`, `Macros...`. Select the Macro `ReplyToAll_AddButtons` and hit `Run`.

### Using Reply To All

Now, in Outlook, click on a `Reply To All` button in the main window. The first time hitting `Reply To All` after starting Outlook you may get a security warning - if so just click `Enable Macros`.

## Additional Features

The associated Macro file contains two additional macros:

* `EditMessage_Run` - edit the currently selected message, useful for when companies have also disabled Edit Message.
* `DeleteImages_Run` - delete all inline images from a message, which sometimes substantially reduces message size.

To use these macros, first open a received email message. If you are using Office 2007, add them to the Quick Access Toolbar at the top of the screen. If you are using Office 2003, add them as custom toolbar buttons.

## Plea to Outlook Administrators

Please do not disable essential email functionality. With the workarounds described above, the attempt is futile, but remains deeply inconvenient. Consider the situation where Alice sends an email to Bob, Charlie and Dave asking for some financial details of Mega Corp. Bob has the details on a post-it note on his desk and quickly replies to everyone with the information. But in a world without Reply To All...

* Bob replies just to Alice. Charlie, being the helpful soul he is, decides to search through the filing cabinet for the information. 15 minutes later Charlie finds the details and emails Alice, only to be told "Thanks, Bob answered this 15 minutes ago". Charlie  realises he just wasted 15 minutes of his life, and goes to get a cookie to make himself feel better.
* Bob replies just to Alice. At the end of the week Dave is reviewing his emails and realises that Alice still hasn't got a reply, but before he potentially wastes 15 minutes, he drops Alice an email - "Do you still want those details?". Alice replies "No". Dave concludes that he's a clever bunny for not going to the filing cabinet, and decides a one week latency when replying to Alice is just common sense.
* Bob replies, but realising the potential cost of replying just to Alice, also replies to Charlie and Dave. Bob spent a minute retyping the recipient list, and wonders "Why can't Outlook have a button that does this for me?".
* Bob replies, also including Charles and Dave. Woops! That should have been Charlie, not Charles. As a best case scenario, Charles gets annoyed with a useless email, but at worst Charles brings down Mega Corp with the sensitive information gleaned from the email. Charlie still ends up going to get a cookie.

Removing Reply To All increases the volume of email required, and increases the risk of email accidents. I've heard only two arguments against Reply To All, both of which are wrong:

* When Bob replies to everyone, the financial information about Mega Corp gets sent to Charlie, but what if Charlie shouldn't have access to that information? Of course, if Charlie shouldn't have access to that information then Alice was at fault for sending the request to Charlie in the first place. Bob should also check when sending sensitive information, but most emails are not sensitive (it is a poor default), checking the senders should not require retyping the senders (it is a poor user interface) and by default mailing lists are not fully expanded (so he won't be able to tell the full list of senders anyway).
* When Zorg, the owner of Mega Corp, sends a Christmas email to all his employees, what if one of them hits Reply To All? That wastes a lot of company time, and should be prevented - but not on the client side. It is possible to restrict the number of recipients to an email, Zorg could send out the email with everyone in the BCC field, or IT could set up a mailing list which does not permit replies.
