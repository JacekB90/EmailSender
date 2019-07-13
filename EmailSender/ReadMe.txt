Hello Folks,

Program works with IBM Notes and you must have been logged.

Program was built in C# and VB that is why it contains 2 .exe files: FastGoupEmail(C#) and EmailNotesSend(VB). 
You need to have FastGoupEmail(C#) and EmailNotesSend(VB) in the same localization to use program.
To start program use FastGoupEmail.exe.

This program send the same message with individual attachment.

All you need to do is prepare email list and attachments.

How to prepare email list:
Email list should be in .txt formt. Structure of email list must contains: name of attachment, comma, email address.
For example:
123,name@name.pl
258,name2@name.pl
963,name3@name.pl

How to prepare attachments:
Name of attachments must be the same as name used into email list.

You can also use button 'Split Excel File' to prepare attachments. 
Data in excel file must be input in first sheet and first column contains key(name of attachment).
For example:
Code	Name	Bonus
123	X	1.000$
258	Y	900$

Important information: program matches email addresses to attachments, so if attachments not exists then email won't be send.
