README
======

The SendJpgEmails.vbs program is a vb script to choose a mail receipient, select a folder with jpg pictures, and then send
the jpg files, **one by one**  to the mail receipient.

The reason to do this vb script is that emails are generally limited to ~ 10M per email, and often digital cameras generate jpgs 
that are quite large (>3M). Thus, only 1 or 2 pictures per email is possible, and it can be tedious generating emails one by one to include
these jpgs.

Second, uploading places did  not or no longer provide full resolution (and memory consumption) for each picture. This is really only 
a problem if  you want to make enlargements of the photo. But if you are going to save the files (maybe in the cloud) then might as well
save in the best resolution possible.

Usage
-----
The SendJpgEmails.vbs is started by clicking on it. ** It assumes you have an outlook account. ** It can be modified to use
a regular MAPI or IPOP server, but does not.

THen the user must respond to the following queries:
- Answer the question to whom the email is going
- Answer the folder selection where the jpg files are located. You can hit cancel to abort.

once these queries have been answered then each jpg file in the folder will be sent individually to the email receipient. For
each jpg file a Message box will query with OK/CANCEL telling the user do you want to send the email. If you hit ok, it will be 
sent, if you hit CANCEL, it will abort the script and stop the sending of jpg files.