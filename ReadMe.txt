HTML Messages Encoder V1.0 ReadMe:
==================================

This is a simple application that converts RichText to Encrypted HTML files. Encrypted means the source RichText will not appear until the user enters the Password that you've chosen to encrypt the source with. What's really cool, it does NOT need any special software to decrypt the encoded file. Any Internet Explorer will handle that! 
 
It is affective for sending confidential messages via email as an attached HTML file. It can also convert RichText to regular HTML files. 
Please note that the decryption process involves using VBScripts, only supported by Internet Explorer 4 or above.

To convert RichText to Encrypted HTML file...
	-Run HTML Messages Encoder, click Open Message, and select your RTF (RichText format) file, or just copy & paste it.
	-In the Opening Message field, type the message that you'd like to display at first when your result HTML file is opened.
	-Click "Export to Encrypted HTML", enter the Password that you wish to have the source encrypted with.

To decrypt the Encrypted HTML file...
	-Open the file with IE 4 or above.
	-The Opening Message appears.
	-Enter the Password and click Decrypt Message
	-Wait for sometime... The source RichText appears.
	*If IE requested interrupting the script, just don't allow it (the decryption process may take few minutes if the source was rather large).

-The encryption cipher is based on the PC1 encryption alogrithm written by Alexander Pukall (alexandermail@hotmail.com). I tried to replace it with faster and more secure ciphers (such as RC4), but I figured out that the PC1 would be more suitable, coz it generates no line feed characters.
-If you encountered Overflow errors when running the application within the VB environment, choose make EXE from the File menu, click on Options, click the Compile tab, make sure it is set to Compile to Native code, click the "Advanced Optimisation" button, and activate "Remove Integer Overflow Checks". Compile to an EXE file, and run the executable. It should NOT generate errors this time.
-If someone managed to break your encrypted message, just don't blame me, I'm NOT responsible. 
-This program's output has been only tested on Win98 platforms with IE 5.5.