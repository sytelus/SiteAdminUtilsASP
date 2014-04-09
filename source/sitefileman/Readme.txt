Programmed By webmaster@pitic.net

This Code is made only for Education and Demonstration purposes. 
Is not intended for Commercial Use and It's still on develop.

You need a folder name 'admin' in the root of your web server with 
permission to run scripts and exe, in that folder must have to be 
5 files (admin.asp, edit.asp, index.asp, manager.asp and upload.exe)
and 1 folder name 'Util'. In 'Util' folder must have to be Database
name 'util.mdb' and some other images and utilities files. The entry
to the program is the 'Index.asp' file. In that file you have to type
your password but first you have to setup accounts, make a test account
and then tray to login. To setup account you have to login with the 
administrator password: Username: admin Password: adm .

Note: For all Users the program makes a folder in the root.

The program Stores Information in the Session Object of the IIS.
The Variables used are:

Session("Nombre") = Name of the user.
Session("Inter") = Path of the admin scripts (in a local web server is something like c:\inetpub\wwwroot\admin)
Session("Direc") = The root of the User.
Session("DirDB") = The Path of the DataBase.
Session("Permiso") = The permission asigned for the user.
	The "Permiso" varable can store "Si" = yes or "No" = no.
