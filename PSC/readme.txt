! Important !
If youre using this program with MySQL, you have to recreate the database due to the updated structure.
! /Important !

This program can use MySQL (with InnoDB) or the MS SQLServer.
If you have a different driver for connecting (we used "MySQL ODBC 5.1 Driver" and "SQL Server Native Client 10.0"),
specify it in "Mod2.bas" in the "OpenDB" function and in "FrmConnDB" in "ConnectDB" function.

The "db.sql" file in Data-Directory contains the db structure and some sample data for using MySQL.
Create a database, grant an user privileges to it and then load this file into the database.
The Data-directory is also used to store the connection info to the Database Server.

If youre using SQLServer, restore the backup file "dbsv_psc_sicher.bak" into a database
and assign a sql user the database user "dbsv" from the database. Check that the user has the "db_owner" role assigned.

There are 2 main directorys, one for the german and one for the english version.
