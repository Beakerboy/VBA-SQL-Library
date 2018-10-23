VBA SQL Library
=====================

### Object-Based Database Interaction for VBA

Easily create SQL queries and execute them on a database

Features
--------
 * [Database](#database)
 * [Insert](#insert)
 * [Select](#select)
 * [Update](#update)
 * [Helper Functions](#helper-functions)
 
 Setup
-----

Import the files into a spreadsheet using Microsoft Visual Basic for Applications. Edit the database class to include the DSN and database type for your database. 
 
 Security
-----
This Library currently does not use prepared statements or string sanitation to prevent SQL Injections.
However, it does require users to authenticate to the database to perform any queries.

 
 Usage
-----

### Database
Create a new database connection:
```vb
'Initialize the object and assign a table name
Set MyDatabase = New SQLDatabase
MyDatabase.DBType = "mssql"
MyDatabase.DSN = "foodb"

'Open UserForm
Login.Show

'After Button is pressed assign values
MyDatabase.UserName = Login.Username
MyDatabase.Password = Login.Password
Unload Login

```

### Insert
The Insert object can create both INSERT VALUES and INSERT SELECT statements.

```sql
INSERT INTO users (username, first_name, password) VALUES ('admin', 'Alice', 'secret');
```

```vb
'Initialize the object and assign a table name
Set MyInsert = new Insert
MyInsert.table = "users"

'Set The Fields
MyInsert.Fields = Array("username", "first_name", "password")

'Set the Values
MyInsert.Values = Array(str("admin"), str("Alice"), str("secret"))

'Execute the query
MyInsert.Insert
```

```sql
INSERT INTO bank_account (account_number, open_date, user_id)
    SELECT (10, 570000051, user_id) FROM users WHERE username = 'admin';
````

```vb
'Initialize the object and assign a table name
Set MyInsert = new Insert
MyInsert.table = "bank_account"

'Set The Fields
MyInsert.Fields = Array("account_number", "open_date", "user_id")

'Create the SELECT Statement
Set SQL = New SQLSelect
Sql.Fields = Array(10, 5770000051, "user_id")
Sql.Table = "users"
Sql.addWhere "username", "=", "admin"
InSQL.setSelect = Sql

'Execute the query
MyInsert.Insert
```


### Select


### Update
Not Yet Implemented

### HelperFunctions
The library includes a handful of helper functions. 
* Date/Time manipulation, toIso() and toUnix().
* String encapsulation str() to add single quotes around strings.
