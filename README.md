VBA SQL Library
=====================

### Object-Based Database Interaction for VBA

Easily create SQL queries and execute them on a database

Features
--------
 * [Login Form](#login-form)
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

### Login Form
This form can be displayed to ask for the database credentials. This avoids having to hard-code username and passwords in the scrips.
```vb
'Open UserForm
Login.Show

'After Button is pressed assign values
MyDatabase.UserName = Login.Username
MyDatabase.Password = Login.Password
Unload Login
```

### Database
Create a new database connection:
```vb
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

Several different types of database execution can occur:
 * Execute(SQL) - Execute a statement (Insert or Update)
 * InsertGetNewID(SQLInsert) - Insert a record, and return the new primary key
 * Execute(SQLSelect, column) - Execute a statement and return a single value
 * Execute(SQLSelect) Execute a statement and return an array of values

### Insert
The Insert object can create both INSERT VALUES and INSERT SELECT statements.

```sql
INSERT INTO users (username, first_name, password) VALUES ('admin', 'Alice', 'secret');
```

```vb
'Initialize the object and assign a table name
Set MyInsert = new SQLInsert
MyInsert.table = "users"

'Set The Fields
MyInsert.Fields = Array("username", "first_name", "password")

'Set the Values
MyInsert.Values = Array(str("admin"), str("Alice"), str("secret"))

'Execute the query
MyDatabase.Execute MyInsert 
```

```sql
INSERT INTO bank_account (account_number, open_date, user_id)
    SELECT (10, 570000051, user_id) FROM users WHERE username = 'admin';
````

```vb
'Initialize the object and assign a table name
Set MyInsert = new SQLInsert
MyInsert.table = "bank_account"

'Set The Fields
MyInsert.Fields = Array("account_number", "open_date", "user_id")

'Create the SELECT Statement
Set SQL = New SQLSelect
Sql.Fields = Array(10, 5770000051, "user_id")
Sql.Table = "users"
Sql.addWhere "username", "=", "admin"
MyInsertSQL.setSelect = Sql

'Execute the query, returning the newly created primary Key
ID = MyDatabase.InsertGetNewID(MyInsert)
```


### Select


### Update
Not Yet Implemented

### HelperFunctions
The library includes a handful of helper functions. 
* Date/Time manipulation, toIso() and toUnix().
* String encapsulation str() to add single quotes around strings and escape contained single-quotes
