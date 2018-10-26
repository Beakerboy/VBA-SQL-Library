VBA SQL Library
=====================

### Object-Based Database Interaction for VBA

Easily create SQL queries and execute them on a database

Features
--------
 * [Database](#database)
 * [Login Form](#login-form)
 * [Insert](#insert)
 * [Select](#select)
 * [Update](#update)
 * [Helper Functions](#helper-functions)
 
 Setup
-----

Import the files into a spreadsheet using Microsoft Visual Basic for Applications.
 
 Security
-----
This Library currently does not use prepared statements but it does escape all single quotes. It also provides a login box to prevent hard-coding database autentication details.

 Usage
-----

### Database
Create a new database connection:
```vb
Set MyDatabase = New SQLDatabase
MyDatabase.DBType = "mssql"
MyDatabase.DSN = "foodb"
```
Several different types of database execution can occur:
 * Execute(SQL) - Execute a statement (Insert or Update)
 * InsertGetNewID(SQLInsert) - Insert a record, and return the new primary key
 * Execute(SQLSelect, column) - Execute a statement and return a single value
 * Execute(SQLSelect) Execute a statement and return an array of values
 
The Database can be placed in "Test Mode" as well by setting ```vb MyDatabase.TestMode = True ```. This will print all queries to Message Boxes, and only execute Select queries. This allows a developer to inspect all the queries for accuracy without risking any changes to the database.

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

### Insert
The Insert object can create both INSERT VALUES and INSERT SELECT statements.

#### Example 1 - Insert Values
To produce this SQL Stament:
```sql
INSERT INTO users (username, first_name, password) VALUES ('admin', 'Alice', 'secret');
```

Use the Following in VBA-SQL-Library
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

#### Example 2 - Insert Select
To produce this SQL Stament:
```sql
INSERT INTO bank_account (account_number, open_date, user_id)
    SELECT (10, 570000051, user_id) FROM users WHERE username = 'admin';
````
Use the Following in VBA-SQL-Library
```vb
'Initialize the object and assign a table name
Set MyInsert = new SQLInsert
MyInsert.table = "bank_account"

'Set The Fields
MyInsert.Fields = Array("account_number", "open_date", "user_id")

'Create the SELECT Statement
Set SQL = New SQLSelect

'We don't escape the "user_id" because it is a field name, not a string
Sql.Fields = Array(10, 5770000051, "user_id")
Sql.Table = "users"
Sql.addWhere "username", "=", str("admin")
MyInsertSQL.setSelect = Sql

'Execute the query, returning the newly created primary Key
ID = MyDatabase.InsertGetNewID(MyInsert)
```


### Select
We can execute a select statement and receive the results as a single value, or an array of values:
```sql
SELECT id FROM users WHERE username='admin';
```

```vb
Set MySelect = New SQLSelect
MySelect.Fields = Array("id")
MySelect.Table = "users"

'Need to escape the string
MySelect.AddWhere "username" "=" str("admin") 

ID = MyDatabase.Execute(MySelect, "id")
```


### Update
Not Yet Implemented

### HelperFunctions
The library includes a handful of helper functions. 
* Date/Time manipulation, toIso() and toUnix().
* String encapsulation str() to add single quotes around strings and escape contained single-quotes
