VBA SQL Library
=====================

### Object-Based Database Interaction for VBA

Easily create SQL queries and execute them on a database

Features
--------
 * [Database](#database)
 * [Insert](#insert)
 * [Select](#select)
 * [Helper Functions](#helper-functions)
 
 Setup
-----

 Copy and paste the code from each file into Excel VBA modules. Edit the Database object to
 include your database authentication details.
 
 Security
-----
This Library currently does not use prepared statments or string sanitation to prevent SQL Injections.
However, it does require users to authenticate to the database to perform any queries.

 
 Usage
-----

### Database
Add the proper values to the class_initialize function to connect to your database.
This object is only used by the SQL objects.

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

### HelperFunctions
The library includes a handful of helper functions. 
* Date/Time manipulation, toIso() and toUnix().
* String encapsulation str() to add single quotes around strings.
