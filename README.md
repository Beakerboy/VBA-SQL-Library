VBA SQL Library
=====================

### Object-Based Database Interaction for VBA

Easily create SQL queries and execute them on a database. For an example on how to use this library with data objects, see the [VBA-Drupal-Library](https://github.com/Beakerboy/VBA-Drupal-Library). For other examples on using the library, refer to the unit tests.

Features
--------
 * [Database](#database)
 * [Login Form](#login-form)
 * [Insert](#insert)
 * [Select](#select)
 * [Update](#update)
 * [Helper Functions](#helper-functions)
 * [Unit Tests](#unit-tests)
 
 Setup
-----

Import the files into a spreadsheet using Microsoft Visual Basic for Applications.
 
 Security
-----
This Library currently does not use prepared statements but it does provide a function to escape all single quotes. It also provides a login box to discourage hard-coding database authentication details. This library should be used within a larger system that provides data integrety checks which ensure SQL injection can not occur. For example, a developer creates an object in VBA with a hard-coded table name to prevent a malicious user from injecting into the MySelect.Table Property. Similarly, they would develop methods to ensure that the MyInsert.Values array has numeric data where it is expected, and escaped string data where it is expected. 

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
 
The Codebase provides a SQLTestRecordset object to test queries without altering any data. If you pass this object to the SQLDatabase, all queries will be printed to a msgBox. Select Queries will be executed, and the results will also be printed to a msgbox.
```vb
Dim MyDatabase As New SQLDatabase
Dim MyTestRecordset As New SQLTestRecordset
MyDatabase.DBType = "mssql"
MyDatabase.DSN = "foodb"
Set MyDatabase.Recordset = MyTestRecordset
```
A Custom Testing Recordset can be created as long as it implements the same Public Properties, Subs, and Functions as the SQLRecordset object

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
The Insert object can create both INSERT VALUES and INSERT SELECT statements. Multiple inserts can be performed in one statement if the values array is 2 Dimensional.

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

#### Example 3 - Insert Multiple Values

To produce this SQL Stament:
```sql
INSERT INTO users (username, first_name, password) VALUES ('admin', 'Alice', 'secret'), ('editor', 'Bob', 'super-secret');
```

Use the Following in VBA-SQL-Library
```vb
'Initialize the object and assign a table name
Set MyInsert = new SQLInsert
MyInsert.table = "users"

'Set The Fields
MyInsert.Fields = Array("username", "first_name", "password")

'Set the Values
Dim Values2D(1, 2) As Variant
Values2D(0) = Array("'admin'", "'Alice'", "'secret'")
Values2D(1) = Array("'editor'","'Bob'", "'super-secret'")
MyInsert.Values = Values2D

'Execute the query
MyDatabase.Execute MyInsert 
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
MySelect.AddWhere "username", "=", str("admin") 

ID = MyDatabase.Execute(MySelect, "id")
```
WHERE clauses can be added and grouped together. The following changes the query to:
```sql
SELECT id FROM users WHERE username='admin' AND id < 10;
```
```vb
MySelect.AddWhere "id", "<", 10, "AND"
```
A SQLWhereGroup can abe be added using SQLSELECT.AddWhereGroup. This is necessary for a where clause like:
```sql
SELECT id FROM users WHERE (a=1 AND b=2) OR (c = 3 AND d = 4)
```
The SQLSelect Object can create Queries with "GROUP BY ... HAVING ..." sections.
```sql
... GROUP BY user_type HAVING age > 1;
```
```vb
MySelect.GroupBy = Array("user_type")
MySelect.AddHaving = "age", ">", "1"
```
A query can be run as DISTINCT by flagging the Distinct property
```vb
MySelect.Distinct
```
#### Example 2
We can add table aliases and joins as well
```sql
SELECT u.id, c.hex FROM users u INNER JOIN colors c ON u.favorite=c.name
```
```vb
Set MySelect = New SQLSelect
With MySelect
    .Fields = Array("u.id", "c.hex")
    .addTable "users", "u"
    .innerJoin "colors", "c", "u.favorite=c.name"
End With
```

### Update
#### Example 1 
To produce this SQL Statement:
```sql
UPDATE users SET username='old_admin' WHERE username='admin'
```
```vb
Set MyUpdate = New SQLUpdate
With MyUpdate
    .Fields = Array("username")
    .Values = Array("old_admin")
    .Table = "users"

    'Need to escape the string
    .AddWhere "username", "=", str("admin") 
End With

MyDatabase.Execute MyUpdate
```
### HelperFunctions
The library includes a handful of helper functions. 
* Date/Time manipulation, toIso() and toUnix().
* String encapsulation str() to add single quotes around strings and escape contained single-quotes

### Unit Tests
If you would like to run the unit tests, import all the library files including the files in "testing" into an Excel workbook. In some cell, type "=RunUnitTests()". Any failures will open a messagebox stating the expected output and what was actually received by the library.
