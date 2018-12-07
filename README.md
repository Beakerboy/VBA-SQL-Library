VBA SQL Library
=====================

### Object-Based Database Interaction for VBA

Easily create SQL queries and execute them on a database. For an example on how to use this library with data objects, see the [VBA-Drupal-Library](https://github.com/Beakerboy/VBA-Drupal-Library). For other examples on using the library, refer to the unit tests.

Features
--------
 * [Database](#database)
 * [Login Form](#login-form)
 * [Static Queries](#static-queries)
 * [Insert](#insert)
 * [Select](#select)
 * [Update](#update)
 * [Helper Functions](#helper-functions)
 * [Unit Tests](#unit-tests)
 
 Setup
-----

Download the Addin (SQLlib.xlam) and enable it in MSExcel. Open Microsoft Visual Basic For Applications, select Tools>References and ensure that both "Microsoft ActiveX Data Objects x.x Library", and SQLlib is selected.
 
 Security
-----
This Library allows developers to create static or dynamic SQL statements using VBA objects. If the table names and field names are all known by the developer, and only field values, and conditional values will be supplied by the user, an SQLStaticQuery might be the best option. All user-supplied information will be sanitized before being added to the query. It also provides a login box to discourage hard-coding database authentication details. The dynamic query generating objects are best for cases where table names and field names are part of larger data objects, and the queries themselves are created by a larger system. This larger system should provide data sanitizing options to ensure malicious data does make it into a query. the [VBA-Drupal-Library](https://github.com/Beakerboy/VBA-Drupal-Library) is an example of such a system.

 Testing
 -----
The unit tests demonstrate many ways to use each of the classes. To run the tests, Import all the modules from the testing directory into a spreadsheet, and run the SQL_RunUnitTests function. Ensure the Setup steps have all been successfully completed.
 
 Usage
-----

### Database
Create a new database connection:
```vb
Dim MyDatabase As SQLDatabase
Set MyDatabase = Create_SQLDatabase
MyDatabase.DBType = "mssql"
MyDatabase.DSN = "foodb"
```
Several different types of database execution can occur:
 * Execute(SQL) - Execute a statement (Insert or Update)
 * InsertGetNewID(SQLInsert) - Insert a record, and return the new primary key
 * Execute(SQLSelect, column) - Execute a statement and return a single value
 * Execute(SQLSelect) Execute a statement and return an array of values

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
### Static Queries
Developers can create static queries, while ensuring that user inputed data will interact with the database successfully.
Items in bold are required
 * .Query = __query__
 * .AddArgument __placeholder__, __value__
 * .ClearArguments
 
#### Example 1
```vb
Dim MyStaic as SQLStaticQuery
Set MyStatic = Create_SQLStaticQuery
MyStatic.Query = "SELECT name FROM users WHERE id=:id"
MYStatic.addArgument ":id", 4
```
Will produce the SQL
```sql
SELECT name FROM users WHERE id=4;
```
The SQL statement can be easily reused with different user-supplied values for the ID without the need to recreate the object.

### Insert
The SQLInsert Object has many options. Items in bold are required
 * .Table     = __table__
 * .Fields    = Array(__field1__, _field2_, ...)
 * .Values    = Array(__value1__, _value2_, ...)
 * .From      = __SQLSelect__
 * .Returning = __field__

The Insert object can create both INSERT VALUES and INSERT SELECT statements. Multiple inserts can be performed in one statement if the values array is 2 Dimensional.

#### Example 1 - Insert Values
To produce this SQL Stament:
```sql
INSERT INTO users (username, first_name, password) VALUES ('admin', 'Alice', 'secret');
```

Use the Following in VBA-SQL-Library
```vb
'Initialize the object and assign a table name
Set MyInsert = Create_SQLInsert
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
'Create the SELECT Statement
Set SQL = Ceate_SQLSelect

'We don't escape the "user_id" because it is a field name, not a string
Sql.Fields = Array(10, 5770000051, "user_id")
Sql.Table = "users"
Sql.addWhere "username", str("admin")

'Initialize the object and assign a table name
Set MyInsert = Create_SQLInsert
With MyInsert
    .table = "bank_account"
    .Fields = Array("account_number", "open_date", "user_id")
    Set .From = Sql
End With

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
Set MyInsert = Create_SQLInsert
MyInsert.table = "users"

'Set The Fields
MyInsert.Fields = Array("username", "first_name", "password")

'Set the Values
Dim Values2D(1) As Variant
Values2D(0) = Array("'admin'", "'Alice'", "'secret'")
Values2D(1) = Array("'editor'","'Bob'", "'super-secret'")
MyInsert.Values = Values2D

'Execute the query
MyDatabase.Execute MyInsert 
```

### Select
The Select Object has many options. Items in bold are required
 * .Table = __table__
 * .addTable __table__, _alias_
 * .Fields = Array(__field1__, _field2_, ...)
 * .AddField __field__, _alias_
 * .AddExpression __expression__, _alias_    
 * .Distinct
 * .InnerJoin __table__, _alias_, _condition_
 * .LeftJoin __table__, _alias_, _condition_
 * .RightJoin __table__, _alias_, _condition_
 * .AddJoin __joinType__, __table__, _alias_, _condition_
 * .OrderBy __field__, _direction_
 * .AddWhere __field__, __value__, _operation_, _groupType_
 * .GroupBy __field_
 * .AddHaving __field__, __value__, _operation_, _groupType_
 * .Union __query__, _type_

#### Example 1
We can execute a select statement and receive the results as a single value, or an array of values:
```sql
SELECT id FROM users WHERE username='admin';
```

```vb
Set MySelect = Create_SQLSelect
With MySelect
    .Fields = Array("id")
    .Table = "users"

    'Need to escape the string
    .AddWhere "username", str("admin")
End With

ID = MyDatabase.Execute(MySelect, "id")
```
WHERE clauses can be added and grouped together. The following changes the query to:
```sql
SELECT id FROM users WHERE username='admin' AND id<10;
```
```vb
MySelect.AddWhere "id", 10, "<", "AND"
```
A SQLWhereGroup can abe be added using SQLSELECT.AddWhereGroup. This is necessary for a where clause like:
```sql
SELECT id FROM users WHERE (a=1 AND b=2) OR (c = 3 AND d = 4)
```
The SQLSelect Object can create Queries with "GROUP BY ... HAVING ..." sections.
```sql
... GROUP BY user_type HAVING age>1;
```
```vb
MySelect.GroupBy = Array("user_type")
MySelect.AddHaving = "age", "1", ">"
```
A query can be run as DISTINCT by flagging the Distinct property
```vb
MySelect.Distinct
```
#### Example 2
We can add table aliases and joins as well
```sql
SELECT u.id, c.hex FROM users u INNER JOIN colors c ON u.favorite=c.name ORDER BY u.id DESC
```
```vb
Set MySelect = Create_SQLSelect
With MySelect
    .Fields = Array("u.id", "c.hex")
    .addTable "users", "u"
    .innerJoin "colors", "c", "u.favorite=c.name"
    .OrderBy "u.id", "DESC"
End With
```

### Update
#### Example 1 
To produce this SQL Statement:
```sql
UPDATE users SET username='old_admin' WHERE username='admin'
```
```vb
Set MyUpdate = Create_SQLUpdate
With MyUpdate
    .Fields = Array("username")
    .Values = Array("old_admin")
    .Table = "users"

    'Need to escape the string
    .AddWhere "username", str("admin") 
End With

MyDatabase.Execute MyUpdate
```
### HelperFunctions
The library includes a handful of helper functions. 
* Date/Time manipulation, toIso() and toUnix().
* String encapsulation str() to add single quotes around strings and escape contained single-quotes

### Unit Tests
If you would like to run the unit tests, import all the library files including the files in "testing" into an Excel workbook. In some cell, type "=RunUnitTests()". Any failures will open a messagebox stating the expected output and what was actually received by the library.
