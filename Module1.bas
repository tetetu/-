Attribute VB_Name = "Module1"
Option Compare Database
Option Explicit

Sub SQL_Run()
Call connect
Dim str_SQL As String
str_SQL = ""
str_SQL = str_SQL & ""
str_SQL = str_SQL & "CREATE TABLE CardInfo ("
str_SQL = str_SQL & "CardID nchar(6),"
str_SQL = str_SQL & "CustomerID nchar(5),"
str_SQL = str_SQL & "IssueDate datetime,"
str_SQL = str_SQL & "ExpireDate datetime,"
str_SQL = str_SQL & "EmployeeID int"
str_SQL = str_SQL & ")"
BeginTransaction
Call execute(str_SQL)
CommitTransaction
End Sub

  
  
  
  
  

