Imports System.Data.OleDb
Module ModCommonClasses
    Public rdr As OleDbDataReader = Nothing
    Public dtable As DataTable
    Public con As OleDbConnection = Nothing
    Public adp As OleDbDataAdapter
    Public ds As DataSet
    Public cmd As OleDbCommand = Nothing
    Public dt As New DataTable
End Module
