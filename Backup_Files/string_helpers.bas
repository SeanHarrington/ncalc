Attribute VB_Name = "string_helpers"
Option Compare Database
Option Explicit ' Explicit 'typing' for variables

' The following function assembles defined SQL statements
' queryType as QueriesType is a public enum in the enums module
' attrs as Variant is an array of string attributes (preferably string) for the statements
' fromTables as Variant is an array of string values (preferably string) for the statements
' whereExpr (whereExpression) is a custom tailored string conditional--easier and more portable this way.

Public Function assemble_query(queryType As QueriesType, ByVal attrs As Variant, ByVal fromTables As Variant _
, ByVal whereExpr As String) As String ' Return  custom query
    If IsNull(attrs) Or IsNull(fromTables) Then ' We must have attributes and from tables for select, insert or delete
        queryType = InvalidQuery ' Overwrite with invalid query
    End If
    
    Select Case queryType
        Case QueriesType.SelectQuery
            assemble_query = gen_select_statement(attrs, fromTables, whereExpr) ' Private call
        Case QueriesType.InsertQuery
            assemble_query = gen_insert_statement() ' Private call
        Case QueriesType.DeleteQuery
            assemble_query = gen_delete_statement()  ' Private call
        Case QueriesType.updateQuery
            assemble_query = gen_update_statement(attrs, fromTables, whereExpr) ' Private call
        Case InvalidQuery
            assemble_query = "" ' Invalid query, return empty
        Case Else
            assemble_query = "" ' Return empty
    End Select
End Function

Private Function gen_select_statement(ByRef selectAttrs As Variant, ByRef fromTables As Variant, ByRef whereExpr As String) _
    As String ' Generate select
    Dim select_part As String: select_part = "SELECT " + gen_attrs_4_sql(selectAttrs)
    Dim from_part As String: from_part = " FROM " + gen_tables_4_sql(fromTables)
        
    gen_select_statement = (select_part + from_part + " " + whereExpr) ' Piece it together and return

End Function



Private Function gen_insert_statement() As String ' Generate insert


End Function

Private Function gen_delete_statement() As String ' Generate delete


End Function

Private Function gen_update_statement(ByRef updateAttrs As Variant, ByRef fromTables As Variant, ByRef whereExpr As String) _
As String ' Generate update
    'Example: UPDATE table_name
    'SET column1=value1, column2,=value2,...
    'WHERE some_column = some_value
    
    Dim update_part As String: update_part = "UPDATE " + gen_tables_4_sql(fromTables)
    Dim set_part As String: set_part = " SET " + gen_attrs_4_sql(updateAttrs)
    
    
    gen_update_statement = (update_part + set_part + " " + whereExpr) ' Piece together update statement and return
    

End Function

Private Function gen_attrs_4_sql(ByRef selectAttrs As Variant)
 ' Assemble the select attributes
    Dim index As Integer
    Dim MAX As Integer: MAX = UBound(selectAttrs) ' Get max count of selectAttrs array
    Dim attrs_list As String: attrs_list = ""
    
    For index = LBound(selectAttrs) To MAX
        If index < MAX Then
            attrs_list = attrs_list + selectAttrs(index) + ","
        ElseIf index = MAX Then
            attrs_list = attrs_list + selectAttrs(index)
        End If
    Next index
    
    gen_attrs_4_sql = attrs_list ' Return list of attributes

End Function

Private Function gen_tables_4_sql(ByRef fromTables As Variant)
    ' Assemble the from tables
    Dim index As Integer
    Dim MAX As Integer: MAX = UBound(fromTables) ' Get max count of fromTables array
    Dim tables_list As String: tables_list = ""
    
    For index = LBound(fromTables) To MAX
        If index < MAX Then
            tables_list = tables_list + fromTables(index) + ","
        ElseIf index = MAX Then
            tables_list = tables_list + fromTables(index)
        End If
    Next index
    
    gen_tables_4_sql = tables_list ' Return list of tables
End Function





