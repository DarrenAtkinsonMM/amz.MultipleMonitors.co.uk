<%
'========================================================================
'// SQL FUNCTIONS AND ROUTINES
'========================================================================
Function TrapSQLError(varTableName)		
    '// -2147217900 = Table 'x' already exists.
    '// -2147217887 = Field 'x' already exists in table 'x'.
    if ((Err.Number=-2147217900) OR (Err.Number=-2147217887)) then
        Err.Description=""
        err.number=0
    else
        ErrStr = ErrStr & "Error Creating Table "&varTableName&": "&Err.Description&"<BR>"
        err.number=0
        iCnt=iCnt+1
    end if
End Function

Function UpdateTableIfValue(tableName, fieldName, searchStr, checkValue, newValue)
    query="SELECT " & fieldName & " FROM " & tableName & " " & searchStr & ";"
    set rs=conntemp.execute(query)
    if not rs.eof then
        if rs(fieldName) = checkValue then
            query="UPDATE " & tableName & " SET " & fieldName & " = '" & newValue & "';" 
            conntemp.execute(query)
        end if
    end if
    set rs=nothing
End Function

Function TableExists(tableName)
    query="IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = N'" & tableName & "') SELECT 1 ELSE SELECT 0;"
    set rs=conntemp.execute(query)
    If rs(0) = "1" Then
        TableExists = true
    Else
        TableExists = false
    End If
    set rs=nothing
    err.clear
    err.number = 0
End Function

Sub AlterTableSQL(strTable,strType,strColName,strDataType,isDefault,intDefaultVal,isNULLFlag)
    on error resume next
    Err.Description=""
    err.number=0
    '//  Add column pcStoreSettings_DisableGiftRegistry for table pcStoreSettings
    query="ALTER TABLE ["&strTable&"] "&strType&" ["&strColName&"] "&strDataType&" NULL;"
    set rs=server.CreateObject("ADODB.RecordSet")
    set rs=conntemp.execute(query)
    if err.number <> 0 then
        '// COLUMN NAMES IN EACH TABLE MUST BE UNIQUE
        if Err.Number = -2147217900 then
            Err.Description=""
            err.number=0
        else
            if Err.Number = -2147217887 then
                Err.Description=""
                err.number=0
            else
                ErrStr=ErrStr&"Unable to update TABLE "&strTable&" COLUMN "&strColName&" - Error: "&Err.Description&"<BR>"
                Err.Description=""
                err.number=0
                iCnt=iCnt+1
            end if
        end if
    else
        if isDefault = 1 OR isDefault = 2 then
            if isDefault = 2 then
                queryDel = "'"
            end if
            query="UPDATE "&strTable&" SET "&strColName&"="&queryDel&intDefaultVal&queryDel&";"
            
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            query="ALTER TABLE ["&strTable&"] ADD CONSTRAINT [DF_"&strTable&"_"&strColName&"] DEFAULT ("&queryDel&intDefaultVal&queryDel&") FOR ["&strColName&"];"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            if isNULLFlag="1" then
                query="ALTER TABLE "&strTable&" ALTER COLUMN "&strColName&" "&strDataType&" NOT NULL"
                set rs=server.CreateObject("ADODB.RecordSet")
                set rs=conntemp.execute(query)
            end if
        end if
    end if
End Sub

'========================================================================
'// END OF SQL FUNCTIONS and ROUTINES
'========================================================================
%>