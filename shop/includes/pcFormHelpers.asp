<%

Sub addFormInput(label, fieldType, fieldID, fieldName, fieldValue, fieldSize, fieldRequired, fieldSubTitle)
	If fieldType <> "text" And fieldType <> "hidden" And fieldType <> "email" And fieldType <> "password" Then
		fieldType = "text"
	End If
	
	attrs = ""
	If Len(fieldID) > 0 Then attrs = attrs & "id='" & fieldID & "'"
	If Len(fieldName) > 0 Then attrs = attrs & " name='" & fieldName & "'"
	If Len(fieldValue) > 0 Then attrs = attrs & " value='" & fieldValue & "'"
	If Len(fieldSize) > 0 Then attrs = attrs & " size='" & fieldSize & "'"
	
%>
    <% formItemStart label, fieldName, fieldRequired  %>
    <input type="<%= fieldType %>" class="form-control" <%= attrs %> />
    <% If Not IsNull(fieldRequired) Then %>
        <% pcs_RequiredImageTagHorizontal fieldName, fieldRequired %>
    <% End If %>
    <% If Len(fieldSubTitle) > 0 Then %>
        <span class="help-block"><%= fieldSubTitle %></span>
    <% End If %>
  <% formItemEnd %>
  
<%
End Sub

Sub formItemStart(label, pcfield, isRequired)
%>

  <div class="form-group">
    <label for="<%=pcfield%>" class="control-label"><%= label %><% If isRequired=True Then %><span class="pcRequiredIcon"><img src="<%=pcv_strRequiredIcon%>" alt="Required"></span><% End If %></label>

<%
End Sub

Sub formItemEnd()
%>
</div> 
<%
End Sub
%>