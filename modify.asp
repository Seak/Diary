<!--#include file="conn.asp"-->
<!--#include file="inc_dim.asp"-->
<!--#include file="inc_config.asp"-->
<% Title = Pro_Name & "—资料" %>
<%
strPassword = Trim(Request.Form("Password"))
strMail = Trim(Request.Form("Mail"))

If Session.Contents("PN_MUDiaryMaster_" & Session.Contents("strUserName")) And strPassword <> "" And strMail <> "" Then
	ConnectionDatabase
	strSQL = "Update [UserList] Set strPassword = '"& strPassword &"', strMail = '"& strMail &"' Where strUserName = '" & Session.Contents("strUserName") & "'"
	Conn.Execute(strSQL)
	Set strSQL = Nothing
	Set Rs = Nothing
	CloseDatabase
	Title = Pro_Name & "—修改成功"
	strMsg = "您的资料修改成功！"
End If
%>
<!--#include file="inc_header.asp"-->
<table border="0" align="center" cellpadding="0" cellspacing="0" style="BORDER: 1px solid #E2E2E2;"> 
<tr> 
  <td width="600" height="30" align="center" valign="middle" background="img/td_bg.gif"><%= Title %></td>
</tr>
<tr> 
  <td align="center" valign="middle"> <br> <table width="500" border="0" align="center" cellpadding="0" cellspacing="0" style="Word-Break:break-all; BORDER: 1px solid #E2E2E2;"> 
    <tr> 
      <td> <div align="center"> 
<% If Session.Contents("PN_MUDiaryMaster_" & Session.Contents("strUserName")) Then %>
<% If strMsg <> "" Then %><%= strMsg %><% Else %><%
ConnectionDatabase
strSQL = "Select strUserName, strPassword, strMail From [UserList] Where strUserName = '" & Session.Contents("strUserName") & "'"
Set Rs = Conn.Execute(strSQL)
%><form name="modify" method="post" action="modify.asp"> 
        昵称： <input name="UserName" type="text" style="BACKGROUND-COLOR: #F6F6F6; BORDER: 1 SOLID; FONT-SIZE: 9pt" value="<%= Rs("strUserName") %>" size="15" ReadOnly>
                <br>
                密码： 
                <input name="Password" type="password" style="BACKGROUND-COLOR: #F6F6F6; BORDER: 1 SOLID; FONT-SIZE: 9pt" value="<%= Rs("strPassword") %>" size="15">
                <br>
                信箱：
                <input name="Mail" type="text" style="BACKGROUND-COLOR: #F6F6F6; BORDER: 1 SOLID; FONT-SIZE: 9pt" value="<%= Rs("strMail") %>" size="15">
                <br>
                <input name="Submit" type="submit" style="BACKGROUND-COLOR: #F6F6F6; BORDER: 1 SOLID; FONT-SIZE: 9pt" value="提交">　<input name="Reset" type="reset" style="BACKGROUND-COLOR: #F6F6F6; BORDER: 1 SOLID; FONT-SIZE: 9pt" value="重置">
            </form>
<%
Set Rs = Nothing
Set strSQL = Nothing
CloseDatabase
%>
<% End If %>
<% End If %>
              </div>
          </td>
        </tr>
      </table>
      <br>
    </td>
  </tr>
</table>
<!--#include file="inc_footer.asp"-->