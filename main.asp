<!--#include file="conn.asp" -->
<!DOCTYPE html>
<html>
<head>
    <title>合同管理系统 - 合同列表</title>
    <link rel="stylesheet" type="text/css" href="css/style.css">
</head>
<body>
    <h1>合同列表</h1>
    <a href="add_contract.asp">添加新合同</a>
    <table border="1">
        <tr>
            <th>合同名称</th>
            <th>合同编号</th>
            <th>签订日期</th>
            <th>合同金额</th>
            <th>操作</th>
        </tr>
        <%
        Dim sql, rs
        sql = "SELECT * FROM Contracts"
        Set rs = conn.Execute(sql)
        Do While Not rs.EOF
        %>
        <tr>
            <td><%= rs("ContractName") %></td>
            <td><%= rs("ContractNumber") %></td>
            <td><%= rs("SignDate") %></td>
            <td><%= rs("Amount") %></td>
            <td>
                <a href="edit_contract.asp?id=<%= rs("ID") %>">编辑</a>
                <a href="delete_contract.asp?id=<%= rs("ID") %>" onclick="return confirm('确定要删除该合同吗？')">删除</a>
            </td>
        </tr>
        <%
        rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        %>
    </table>
</body>
</html>
