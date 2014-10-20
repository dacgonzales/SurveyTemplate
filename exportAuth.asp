<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title></title>
    </head>
    <body style="text-align:center;">
        <form id="exportPasswordProc" method="post" action="save_excel.asp">
            <div style="width:100%; text-align:center;">
                <div style="width:400px; margin-left:auto; margin-right:auto; text-align:center;">
                <div style="height:400px;"></div>
                    <table width="100%">
                        <tr>
                            <td>
                                Please enter password:
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <input type="password" id="expPass" name="expPass" />
                                <br />
                                <input type="submit" value="Download" />
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <%if request.querystring("qryErr") = 1 then %>
                                    <b style="color:red">Please enter the correct password.</b>
                                <%end if %>
                            </td>
                        </tr>
                    </table>
                </div>
            </div>
        </form>
    </body>
</html>
