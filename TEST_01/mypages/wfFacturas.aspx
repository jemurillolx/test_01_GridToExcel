<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="wfFacturas.aspx.cs" Inherits="TEST_01.mypages.wfFacturas" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous">
    
</head>
 
<body>
    <form id="form1" runat="server">
        <div>
            <div class="row justify-content-center"" >
                <div class="col-2">
                    <label for="inputMonto" class="col-form-label">Monto Factura</label>
                  
                </div>
                <div class="col-6 offset-1">
                    <asp:TextBox ID="tbMontoFactura" class="form-control" runat="server" placeholder="Monto Factura"> </asp:TextBox>
                   
                </div>

                <div class="col-4">
                </div>

            </div>
            <div class="row justify-content-center"">
                <div class="col-3">
                    
                    <label for="inputcl" class="col-form-label">Tipos de Clientes</label>
                   
                        
                        
                        <asp:CheckBoxList ID="lstClientes" runat="server"></asp:CheckBoxList>
                      
                    
                </div>
                <div class="col-2">
                    
                    <label for="inputcl" class="col-form-label">Fecha Inicial</label>

                    <asp:Calendar ID="Calendar1" runat="server" BackColor="White" BorderColor="#3366CC" BorderWidth="1px" CellPadding="1" DayNameFormat="Shortest" Font-Names="Verdana" Font-Size="8pt" ForeColor="#003399" Height="200px" Width="220px">
                        <DayHeaderStyle BackColor="#99CCCC" ForeColor="#336666" Height="1px" />
                        <NextPrevStyle Font-Size="8pt" ForeColor="#CCCCFF" />
                        <OtherMonthDayStyle ForeColor="#999999" />
                        <SelectedDayStyle BackColor="#009999" Font-Bold="True" ForeColor="#CCFF99" />
                        <SelectorStyle BackColor="#99CCCC" ForeColor="#336666" />
                        <TitleStyle BackColor="#003399" BorderColor="#3366CC" BorderWidth="1px" Font-Bold="True" Font-Size="10pt" ForeColor="#CCCCFF" Height="25px" />
                        <TodayDayStyle BackColor="#99CCCC" ForeColor="White" />
                        <WeekendDayStyle BackColor="#CCCCFF" />
                    </asp:Calendar>
                </div>
                <div class="col-2">
                    
                    <label for="inputcl" class="col-form-label">Fecha Final</label>

                   
                    <asp:Calendar ID="Calendar2" runat="server" BackColor="White" BorderColor="#3366CC" BorderWidth="1px" CellPadding="1" DayNameFormat="Shortest" Font-Names="Verdana" Font-Size="8pt" ForeColor="#003399" Height="200px" Width="220px">
                        <DayHeaderStyle BackColor="#99CCCC" ForeColor="#336666" Height="1px" />
                        <NextPrevStyle Font-Size="8pt" ForeColor="#CCCCFF" />
                        <OtherMonthDayStyle ForeColor="#999999" />
                        <SelectedDayStyle BackColor="#009999" Font-Bold="True" ForeColor="#CCFF99" />
                        <SelectorStyle BackColor="#99CCCC" ForeColor="#336666" />
                        <TitleStyle BackColor="#003399" BorderColor="#3366CC" BorderWidth="1px" Font-Bold="True" Font-Size="10pt" ForeColor="#CCCCFF" Height="25px" />
                        <TodayDayStyle BackColor="#99CCCC" ForeColor="White" />
                        <WeekendDayStyle BackColor="#CCCCFF" />
                    </asp:Calendar>

                </div>
                

            </div>
             <div class="row justify-content-center">
                <div class="col-2"><asp:Button ID="BtVerInfo" runat="server" Text="Ver información" class="btn btn-info" OnClick="BtVerInfo_Click" />

                </div>
                <div class="col-1">
                    <asp:ImageButton ID="imbExcel" runat="server" ImageUrl="~/mypages/bexcel.jpg" widt="50px" Height="50px" OnClick="imbExcel_Click"/>
                </div>

            </div>
           

            <div class="row">
                <div class="col-12">
                    
                <asp:GridView ID="GridView1" runat="server" class="table table-striped table-bordered" style="width:100%">
                </asp:GridView>
                </div>
            </div>
            </div>
    </form>

   


     <!-- Latest jQuery minified -->
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.2.1/jquery.min.js"></script>
    <!-- Latest compiled and minified JavaScript -->
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js" integrity="sha384-Tc5IQib027qvyjSMfHjOMaLkfuWVxZxUPnCJA7l2mCWNIpG9mGCD8wGNIcPD7Txa" crossorigin="anonymous"></script>
    <script type="text/javascript" src="https://cdn.datatables.net/v/bs/dt-1.10.15/datatables.min.js"></script>


    <script src="https://code.jquery.com/jquery-3.2.1.slim.min.js" integrity="sha384-KJ3o2DKtIkvYIK3UENzmM7KCkRr/rE9/Qpg6aAZGJwFDMVNA/GpGFF93hXpG5KkN" crossorigin="anonymous"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.9/umd/popper.min.js" integrity="sha384-ApNbgh9B+Y1QKtv3Rn7W3mgPxhU9K/ScQsAP7hUibX39j7fakFPskvXusvfa0b4Q" crossorigin="anonymous"></script>
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js" integrity="sha384-JZR6Spejh4U02d8jOt6vLEHfe/JQGiRRSQQxSfFWpi1MquVdAyjUar5+76PVCmYl" crossorigin="anonymous"></script>
</body>
</html>