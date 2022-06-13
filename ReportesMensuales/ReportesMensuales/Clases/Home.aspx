<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Home.aspx.cs" Inherits="ReportesMensuales.Clases.Home" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
 <title>Home</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <%--CSS--%>
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.1.1/css/bootstrap.min.css" />
    <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.6.3/css/all.css"  />
    <link rel="stylesheet" href="../css/StyleSheet.css" />
    <link rel="stylesheet" href="../PNotify/animate.css" />
    <link rel="stylesheet" href="../PNotify/pnotify.custom.min.css" />

    <%--JS--%>
      <script src="https://code.jquery.com/jquery-3.3.1.slim.min.js" integrity="sha384-q8i/X+965DzO0rT7abK41JStQIAqVgRVzpbzo5smXKp4YfRvH+8abtTE1Pi6jizo" crossorigin="anonymous"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.14.3/umd/popper.min.js" ></script>
    <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.1.1/js/bootstrap.min.js"></script>
    <script src="../js/Sesion.js"></script>
     <script src="../PNotify/pnotify.custom.min.js"></script>
</head>
<body>
    <form id="form1" runat="server">
        
        <asp:ScriptManager ID="ajaxHome" runat="server" />
       <asp:UpdatePanel ID="renderPanel" runat="server" UpdateMode="Conditional">
           <ContentTemplate>
        <div style="background-color: #9d2424; width: 100%;height: 130px;">
             <asp:Label runat="server" CssClass="d-flex justify-content-center text-white" style="font-size: 30px;font-weight: bold;">Reportes Mensuales</asp:Label>
        </div>

                            <div class="col-sm-12 d-flex justify-content-center m-0 p-2 form-group">
                            <div class="col-sm-4  m-0 p-2 form-group">
                            <asp:Label runat="server" CssClass="" Style="font-weight: bold;color: gray;">Fecha Inicio</asp:Label>
                            <asp:TextBox type="date" CssClass="form-control" ID="fecIni" runat="server" name="Fecha" placeholder="Fecha Inicio"/>
                            </div>

                            <div class="col-sm-4  m-0 p-2 form-group">
                            <asp:Label runat="server" CssClass="" Style="font-weight: bold;color: gray;">Fecha Fin</asp:Label>
                            <asp:TextBox type="date" CssClass="form-control" ID="fecFin" runat="server" name="Fecha" placeholder="Fecha Fin"/>
                            </div>
                            </div>

                            <div class="container p-5 my-5">
                            <asp:Label runat="server" CssClass="d-flex justify-content-center" style="font-size: 30px;font-weight: bold;color: gray;">Seleccione el reporte que desea generar</asp:Label>
                            </div>


                            <div class="container p-5 my-5">
                            <div class="row">  
                              <div class="col-sm-3"><asp:LinkButton runat="server" OnClick="GeneraReporteUno" class="btn btn-secondary">Reporte 1</asp:LinkButton></div>
                              <div class="col-sm-3"><asp:LinkButton runat="server" class="btn btn-secondary">Reporte 2</asp:LinkButton></div>
                              <div class="col-sm-3"><asp:LinkButton runat="server" class="btn btn-secondary">Reporte 3</asp:LinkButton></div>
                              <div class="col-sm-3"><asp:LinkButton runat="server" class="btn btn-secondary">Reporte 4</asp:LinkButton></div>
                            </div>
                            </div>

               
                <asp:UpdateProgress runat="server" ID="progressPanel" runAssociatedUpdatePanelID="renderPanel" AsyncPostBackTimeout="36000">
                    <ProgressTemplate>
                        <!-- The Modal -->
                        <asp:Panel runat="server" ID="mascara" Style="font-size: 50vh">
                           <i class="fas fa-spinner fa-pulse" style="color: white"></i>
                       </asp:Panel>
                   </ProgressTemplate>
                </asp:UpdateProgress>
                
        </ContentTemplate>
       </asp:UpdatePanel>
            
    </form>
</body>
</html>
