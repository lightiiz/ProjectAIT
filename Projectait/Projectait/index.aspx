<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="index.aspx.cs" Inherits="Projectait.index" %>


<!DOCTYPE html>
<!--[if lt IE 7]>      <html class="no-js lt-ie9 lt-ie8 lt-ie7"> <![endif]-->
<!--[if IE 7]>         <html class="no-js lt-ie9 lt-ie8"> <![endif]-->
<!--[if IE 8]>         <html class="no-js lt-ie9"> <![endif]-->
<!--[if gt IE 8]><!-->
<html class="no-js" xmlns="http://www.w3.org/1999/xhtml">
<!--<![endif]-->

<head runat="server">
  <!-- BASICS -->
  <meta charset="utf-8"/>
  <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1"/>
  <title>Amoeba free one page responsive bootstrap site template</title>
  <meta name="description" content=""/>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <link rel="stylesheet" type="text/css" href="css/isotope.css" media="screen" />
  <link rel="stylesheet" href="js/fancybox/jquery.fancybox.css" type="text/css" media="screen" />
  <link rel="stylesheet" href="css/bootstrap.css"/>
  <link rel="stylesheet" href="css/bootstrap-theme.css"/>
  <link rel="stylesheet" href="css/style.css"/>
  <!-- skin -->
  <link rel="stylesheet" href="skin/default.css"/>
</head>

<body>
   <form id="form2" runat="server">
  <section id="header" class="appear"></section>
  <div class="navbar navbar-fixed-top" role="navigation" data-0="line-height:100px; height:100px; background-color:rgba(0,0,0,0.3);" data-300="line-height:60px; height:60px; background-color:rgba(0,0,0,1);">
    <div class="container">
      <div class="navbar-header">
        <button type="button" class="navbar-toggle" data-toggle="collapse" data-target=".navbar-collapse">
						<span class="fa fa-bars color-white"></span>
					</button>
        <h1><a class="navbar-brand" href="#" data-0="line-height:90px;" data-300="line-height:50px;">			Gin
					</a></h1>
      </div>
      <div class="navbar-collapse collapse">
        <ul class="nav navbar-nav" data-0="margin-top:20px;" data-300="margin-top:5px;">
          <li class="active"><a href="#">Home</a></li>
          <li><a href="#section-osticket">Osticket</a></li>
          <li><a href="#section-netka">Netka</a></li>
          <li><a href="#section-uploadfile">Uploadfile</a></li>
        </ul>
      </div>
      <!--/.navbar-collapse -->
    </div>
  </div>

  <section class="featured">
    <div class="container">
      <div class="row mar-bot40">
        <div class="col-md-6 col-md-offset-3">

          <div class="align-center">
            <i class="fa fa-flask fa-5x mar-bot20"></i>
            <h2 class="slogan">Welcome to Web report</h2>
            <p>
              Lorem ipsum dolor sit amet, natum bonorum expetendis usu ut. Eum impetus offendit disputationi eu, at vim aliquip lucilius praesent. Alia laudem antiopam te ius, sed ad munere integre, ubique facete sapientem nam ut.

            </p>
          </div>
        </div>
      </div>
    </div>
  </section>

  <!-- services -->
  <section id="section-services" class="section pad-bot30 bg-white">
    <div class="container">

    </div>
  </section>

  <!-- spacer section:testimonial -->
  <section id="testimonials" class="section" data-stellar-background-ratio="0.5">
    <div class="container">
      <div class="row">
        <div class="col-lg-12">
          <div class="align-center">
            <div class="testimonial pad-top40 pad-bot40 clearfix">
              <h5>
        								Nunc velit risus, dapibus non interdum quis, suscipit nec dolor. Vivamus tempor tempus mauris vitae fermentum. In vitae nulla lacus. Sed sagittis tortor vel arcu sollicitudin nec tincidunt metus suscipit.Nunc velit risus, dapibus non interdum.
        							</h5>
              <br/>
              <span class="author">&mdash; MIKE DOE <a href="#">www.siteurl.com</a></span>
            </div>

          </div>
        </div>

      </div>
    </div>
  </section>

  <!-- about -->
  <section id="section-osticket" class="section appear clearfix">
    <div class="container">

      <div class="row mar-bot40">
        <div class="col-md-offset-3 col-md-6">
          <div class="section-header">
            <h2 class="section-heading animated" data-animation="bounceInUp">Osticket Report</h2>
          </div>
        </div>
      </div>

      <div class="row">
        <div class="col-md-4 col-md-offset-4">

            Please input date.<br />
             <div class="form-group">
                <label for="date">date</label>
                <asp:TextBox ID="TextBox1" runat="server"  class="form-control" placeholder="sinc"  data-msg="Please enter a valid date" />
                <div class="validation"></div><br />
                <asp:TextBox ID="TextBox2" runat="server"  class="form-control" placeholder="until"  data-msg="Please enter a valid date" />
                <div class="validation"></div>
              </div>
            <asp:Button ID="Button1" runat="server" class="btn btn-theme pull-left" OnClick="Button1_Click" Text="Search" />
             <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="false" CellPadding="19">
                <Columns>
                    <asp:BoundField DataField="หมายเลขวงจร" HeaderText="หมายเลขวงจร" />
                    <asp:BoundField DataField="หน่วยงานผู้ใช้" HeaderText="หน่วยงานผู้ใช้"  />
                    <asp:BoundField DataField="SLA" HeaderText="SLA"  />
                    <asp:BoundField DataField="วันที่แจ้งเหตุ" HeaderText="วันที่แจ้งเหตุ"  />
                    <asp:BoundField DataField="เวลาแจ้งเหตุ" HeaderText="เวลาแจ้งเหตุ"  />
                    <asp:BoundField DataField="ประเภทเหตุขัดข้อง" HeaderText="ประเภทเหตุขัดข้อง" />
                    <asp:BoundField DataField="สาเหตุ" HeaderText="สาเหตุ" />
                    <asp:BoundField DataField="การแก้ไข" HeaderText="การแก้ไข"/>
                    <asp:BoundField DataField="วันที่แก้ไข" HeaderText="วันที่แก้ไข"  />
                    <asp:BoundField DataField="เวลาแก้ไข" HeaderText="เวลาแก้ไข"  />
                    <asp:BoundField DataField="ระยะเวลาการแก้ไข" HeaderText="ระยะเวลาการแก้ไข"/>
                    <asp:BoundField DataField="OS Ticket Number" HeaderText="OS Ticket Number"/>
                    <asp:BoundField DataField="หมายเหตุ" HeaderText="หมายเหตุ"  />
                    <asp:BoundField DataField="Root Cause" HeaderText="Root Cause"  />
                    <asp:BoundField DataField="Hardware" HeaderText="Hardware"  />
                    <asp:BoundField DataField="ปัญหาที่เกิด" HeaderText="ปัญหาที่เกิด" />
                    <asp:BoundField DataField="Breach/Meet" HeaderText="Breach/Meet"  />
                    <asp:BoundField DataField="ข้อยกเว้น" HeaderText="ข้อยกเว้น"/>
                    <asp:BoundField DataField="วิเคราะห์ Customer" HeaderText="วิเคราะห์ Customer"  />
                </Columns>
            </asp:GridView>
            <br />
            <br />
      </div>   
      </div>
       
    </div>
  </section>
  <!-- /about -->

  <!-- spacer section:stats -->
  <section id="parallax1" class="section pad-top40 pad-bot40" data-stellar-background-ratio="0.5">
    <div class="container">
      <div class="align-center pad-top40 pad-bot40">
        <blockquote class="bigquote color-white">:)</blockquote>
        <p class="color-white">Bootstraptaste</p>
      </div>
    </div>
  </section>


 <!-- netka -->
  <section id="section-netka" class="section appear clearfix">
    <div class="container">

      <div class="row mar-bot40">
        <div class="col-md-offset-3 col-md-6">
          <div class="section-header">
            <h2 class="section-heading animated" data-animation="bounceInUp">Netka Report</h2>
          </div>
        </div>
      </div>

      <div class="row">
        <div class="col-md-4 col-md-offset-4">

            Please input date.<br />
             <div class="form-group">
                <label for="date">date</label>
                <asp:TextBox ID="TextBox3" runat="server"  class="form-control" placeholder="sinc"  data-msg="Please enter a valid date" />
                <div class="validation"></div><br />
                <asp:TextBox ID="TextBox4" runat="server"  class="form-control" placeholder="until"  data-msg="Please enter a valid date" />
                <div class="validation"></div>
              </div>
            <asp:Button ID="Button2" runat="server" class="btn btn-theme pull-left" OnClick="Button2_Click" Text="Search" />
            <br />
            <br />
            <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="false" CellPadding="19">
                <Columns>
                    <asp:BoundField DataField="หน่วยงานผู้ใช้" HeaderText="หน่วยงานผู้ใช้" />
                    <asp:BoundField DataField="Case ID" HeaderText="Case ID"  />
                    <asp:BoundField DataField="ผู้แจ้งเหตุ" HeaderText="ผู้แจ้งเหตุ"  />
                    <asp:BoundField DataField="ผู้รับแจ้งเหตุ" HeaderText="ผู้รับแจ้งเหตุ"  />
                    <asp:BoundField DataField="วันที่แจ้งเหตุ" HeaderText="วันที่แจ้งเหตุ"  />
                    <asp:BoundField DataField="เวลาแจ้งเหตุ" HeaderText="เวลาแจ้งเหตุ"  />
                    <asp:BoundField DataField="เหตุขัดข้อง" HeaderText="เหตุขัดข้อง" />
                    <asp:BoundField DataField="สาเหตุ" HeaderText="สาเหตุ" />
                    <asp:BoundField DataField="การแก้ไข" HeaderText="การแก้ไข"/>
                    <asp:BoundField DataField="วันที่แก้ไข" HeaderText="วันที่แก้ไข"  />
                    <asp:BoundField DataField="เวลาแก้ไข" HeaderText="เวลาแก้ไข"  />
                    <asp:BoundField DataField="ระยะเวลาการแก้ไข" HeaderText="ระยะเวลาการแก้ไข"/>
                    <asp:BoundField DataField="ความรุนแรงของเหตุการ" HeaderText="ความรุนแรงของเหตุการ"/>
                    <asp:BoundField DataField="หมายเหตุ" HeaderText="หมายเหตุ"  />
                    <asp:BoundField DataField="Root Cause" HeaderText="Root Cause"  />
                    <asp:BoundField DataField="Case Category" HeaderText="Case Category"  />
                    <asp:BoundField DataField="Case Sub Category" HeaderText="Case Sub Category" />
                </Columns>
            </asp:GridView>
      </div>   
      </div>
       
    </div>
  </section>
  <!-- /netka -->

     <section id="parallax2" class="section parallax" data-stellar-background-ratio="0.5">
    <div class="align-center pad-top40 pad-bot40">
      <blockquote class="bigquote color-white">:D</blockquote>
      <p class="color-white">Bootstraptaste</p>
    </div>
  </section>

  <!-- uploadfile -->
  <section id="section-uploadfile" class="section appear clearfix">
    <div class="container">

      <div class="row mar-bot40">
        <div class="col-md-offset-3 col-md-6">
          <div class="section-header">
            <h2 class="section-heading animated" data-animation="bounceInUp">Upload Excel File</h2>
          </div>
        </div>
      </div>
      <div class="row">
        <div class="col-md-8 col-md-offset-2">

         <table class="auto-style1">
                <tr>
                    <td class="auto-style3">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Upload Excel File&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>
                    <td class="auto-style4">
                        <asp:FileUpload ID="FileUpload1"  runat="server" />
                        <br />
                    </td>
                </tr>
                <tr>
                    <td class="auto-style2">
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <asp:Label ID="Label1" runat="server" Text=""></asp:Label>
                    </td>
                    <td class="auto-style5">
                        <asp:Button ID="Button3" runat="server" class="btn btn-theme pull-left" OnClick="Button3_Click" Text="Upload and Save to SQL Server" Height="33px" Width="252px" />
                    </td>
                </tr>
            </table>
         
          </div>
        </div>
        <!-- ./span12 -->
      </div>
  </section>
  
  <a href="#header" class="scrollup"><i class="fa fa-chevron-up"></i></a>

  <script src="js/modernizr-2.6.2-respond-1.1.0.min.js"></script>
  <script src="js/jquery.js"></script>
  <script src="js/jquery.easing.1.3.js"></script>
  <script src="js/bootstrap.min.js"></script>
  <script src="js/jquery.isotope.min.js"></script>
  <script src="js/jquery.nicescroll.min.js"></script>
  <script src="js/fancybox/jquery.fancybox.pack.js"></script>
  <script src="js/skrollr.min.js"></script>
  <script src="js/jquery.scrollTo.js"></script>
  <script src="js/jquery.localScroll.js"></script>
  <script src="js/stellar.js"></script>
  <script src="js/jquery.appear.js"></script>
  <script src="https://maps.googleapis.com/maps/api/js?key=AIzaSyD8HeI8o-c1NppZA-92oYlXakhDPYR7XMY"></script>
  <script src="js/main.js"></script>
  <script src="contactform/contactform.js"></script>
</form>
</body>

</html>
