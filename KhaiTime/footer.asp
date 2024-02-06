<!--#include file="DBconnection.asp"-->


<!DOCTYPE html>
<html lang="en">

<head>
  <meta charset="utf-8">
  <meta content="width=device-width, initial-scale=1.0" name="viewport">
  <%
          set rsTitle = CreateObject("ADODB.Recordset")
          sql = "select * from CompanyInfo"
          rsTitle.open sql, conn
            title = rsTitle("pageTitle")
            name = rsTitle("name")
            logo = rsTitle("logo")
            picture = rsTitle("picture")
    %>

  <title><%=name%></title>
  <meta content="" name="description">
  <meta content="" name="keywords">

  <!-- Favicons -->
  <link href="image\icon.ico" rel="icon">
  <link href="image\icon.ico" rel="apple-touch-icon">

  <!-- Google Fonts -->
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Open+Sans:ital,wght@0,300;0,400;0,500;0,600;0,700;1,300;1,400;1,600;1,700&family=Poppins:ital,wght@0,300;0,400;0,500;0,600;0,700;1,300;1,400;1,500;1,600;1,700&family=Inter:ital,wght@0,300;0,400;0,500;0,600;0,700;1,300;1,400;1,500;1,600;1,700&display=swap" rel="stylesheet">

  <!-- Vendor CSS Files -->
  <link href="assets/vendor/bootstrap/css/bootstrap.min.css" rel="stylesheet">
  <link href="assets/vendor/bootstrap-icons/bootstrap-icons.css" rel="stylesheet">
  <link href="assets/vendor/fontawesome-free/css/all.min.css" rel="stylesheet">
  <link href="assets/vendor/glightbox/css/glightbox.min.css" rel="stylesheet">
  <link href="assets/vendor/swiper/swiper-bundle.min.css" rel="stylesheet">
  <link href="assets/vendor/aos/aos.css" rel="stylesheet">

  <!-- Template Main CSS File -->
  <link href="assets/css/main.css" rel="stylesheet">

  <!-- Font awesome-->
    

  <!-- =======================================================
  * Template Name: Logis
  * Updated: Jan 09 2024 with Bootstrap v5.3.2
  * Template URL: https://bootstrapmade.com/logis-bootstrap-logistics-website-template/
  * Author: BootstrapMade.com
  * License: https://bootstrapmade.com/license/
  ======================================================== -->
</head>

<body>
      

  <!-- ======= Footer ======= -->
  <footer id="footer" class="footer">
    <div class="container">
      <div class="row gy-4">
            <%
                set rs = CreateObject("ADODB.recordset")
                sql = "Select * from CompanyInfo "
                rs.open sql, conn
                name = rs("name")
                address = rs("address")
                taxcode = rs("taxcode")
                logo = rs("logo")
                picture = rs("picture")
                phoneNumber = rs("phoneNumber")
                hotline = rs("hotline")
                email = rs("email")
                pageTitle = rs("pageTitle")
                linkFaceBook = rs("linkFaceBook")
                linkYoutube = rs("linkYouTube")
                linkMap = rs("linkMap")
            %>
            <div class="col-lg-4 col-6 footer-links">
            <h4><%=name%></h4>
            <ul>
                <li>Mã số thuế: <%=taxcode%></li>
                <li>Hotline: <%=hotline%></li>
                <li>Email: <%=email%></li>
                <li><%=address%></li>
            </ul>
            </div>


            <div class="col-lg-4 col-md-12 footer-contact text-center text-md-start">
            <h4>Contact Us</h4>
            <p>
                A108 Adam Street <br>
                New York, NY 535022<br>
                United States <br><br>
                <strong>Phone:</strong> +1 5589 55488 55<br>
                <strong>Email:</strong> info@example.com<br>
            </p>

            </div>

            <div class="col-lg-4 col-md-12 footer-info">
            <a href="index.html" class="logo d-flex align-items-center">
                <span>Map</span>
            </a>
            <p><iframe src="https://www.google.com/maps/embed?pb=!1m18!1m12!1m3!1d14902.842060550984!2d105.82058454947874!3d20.964137390702!2m3!1f0!2f0!3f0!3m2!1i1024!2i768!4f13.1!3m3!1m2!1s0x3135ad1464f75923%3A0xcad6f0f26f397daf!2zQ2h1bmcgY8awIFZQMiBMaW5oIMSQw6Bt!5e0!3m2!1svi!2s!4v1706775591997!5m2!1svi!2s" width="300" height="170" style="border:0;" allowfullscreen="" loading="lazy" referrerpolicy="no-referrer-when-downgrade"></iframe></p>
            <div class="social-links d-flex mt-4">
                <a href="#" class="twitter"><i class="bi bi-twitter"></i></a>
                <a href="#" class="facebook"><i class="bi bi-facebook"></i></a>
                <a href="#" class="instagram"><i class="bi bi-instagram"></i></a>
                <a href="#" class="linkedin"><i class="bi bi-linkedin"></i></a>
            </div>
            </div>

      </div>
    </div>

  </footer>
  <!-- End Footer -->
  </body>
  <html>