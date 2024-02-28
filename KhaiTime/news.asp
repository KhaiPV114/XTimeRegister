<%@  language="VBSCRIPT" codepage="65001" %>
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
  <link href="./image/icon.ico" rel="icon">
  <link href="assets/img/apple-touch-icon.png" rel="apple-touch-icon">

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

  <!-- =======================================================
  * Template Name: Logis
  * Updated: Jan 09 2024 with Bootstrap v5.3.2
  * Template URL: https://bootstrapmade.com/logis-bootstrap-logistics-website-template/
  * Author: BootstrapMade.com
  * License: https://bootstrapmade.com/license/
  ======================================================== -->
</head>

<body>

  <!-- ======= Header ======= -->
  <!--#include file="header.asp"-->
  <!-- End Header -->

  <main id="main">

    <!-- ======= Breadcrumbs ======= -->
    <div class="breadcrumbs">
      <div class="page-header d-flex align-items-center" style="background-image: url('./image/15.png');">
        <div class="container position-relative">
          <div class="row d-flex justify-content-center">
            <%
                    set rsArticle = CreateObject("ADODB.recordset")
                    sql = "select * from Article a join Category_Article ca on a.articleId = ca.articleId  where a.interface = 1 and ca.categoryID = 4"
                    rsArticle.open sql, conn
                      articleId = rsArticle("articleId")
                      articleTitle = rsArticle("articleTitle")
                      articleBody = rsArticle("articleBody")
                      interface = rsArticle("interface")
              %>
            <div class="col-lg-6 text-center">
              <h2><%=articleTitle%></h2>
              <p></p>
            </div>
            <%
                rsArticle.close
            %>
          </div>
        </div>
      </div>
      <nav>
        <div class="container">
          <ol>
            <li><a href="index.html">Home</a></li>
            <li>Services</li>
          </ol>
        </div>
      </nav>
    </div><!-- End Breadcrumbs -->

    <!-- ======= News Section ======= -->
    <section id="service" class="services pt-0">
      <div class="container" data-aos="fade-up">
          
        <%
                    set rsArticle = CreateObject("ADODB.recordset")
                    sql = "select * from Article a join Category_Article ca on a.articleId = ca.articleId  where a.interface = 2 and ca.categoryID = 4"
                    rsArticle.open sql, conn
                      articleId = rsArticle("articleId")
                      articleTitle = rsArticle("articleTitle")
                      articleBody = rsArticle("articleBody")
                      interface = rsArticle("interface")
        %>

        <div class="section-header">
          <span><%=articleTitle%></span>
          <h2><%=articleTitle%></h2>

        </div>

        <div class="row gy-4">
          <%
                set rsArticle = CreateObject("ADODB.recordset")
              sql = "select * from Article a join Category_Article ca on a.articleId = ca.articleId  where a.interface = 7 and ca.categoryID = 1"
                  rsArticle.open sql, conn
                  Do Until rsArticle.eof
                    articleId = rsArticle("articleId")
                    articleTitle = rsArticle("articleTitle")
                    articleBody = rsArticle("articleBody")
                    interface = rsArticle("interface")
          %>
          <div class="col-lg-4 col-md-6" data-aos="fade-up" data-aos-delay="100">
          
            <div class="card">
              <%
                set rsPicture = CreateObject("ADODB.recordset")
                  sql = "select * from Article_Picture where articleId = " & articleId
                  rsPicture.open sql, conn
                  Do Until rsPicture.eof
                    pictureName = rsPicture("pictureName")
                    pictureUrl = rsPicture("pictureUrl")
          %>
              <div class="card-img">
                <img src="<%=pictureUrl%>" alt="<%=pictureName%>" class="img-fluid">
              </div>
              <%
                rsPicture.movenext
                loop
                rsPicture.close
              %>
              <h3><a href="service-details.html" class="stretched-link"><%=articleTitle%></a></h3>
              <p><%=articleBody%></p>
            </div>
            <!-- End Card Item -->
            </div>
          <%
            rsArticle.movenext
            loop
            rsArticle.close
          %>
        </div>
        

      </div>
    </section><!-- End News Section -->

  </main><!-- End #main -->

  <!-- ======= Footer ======= -->
    <!--#include file="footer.asp"-->

  <!-- End Footer -->

  <a href="#" class="scroll-top d-flex align-items-center justify-content-center"><i class="bi bi-arrow-up-short"></i></a>

  <div id="preloader"></div>

  <!-- Vendor JS Files -->
  <script src="assets/vendor/bootstrap/js/bootstrap.bundle.min.js"></script>
  <script src="assets/vendor/purecounter/purecounter_vanilla.js"></script>
  <script src="assets/vendor/glightbox/js/glightbox.min.js"></script>
  <script src="assets/vendor/swiper/swiper-bundle.min.js"></script>
  <script src="assets/vendor/aos/aos.js"></script>
  <script src="assets/vendor/php-email-form/validate.js"></script>

  <!-- Template Main JS File -->
  <script src="assets/js/main.js"></script>

</body>

</html>