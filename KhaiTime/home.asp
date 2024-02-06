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

  <!-- ======= Hero Section ======= -->

  <section id="hero" class="hero d-flex align-items-center">
  <div class="container-fluid">
    <div class="row gy-4 d-flex justify-content-between">
    <div class="container">
      <div class="row gy-4 d-flex justify-content-between">
        <div class="col-lg-8 order-2 order-lg-1 d-flex flex-column justify-content-center ms-5">
          <%
          set rsArticle = CreateObject("ADODB.recordset")
            sql = "select * from Article "
            rsArticle.open sql, conn
            Do Until rsArticle.eof
              articleId = rsArticle("articleId")
              articleTitle = rsArticle("articleTitle")
              articleBody = rsArticle("articleBody")
              interface = rsArticle("interface")
          %>

          <%
            If (interface = 1) Then
          %>
          <%If (articleId = 1) Then%>
          <h2 data-aos="fade-up"><%=articleTitle%></h2>
          <p data-aos="fade-up" data-aos-delay="100"><%=articleBody%></p>
          <%
          End If%>

          <div class="row gy-4" data-aos="fade-up" data-aos-delay="400">
            <%If (articleId = 3) Then%>
            <h4 class="ms-5"><%=articleBody%></h4>
            <%
              set rsItem = CreateObject("ADODB.recordset")
              sql = "select * from Article_Items where articleId = " & articleId
              rsItem.open sql, conn
              Do Until rsItem.eof
                itembody = rsItem("itemBody")
              %>

            <div class="col-lg-3 col-6">
              <div class="stats-item text-center w-100 h-100">
                 <i class="fas fa-check-circle"></i><p><%=itemBody%></p>
              </div>
            </div><!-- End Stats Item -->
              <%
                rsItem.movenext
                Loop
                rsItem.close
              %>
            <%End If%> 
          </div>
        </div>
           <%   
              End If
                rsArticle.movenext
                Loop
                rsArticle.close
            %>
        <div class="col-lg-4 order-1 order-lg-2 hero-img" data-aos="zoom-out">
          <img src="assets/img/hero-img.svg" class="img-fluid mb-3 mb-lg-0" alt="">
        </div>
        </div>
        </div>
      </div>
    </div>
  </section><!-- End Hero Section -->


  <main id="main">

    <!-- Video Action Section -->
    <section>
        <div class="container">
          <div class="row gy-4">
            <%
                set rsVideo = CreateObject("ADODB.recordset")
                  sql = "select * from Article where articleId = 4"
                  rsVideo.open sql, conn
                    articleId = rsVideo("articleId")
                    articleTitle = rsVideo("articleTitle")
                    articleBody = rsVideo("articleBody")
                    interface = rsVideo("interface")
              %>
                <div class="section-header">
                    <span><%=articleTitle%></span>
                    <h2><%=articleTitle%></h2>
                </div>
              <div class="video-action d-flex justify-content-center w-px">
                <%=articleBody%>
              </div>
              <%=rsVideo.close%>
          </div>
        </div>
    </section>

    <!-- ======= Customer Reaction Section ======= -->
    <section id="featured-services" class="featured-services">
      <div class="container">


        <%
                set rsArticle = CreateObject("ADODB.recordset")
                  sql = "select * from Article where articleId = 5"
                  rsArticle.open sql, conn
                    articleId = rsArticle("articleId")
                    articleTitle = rsArticle("articleTitle")
                    articleBody = rsArticle("articleBody")
                    interface = rsArticle("interface")
        %>
        <div class="row gy-4">
          <div class="section-header">
            <span><%=articleTitle%></span>
            <h2><%=articleBody%></h2>
          </div>  
              <%
                set rsItem = CreateObject("ADODB.recordset")
                  sql = "select * from Article_Items where articleId = " & articleId
                  rsItem.open sql, conn
                  Do Until rsItem.eof
                    itembody = rsItem("itemBody")
                    pictureUrl = rsItem("pictureUrl")
                    itemTitle = rsItem("itemTitle")
                %>
              <div class="col-lg-4 col-md-6 service-item" data-aos="fade-up">
                <div class="icon flex-shrink-0 d-flex justify-content-center "><img class="border border-secondary rounded mb-3" width="150px" height="150px" src="<%=pictureUrl%>"/></div>
                <div>
                  <h4 class="title"><%=itemTitle%></h4>
                  <p class="description"><%=itemBody%></p>
                </div>
              </div>
              <%
                rsItem.movenext
                Loop
                rsItem.close
              %>
          <!-- End Service Item -->
        </div>
      </div>
    </section><!-- End Featured Services Section -->


    <!-- ======= Why Action Section ======= -->
    <section id="call-to-action" class="call-to-action">
      <div class="container" data-aos="zoom-out">

        <div class="row justify-content-center">
          <%
            set rsArticle = CreateObject("ADODB.recordset")
                  sql = "select * from Article where interface = 4"
                  rsArticle.open sql, conn
                    articleId = rsArticle("articleId")
                    articleTitle = rsArticle("articleTitle")
                    articleBody = rsArticle("articleBody")
                    interface = rsArticle("interface")
          %>
          <div class="col-lg-8 text-center">
            <h3><%=articleTitle%></h3>
            <p><%=articleBody%></p>
            <a class="cta-btn" href="#">Xem thêm</a>
          </div>
          <%
            rsArticle.close
          %>
        </div>
      </div>
    </section><!-- Why Action Section -->


    <!-- ======= Product Section ======= -->
    <section id="pricing" class="pricing pt-0">
      <div class="container" data-aos="fade-up">

        <%
            set rsArticle = CreateObject("ADODB.recordset")
                  sql = "select * from Article where interface = 5"
                  rsArticle.open sql, conn
                    articleId = rsArticle("articleId")
                    articleTitle = rsArticle("articleTitle")
                    articleBody = rsArticle("articleBody")
                    interface = rsArticle("interface")
          %>

        <div class="section-header">
          <span><%=articleTitle%></span>
          <h2><%=articleBody%></h2>

        </div>

        <div class="row gy-4 d-flex jusify-content-center">
            <%
            set rsItem = CreateObject("ADODB.recordset")
                  sql = "select * from Article_Items where articleId = 7" 
                  rsItem.open sql, conn
                  Do Until rsItem.eof
                    itemBody = rsItem("itemBody")
                    itemTitle = rsItem("itemTitle")
                    itemDescribe = rsItem("itemDescribe")
                    itemPicture = rsItem("pictureUrl")
            %>
          <div class="col-lg-6" data-aos="fade-up" data-aos-delay="100">
            <div class="pricing-item">
              <div>
                <img src="<%=itemPicture%>" alt="<%=itemName%>" class="img-fluid">
              </div>
              <h3><%=itemTitle%></h3>
              <h4><sup>$</sup>0<span> / month</span></h4>
              <p><%=itemBody%></p>
              <a href="#" class="buy-btn">Đăng kí</a>
              <a href="#" class="buy-btn">Chi tiết</a>
            </div>
          </div><!-- End Pricing Item -->
          <%
            rsItem.movenext
            loop
            rsItem.close
            rsArticle.close
          %>
        </div>
      </div>
    </section><!-- End Product Section -->


    <!-- ======= Application Section ======= -->
    <section id="service" class="services pt-0">
      <div class="container" data-aos="fade-up">

        <div class="section-header">
          <span>Ứng dụng của phần mềm</span>
          <h2>Ứng dụng của phần mềm</h2>

        </div>

        <div class="row gy-4">
          <%
                set rsArticle = CreateObject("ADODB.recordset")
                  sql = "select * from Article where interface = 7"
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
          </div><!-- End Card Item -->
          <%
            rsArticle.movenext
            loop
            rsArticle.close
          %>
        </div>

      </div>
    </section><!-- End Services Section -->



    <!-- ======= Frequently Asked Questions Section ======= -->
    <section id="faq" class="faq">
      <div class="container" data-aos="fade-up">
        <%
            set rsArticle = CreateObject("ADODB.recordset")
                  sql = "select * from Article where interface = 8"
                  rsArticle.open sql, conn
                    articleId = rsArticle("articleId")
                    articleTitle = rsArticle("articleTitle")
                    articleBody = rsArticle("articleBody")
                    interface = rsArticle("interface")
            %>
        <div class="section-header">
          <span><%=articleTitle%></span>
          <h2><%=articleBody%></h2>

        </div>

        <div class="row justify-content-center" data-aos="fade-up" data-aos-delay="200">
          <div class="col-lg-10">

            <div class="accordion accordion-flush" id="faqlist">

              <div class="accordion-item">
                <h3 class="accordion-header">
                  <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#faq-content-1">
                    <i class="bi bi-question-circle question-icon"></i>
                    Non consectetur a erat nam at lectus urna duis?
                  </button>
                </h3>
                <div id="faq-content-1" class="accordion-collapse collapse" data-bs-parent="#faqlist">
                  <div class="accordion-body">
                    Feugiat pretium nibh ipsum consequat. Tempus iaculis urna id volutpat lacus laoreet non curabitur gravida. Venenatis lectus magna fringilla urna porttitor rhoncus dolor purus non.
                  </div>
                </div>
              </div><!-- # Faq item-->

              <div class="accordion-item">
                <h3 class="accordion-header">
                  <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#faq-content-2">
                    <i class="bi bi-question-circle question-icon"></i>
                    Feugiat scelerisque varius morbi enim nunc faucibus a pellentesque?
                  </button>
                </h3>
                <div id="faq-content-2" class="accordion-collapse collapse" data-bs-parent="#faqlist">
                  <div class="accordion-body">
                    Dolor sit amet consectetur adipiscing elit pellentesque habitant morbi. Id interdum velit laoreet id donec ultrices. Fringilla phasellus faucibus scelerisque eleifend donec pretium. Est pellentesque elit ullamcorper dignissim. Mauris ultrices eros in cursus turpis massa tincidunt dui.
                  </div>
                </div>
              </div><!-- # Faq item-->

              <div class="accordion-item">
                <h3 class="accordion-header">
                  <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#faq-content-3">
                    <i class="bi bi-question-circle question-icon"></i>
                    Dolor sit amet consectetur adipiscing elit pellentesque habitant morbi?
                  </button>
                </h3>
                <div id="faq-content-3" class="accordion-collapse collapse" data-bs-parent="#faqlist">
                  <div class="accordion-body">
                    Eleifend mi in nulla posuere sollicitudin aliquam ultrices sagittis orci. Faucibus pulvinar elementum integer enim. Sem nulla pharetra diam sit amet nisl suscipit. Rutrum tellus pellentesque eu tincidunt. Lectus urna duis convallis convallis tellus. Urna molestie at elementum eu facilisis sed odio morbi quis
                  </div>
                </div>
              </div><!-- # Faq item-->

              <div class="accordion-item">
                <h3 class="accordion-header">
                  <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#faq-content-4">
                    <i class="bi bi-question-circle question-icon"></i>
                    Ac odio tempor orci dapibus. Aliquam eleifend mi in nulla?
                  </button>
                </h3>
                <div id="faq-content-4" class="accordion-collapse collapse" data-bs-parent="#faqlist">
                  <div class="accordion-body">
                    <i class="bi bi-question-circle question-icon"></i>
                    Dolor sit amet consectetur adipiscing elit pellentesque habitant morbi. Id interdum velit laoreet id donec ultrices. Fringilla phasellus faucibus scelerisque eleifend donec pretium. Est pellentesque elit ullamcorper dignissim. Mauris ultrices eros in cursus turpis massa tincidunt dui.
                  </div>
                </div>
              </div><!-- # Faq item-->

              <div class="accordion-item">
                <h3 class="accordion-header">
                  <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#faq-content-5">
                    <i class="bi bi-question-circle question-icon"></i>
                    Tempus quam pellentesque nec nam aliquam sem et tortor consequat?
                  </button>
                </h3>
                <div id="faq-content-5" class="accordion-collapse collapse" data-bs-parent="#faqlist">
                  <div class="accordion-body">
                    Molestie a iaculis at erat pellentesque adipiscing commodo. Dignissim suspendisse in est ante in. Nunc vel risus commodo viverra maecenas accumsan. Sit amet nisl suscipit adipiscing bibendum est. Purus gravida quis blandit turpis cursus in
                  </div>
                </div>
              </div><!-- # Faq item-->

            </div>

          </div>
        </div>

      </div>
    </section><!-- End Frequently Asked Questions Section -->

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