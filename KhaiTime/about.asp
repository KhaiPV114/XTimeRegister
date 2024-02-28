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
  <!-- ======= Header ======= -->
  <!--#include file="header.asp"-->
  <!-- End Header -->

<body>



  <main id="main">

    <!-- ======= Breadcrumbs ======= -->
    <div class="breadcrumbs">
      <div class="page-header d-flex align-items-center" style="background-image: url('./image/15.png'); padding-top: 80px">
        <div class="container position-relative">
          <div class="row d-flex justify-content-center">
              <%
                    set rsArticle = CreateObject("ADODB.recordset")
                    sql = "select * from Article a join Category_Article ca on a.articleId = ca.articleId  where a.interface = 1 and ca.categoryID = 2"
                    rsArticle.open sql, conn
                      articleId = rsArticle("articleId")
                      articleTitle = rsArticle("articleTitle")
                      articleBody = rsArticle("articleBody")
                      interface = rsArticle("interface")
              %>
            <div class="col-lg-6 text-center">
              <%
                    set rsPicture = CreateObject("ADODB.recordset")
                    sql = "select * from Article_Picture where articleId = " & articleId
                    rsPicture.open sql, conn
                      pictureUrl = rsPicture("pictureUrl")
              %>
              <img src="<%=pictureUrl%>" alt="logo"/>
              <h2><%=articleTitle%></h2>
              <p><%=articleBody%></p>
              <%
                    set rsItem = CreateObject("ADODB.recordset")
                    sql = "select * from Article_Items where articleId = " & articleId
                    rsItem.open sql, conn
                      itemTitle = rsItem("itemTitle")
                      itemBody = rsItem("itemBody")
              %>
              <a class="cta-btn" href="#"><%=itemBody%></a>
              <%
                rsItem.close
                rsPicture.close
                rsArticle.close
              %>
            </div>
            <div>

            </div>  
          </div>
        </div>
      </div>
      <nav>
        <div class="container">
          <ol>
            <li><a href="index.html">Home</a></li>
            <li>About</li>
          </ol>
        </div>
      </nav>
    </div><!-- End Breadcrumbs -->

    <!-- ======= About Us Section ======= -->
    <section id="about" class="about">
      <div class="container" data-aos="fade-up">

        <div class="row gy-4">
          <div class="col-lg-6 position-relative align-self-start order-lg-last order-first">
            <img src="assets/img/about.jpg" class="img-fluid" alt="">
            <a href="https://www.youtube.com/watch?v=LXb3EKWsInQ" class="glightbox play-btn"></a>
          </div>
          <div class="col-lg-6 content order-last  order-lg-first">
            <h3>About Us</h3>
            <p>
              Dolor iure expedita id fuga asperiores qui sunt consequatur minima. Quidem voluptas deleniti. Sit quia molestiae quia quas qui magnam itaque veritatis dolores. Corrupti totam ut eius incidunt reiciendis veritatis asperiores placeat.
            </p>
            <ul>
              <li data-aos="fade-up" data-aos-delay="100">
                <i class="bi bi-diagram-3"></i>
                <div>
                  <h5>Ullamco laboris nisi ut aliquip consequat</h5>
                  <p>Magni facilis facilis repellendus cum excepturi quaerat praesentium libre trade</p>
                </div>
              </li>
              <li data-aos="fade-up" data-aos-delay="200">
                <i class="bi bi-fullscreen-exit"></i>
                <div>
                  <h5>Magnam soluta odio exercitationem reprehenderi</h5>
                  <p>Quo totam dolorum at pariatur aut distinctio dolorum laudantium illo direna pasata redi</p>
                </div>
              </li>
              <li data-aos="fade-up" data-aos-delay="300">
                <i class="bi bi-broadcast"></i>
                <div>
                  <h5>Voluptatem et qui exercitationem</h5>
                  <p>Et velit et eos maiores est tempora et quos dolorem autem tempora incidunt maxime veniam</p>
                </div>
              </li>
            </ul>
          </div>
        </div>

      </div>
    </section><!-- End About Us Section -->


      <!-- ======= VISION AND MISSION Section ======= -->
    <section id="service" class="services pt-0">
      <div class="container" data-aos="fade-up">
          <%
                    set rsArticle = CreateObject("ADODB.recordset")
                    sql = "select * from Article a join Category_Article ca on a.articleId = ca.articleId  where a.interface = 2 and ca.categoryID = 2"
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
                set rsItem = CreateObject("ADODB.recordset")
                  sql = "select * from Article_Items where articleId = " & articleId
                  rsItem.open sql, conn
                  Do Until rsItem.eof
                    itemTitle = rsItem("itemTitle")
                    itemPicture = rsItem("pictureUrl")
                    itemBody = rsItem("itemBody")
              %>
          <div class="col-12 col-md-6" data-aos="fade-up" data-aos-delay="100">
            <div class="border shadow p-3 mb-5 bg-body rounded">
                <div class="icon flex-shrink-0 d-flex justify-content-center "><img class=" rounded m-5" width="256px" height="160px" src="<%=itemPicture%>"/></div>
              
              <h3 class="text-center"><a href="service-details.html" class="stretched-link "><%=itemTitle%></a></h3>
              <p><%=itemBody%></p>
            </div>
          </div><!-- End Card Item -->
          <%
                rsItem.movenext
                loop
                rsItem.close
              %>
          <%
            rsArticle.movenext
            rsArticle.close
          %>
        </div>

      </div>
    </section><!-- End Services Section -->


    <!-- Technology Section -->
    <section style="background-color: #888e97; background-image: url('./image/11.png')">
        <div class="container" >
          <div class="row gy-4">
            <%
                set rsArticle = CreateObject("ADODB.recordset")
                    sql = "select * from Article a join Category_Article ca on a.articleId = ca.articleId  where a.interface = 3 and ca.categoryID = 2"
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
              <div>
                  <p style="text-align:center"><%=articleBody%></p>
                  
                  <div class="icon flex-shrink-0 ">
                    <div class="row " style="justify-content: space-evenly !important ;">
                      <%
                        set rsPicture = CreateObject("ADODB.recordset")
                            sql = "select * from Article_Picture where articleId = 19"
                            rsPicture.open sql, conn
                              Do until rsPicture.eof
                                pictureUrl = rsPicture("pictureUrl")
                      %>
                          
                                  <div class="col-md-3 col-12 py-3">
                                      <div>
                                          <a href="">
                                              <img src="<%=pictureUrl%>" style="aspect-ratio: 1/1;" alt=" ">
                                          </a>
                                      </div>
                                  </div>
                      
                      <%
                        rsPicture.movenext
                        loop
                        rsPicture.close
                        rsArticle.close
                      %>
                    </div>
                  </div>
              </div>
            </div>
          </div>
          </div>
        </div>
    </section>
     <!-- End Technology Section -->



    <!-- ======= Stats Counter Section ======= -->
    <section id="stats-counter" class="stats-counter pt-0">
      <div class="container" data-aos="fade-up">

        <div class="row gy-4">

          <div class="col-lg-3 col-md-6">
            <div class="stats-item text-center w-100 h-100">
              <span data-purecounter-start="0" data-purecounter-end="232" data-purecounter-duration="1" class="purecounter"></span>
              <p>Clients</p>
            </div>
          </div><!-- End Stats Item -->

          <div class="col-lg-3 col-md-6">
            <div class="stats-item text-center w-100 h-100">
              <span data-purecounter-start="0" data-purecounter-end="521" data-purecounter-duration="1" class="purecounter"></span>
              <p>Projects</p>
            </div>
          </div><!-- End Stats Item -->

          <div class="col-lg-3 col-md-6">
            <div class="stats-item text-center w-100 h-100">
              <span data-purecounter-start="0" data-purecounter-end="1453" data-purecounter-duration="1" class="purecounter"></span>
              <p>Hours Of Support</p>
            </div>
          </div><!-- End Stats Item -->

          <div class="col-lg-3 col-md-6">
            <div class="stats-item text-center w-100 h-100">
              <span data-purecounter-start="0" data-purecounter-end="32" data-purecounter-duration="1" class="purecounter"></span>
              <p>Workers</p>
            </div>
          </div><!-- End Stats Item -->

        </div>

      </div>
    </section><!-- End Stats Counter Section -->

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