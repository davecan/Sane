
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">

    <title>Northwind Intranet</title>

    <!-- Latest compiled and minified CSS -->
    <!--<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/css/bootstrap.min.css" integrity="sha384-1q8mTJOASx8j1Au+a5WDVnPi2lkFfwwEAa8hDDdjZlpLegxhjVME1fgjWPGmkzs7" crossorigin="anonymous">-->
    <link rel="stylesheet" crossorigin="anonymous" href="https://bootswatch.com/cerulean/bootstrap.min.css" />

    <style type="text/css">
        th, td { font-size: 15px !important; }

        input[type=checkbox].form-control { height: 20px !important; display: inline; }
    </style>
  </head>

  <body id="MVC-<%= H(MVC.ControllerName) & "-" & H(MVC.ActionName) %>">

    <nav class="navbar navbar-default navbar-fixed-top">
      <div class="container">
        <div class="navbar-header">
          <button type="button" class="navbar-toggle collapsed" data-toggle="collapse" data-target="#navbar" aria-expanded="false" aria-controls="navbar">
            <span class="sr-only">Toggle navigation</span>
            <span class="icon-bar"></span>
            <span class="icon-bar"></span>
            <span class="icon-bar"></span>
          </button>
          <%= Html.LinkToExt("Northwind Intranet", "Home", "Index", empty, Array("class", "navbar-brand")) %>
        </div>
        <div id="navbar" class="collapse navbar-collapse">
          <ul class="nav navbar-nav">
            <li><%= Html.LinkTo("Products", "Products", "Index") %></li>
            <li><%= Html.LinkTo("Orders", "Orders", "Index") %></li>
          </ul>
        </div><!--/.nav-collapse -->
      </div>
    </nav>

    <div class="container" style="margin-top: 50px">

        <%
        Flash.ShowErrorsIfPresent
        Flash.ShowSuccessIfPresent
        %>