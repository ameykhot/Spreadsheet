
<!-- -*- mode: html; -*- -->
<!DOCTYPE html>
<html>
  <head>
    <title>CS 444/544 Spreadsheet Selection</title>
    <link rel="stylesheet" href="/style.css">
  </head>
  <body>
    <h1>CS 444/544 Spreadsheet Selection</h1>
    <p class="error"></p>
    <form action="/" method="POST">
      <label for="ssName">Open Spreadsheet Name</label>
      <span>
	<input name="ssName" value="" id="ssName"/>
	<br/>
	<span class="error" id="ssNameError">{{msg}}</span>
      </span>
      <input type="submit" value="Open"/>
    </form>
  </body>
</html>
