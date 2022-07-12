<!Doctype html>
<html>
    <head>
        <meta charset='utf-8' />
        <title>Download</title>
        <style type="text/css">
            html, body {
                font-family: consolas, monospace;
            }
            input, textarea, select {
                font-family: consolas, monospace;
                box-sizing: border-box;
            }
        </style>
        <script src="./js/download.js"></script>
        <script type="text/javascript">
            /* download('./excel/phpexcel-output/HARDEXAMPLE01.xlsx','COBA01.xlsx'); */
        </script>
    </head>
    <body>
        <input type="button" value="Download" onclick="download('./excel/phpexcel-output/HARDEXAMPLE02.xlsx','COBA02.xlsx');" /><br />
        <input type="button" value="Download" onclick="download('./excel/phpexcel-output/HARDEXAMPLE02.xlsx','COBA02.xlsx');" /><br /><!-- 
        <a href="data:application/xml;charset=utf-8," download="filename.html">Save</a> -->
    </body>
</html>