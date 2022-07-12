<!Doctype html>
<html>
    <head>
        <meta charset="utf-8" />
        <link rel="shortcut icon" href="plugins/images/LOGOYKKBI.png" type="image/x-icon">
        <title>Excel</title>
        <style type="text/css">
            html, body {
                width: 100%;
                height: 100%;
                padding: 0px;
                margin: 0px;
            }
        </style>
        <link rel='stylesheet' href='css/pluginsCss.css' />
        <link rel='stylesheet' href='css/plugins.css' />
        <link rel='stylesheet' href='css/luckysheet.css' />
        <link rel='stylesheet' href='css/iconfont.css' />
        
    </head>
    <body>
        <div id="luckysheet" style="margin:0px;padding:0px;position:absolute;width:100%;height:100%;left: 0px;top: 0px;"></div>
        <script src="js/plugin.js"></script>
        <script src="js/luckysheet.umd.js"></script>
        <script>
            $(function () {
                //Configuration item
                var options = {
                    container: 'luckysheet' //luckysheet is the container id
                }
                luckysheet.create(options)
            })
        </script>
    </body>
</html>