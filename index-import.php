<!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8" />
        <link rel="shortcut icon" href="plugins/images/LOGOYKKBI.png" type="image/x-icon">
        <title>Excel</title>
        <link rel='stylesheet' href='./luckysheet/plugins/css/pluginsCss.css' />
        <link rel='stylesheet' href='./luckysheet/plugins/plugins.css' />
        <link rel='stylesheet' href='./luckysheet/css/luckysheet.css' />
        <link rel='stylesheet' href='./luckysheet/assets/iconfont/iconfont.css' />
        <script src="./luckysheet/plugins/js/plugin.js"></script>
        <script src="./luckysheet/luckysheet.umd.js"></script>
        <script>
            $(function () {
                //Configuration item
                var options = {
                    container: 'luckysheet', //luckysheet is the container id
                    showinfobar: false,
                }
                luckysheet.create(options)
            });
        </script>
        <style type="text/css">
            html, body {
                font-family: consolas, monospace;
            }
            input, textarea, select {
                font-family: consolas, monospace;
            }
        </style>
    </head>
    <body>
        <div id="lucky-mask-demo" style="position: absolute;z-index: 1000000;left: 0px;top: 0px;bottom: 0px;right: 0px; background: rgba(255, 255, 255, 0.8); text-align: center;font-size: 40px;align-items:center;justify-content: center;display: none;">Downloading</div>
        <div id="lucky-mask-export" style="position: absolute;z-index: 1000000;left: 0px;top: 0px;bottom: 0px;right: 0px; background: rgba(255, 255, 255, 0.8); text-align: center;font-size: 40px;align-items:center;justify-content: center;display: none;">Exporting</div>
        <p style="text-align:center; font-family: consolas, monospace;"> 
            <input style="font-size:16px;" type="file" id="Luckyexcel-demo-file" name="Luckyexcel-demo-file" change="demoHandler" /> 
            Or Load remote xlsx file: 
            <select style="height: 27px;top: -2px;position: relative; font-family: consolas, monospace; box-sizing: border-box;" id="Luckyexcel-select-demo"> 
                <option value="">select a demo</option> 
                <option value="https://minio.cnbabylon.com/public/luckysheet/money-manager-2.xlsx">Money Manager.xlsx</option> 
                <option value="https://minio.cnbabylon.com/public/luckysheet/Activity%20costs%20tracker.xlsx">Activity costs tracker.xlsx</option>
                <option value="https://minio.cnbabylon.com/public/luckysheet/House%20cleaning%20checklist.xlsx">House cleaning checklist.xlsx</option>
                <option value="https://minio.cnbabylon.com/public/luckysheet/Student%20assignment%20planner.xlsx">Student assignment planner.xlsx</option>
                <option value="https://minio.cnbabylon.com/public/luckysheet/Credit%20card%20tracker.xlsx">Credit card tracker.xlsx</option>
                <option value="https://minio.cnbabylon.com/public/luckysheet/Blue%20timesheet.xlsx">Blue timesheet.xlsx</option>
                <option value="https://minio.cnbabylon.com/public/luckysheet/Student%20calendar%20%28Mon%29.xlsx">Student calendar (Mon).xlsx</option>
                <option value="https://minio.cnbabylon.com/public/luckysheet/Blue%20mileage%20and%20expense%20report.xlsx">Blue mileage and expense report.xlsx</option> 
            </select> 
            <a href="javascript:void(0)" id="Luckyexcel-downlod-file">Download</a>
            <a href="javascript:void(0)" id="Luckyexcel-export-xlsx">Export</a>
        </p>
        <div id="luckysheet" style="margin:0px;padding:0px;position:absolute;width:100%;left: 0px;top: 50px;bottom: 0px;outline: none;"></div>
        <script src="./luckysheet/luckyexcel.umd.js"></script>
        <!-- 
        <script type="module">/*
            import l from './luckyexcel.js';
            console.info('=====',l)
            window.onload = () => {
                let upload = document.getElementById("file");
                upload.addEventListener("change", function(evt){
                    var files = evt.target.files;   
                    importFile(files[0]);
                });
            } */
        </script> -->
        <script>
            function demoHandler() {
                let upload = document.getElementById("Luckyexcel-demo-file");
                let selectADemo = document.getElementById("Luckyexcel-select-demo");
                let downlodDemo = document.getElementById("Luckyexcel-downlod-file");
                let export_ = document.getElementById("Luckyexcel-export-xlsx");
                let mask = document.getElementById("lucky-mask-demo");
                let export_mask = document.getElementById("lucky-mask-export");
                if (upload) {

                    window.onload = () => {

                        upload.addEventListener("change", function (evt) {
                            var files = evt.target.files;
                            if (files == null || files.length == 0) {
                                alert("No files wait for import");
                                return;
                            }

                            let name = files[0].name;
                            let suffixArr = name.split("."), suffix = suffixArr[suffixArr.length - 1];
                            if (suffix != "xlsx") {
                                alert("Currently only supports the import of xlsx files");
                                return;
                            }
                            LuckyExcel.transformExcelToLucky(files[0], function (exportJson, luckysheetfile) {

                                if (exportJson.sheets == null || exportJson.sheets.length == 0) {
                                    alert("Failed to read the content of the excel file, currently does not support xls files!");
                                    return;
                                }
                                window.luckysheet.destroy();
                                console.log(exportJson.sheets);
                                window.luckysheet.create({
                                    container: 'luckysheet', //luckysheet is the container id
                                    showinfobar: false,
                                    data: exportJson.sheets,
                                    title: exportJson.info.name,
                                    userInfo: exportJson.info.name.creator
                                });
                            });
                        });

                        selectADemo.addEventListener("change", function (evt) {
                            var obj = selectADemo;
                            var index = obj.selectedIndex;
                            var value = obj.options[index].value;
                            var name = obj.options[index].innerHTML;
                            if (value == "") {
                                return;
                            }
                            mask.style.display = "flex";
                            LuckyExcel.transformExcelToLuckyByUrl(value, name, function (exportJson, luckysheetfile) {

                                if (exportJson.sheets == null || exportJson.sheets.length == 0) {
                                    alert("Failed to read the content of the excel file, currently does not support xls files!");
                                    return;
                                }
                                // console.log(exportJson, luckysheetfile);
                                mask.style.display = "none";
                                window.luckysheet.destroy();
                                console.log(exportJson.sheets);
                                window.luckysheet.create({
                                    container: 'luckysheet', //luckysheet is the container id
                                    showinfobar: false,
                                    data: exportJson.sheets,
                                    title: exportJson.info.name,
                                    userInfo: exportJson.info.name.creator
                                });
                            });
                        });
                        
                        export_.addEventListener("click", function () {
                            export_mask.style.display = "flex";
                            var isi_excel = window.luckysheet.getAllSheets();
                            console.log(this);
                            console.log(isi_excel);/*
                            var json_upload = "json_luckyexcel=" + JSON.stringify(isi_excel);
                            // console.log(json_upload);
                            var xmlhttp = new XMLHttpRequest(); 
                            xmlhttp.onreadystatechange = function(){
                                if(this.status === 200 && this.readyState === 4){
                                    export_mask.style.display = "none";
                                }
                            };
                            xmlhttp.open("POST", "excel/cobaexcel-luckyexport.php");
                            xmlhttp.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                            xmlhttp.send(json_upload); */
                            $.ajax({
                                type:'post',
                                url:'excel/cobaexcel-luckyexport.php',
                                data: {json_luckyexcel:isi_excel},
                                success: function(response) {
                                    response;
                                    console.log('SUCCESS BLOCK');
                                    export_mask.style.display = "none";
                                },
                                error: function(response) {
                                    console.log('ERROR BLOCK');
                                    console.log(response);
                                }
                            });
                        });
                        
                        downlodDemo.addEventListener("click", function () {
                            var obj = selectADemo;
                            var index = obj.selectedIndex;
                            var value = obj.options[index].value;

                            if (value.length == 0) {
                                alert("Please select a demo file");
                                return;
                            }

                            var elemIF = document.getElementById("Lucky-download-frame");
                            if (elemIF == null) {
                                elemIF = document.createElement("iframe");
                                elemIF.style.display = "none";
                                elemIF.id = "Lucky-download-frame";
                                document.body.appendChild(elemIF);
                            }
                            elemIF.src = value;
                        });
                    }
                }
            }
            demoHandler();
        </script>
    </body>
</html>