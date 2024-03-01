$computer = $env:COMPUTERNAME
$reportCompanyName = "CannIT"
$runby = $env:USERNAME
$version = "v1.0"
$dt1    = (Get-Date -Format 'yyyy/MM/dd HH:mm')    
[string]$un     = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name.ToLower()


$path = "output.htm"
$html = @'
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="en">
<head>
    <meta charset="utf-8">
    <title>QA Report</title>
    <style>
        @charset UTF-8;
        html body { font-family: Segoe UI, Verdana, Geneva, sans-serif; font-size: 12px; height: 100%; overflow: auto; color: #000000; }
        .header1  { width: 99%; margin: 0px 10px 0px auto; }
        .header2  { width: 99%; margin: 0px 10px 0px auto; padding-top: 10px; clear: both; min-height: 80px; }

        .header1 > .headerCompany { float: left;  font-size: 333%; font-weight: bold; }
        .header1 > .headerQA      { float: left;  font-size: 333%; }
        .header1 > .headerDetails { float: right; font-size: 100%; text-align: right;  }
        .header1 > .headerDetails > .item { display:block; padding: 0 0 3px 0; }

        

        .header2 > .headerServer { float: left; font-weight: normal; height: 88px; }
        .header2 > .headerServer > .serverName { font-size: 266%; line-height: 35px; text-transform: uppercase; }
        .header2 > .headerServer > .row { font-size: 100%; padding-left: 3px; padding-top: 2px; }

        /*  Size of check code/num boxes :                  (6 x (100 + 12)) + 10  = 794px  */
        /*  Slightly larger boxes        :                  (6 x (110 + 12)) + 10  = 864px  */
        /*  Just a bit bigger boxes      :                  (6 x (115 + 12)) + 10  = 899px  */

        .header2 > .summary { float:right; background: #f8f8f8; height: 77px; width: 864px; padding-top: 10px; border-right: 1px solid #ccc; border-bottom: 1px solid #cccccc; }
        .header2 > .summary > a > .summaryBox { float: left; height: 65px; width: 110px; text-align: center; margin-left: 10px; padding: 0px; border: 1px solid #000; }
        .header2 > .summary > a > .summaryBox > .code { font-size: 133%; padding-top: 5px; display: block; font-weight: bold; }
        .header2 > .summary > a > .summaryBox > .num  { font-size: 233%; }

        .sectionTitle    { padding: 5px; font-size: 233%; text-align: center; letter-spacing: 3px; display: block; }
        .sectionItem     { background: #707070; color: #ffffff; width: 99%; display: block; margin: 25px auto  5px auto; padding: 0; overflow: auto; }
        .checkItem       { background: #f8f8f8;                 width: 99%; display: block; margin: 10px auto 10px auto; padding: 0; overflow: auto; border-right: 1px solid #cccccc; border-bottom: 1px solid #cccccc; }
        .checkItem:hover { background: #f2f2f2; }

        .boxContainer { float: left; width: 80px; height: 77px; }
        .boxContainer > .check { position: relative; top: 0; left: 0; height: 65px; width: 100px; text-align: center; margin: 5px 0px 5px 5px; padding: 0px; border: 1px solid #707070; background: #ff00ff; cursor: default; }
        .boxContainer > .check > .code { font-size: 133%; padding-top: 5px; font-weight: bold; display: block; }
        .boxContainer > .check > .num { font-size: 233%; }

        .contentContainer { margin-left: 100px; padding: 10px 10px 10px 15px; overflow: auto; }
        .checkContainer  { float: left; width: 45%; }
        .checkContainer  > .name    { font-size: 125%; margin: 0 0 5px 0; font-weight: bold; }
        .checkContainer  > .message { font-size: 110%; }
        .resultContainer { float: left; width: 50%; }
        .resultContainer > .data > .dataHeader { font-weight: bold; margin-bottom: 5px; }
       
        .arrow          { border-right: 7px solid #000000; border-bottom: 7px solid #000000; width: 10px; height: 10px; transform: rotate(-135deg); margin-top: 5px; }
        .btt            { color: #000000; background: #ffffff; font-size: 125%; border: 1px solid #707070; margin: 0px; padding: 12px 15px; font-weight: bold; display: block; right: 10px; position: fixed; text-align: center; text-decoration: none; bottom: 10px; z-index: 100; border-radius: 50px; }
        .tocEntry       { color: #000000; background: #f8f8f8; font-size: 125%; border: 1px solid #707070; margin: 2px; padding:  5px 10px; font-weight: bold; }
        .btt:hover      { color: #ffffff; background: #707070; border: 1px solid #000000; }
        .tocEntry:hover { color: #ffffff; background: #707070; border: 1px solid #000000; }
        a               { color: inherit; text-decoration: none; }

        .note                { text-decoration: none; }
        .note div.help       { display: none; }
        .note:hover          { cursor: help; position: relative; }
        .note:hover div.help { color: #000000; background: #ffffdd; border: #000000 3px solid; display: block; right: 10px; margin: 10px; padding: 15px; position: fixed; text-align: left; text-decoration: none; top: 10px; width: 600px; z-index: 100; }
        .note li             { display: table-row-group; list-style: none; }
        .note li span        { display: table-cell; vertical-align: top; padding: 3px 0; }
        .note li span:first-child { text-align: right; min-width: 120px; max-width: 120px; font-weight: bold; padding-right: 7px; }
        .note li span:last-child  { padding-left: 7px; border-left: 1px solid #000000; }

        .x  { background: #ffffff !important; }
        .p  { background: #b3ffb3 !important; }
        .w  { background: #ffffb3 !important; }
        .f  { background: #ffb3b3 !important; }
        .m  { background: #b3b3ff !important; }
        .n  { background: #e2e2e2 !important; }
        .e  { background: #c80000 !important; color: #ffffff !important; }
        .eB { background: #c80000 !important; color: #ffffff !important; border: 1px solid #ffffff !important; }
    </style>

    <script>
        function showSwitch(sId) {
            var d = document.getElementsByTagName("div");
            for (var i = 0; i < d.length; i++) { if (d[i].getAttribute("id") == sId) { d[i].style.display = 'block'; } }
        }

        function hideSwitch(sId) {
            var d = document.getElementsByTagName("div");
            for (var i = 0; i < d.length; i++) { if (d[i].getAttribute("id") == sId) { d[i].style.display = 'none';  } }
        }

        function showall() { showSwitch("p"); showSwitch("w"); showSwitch("f"); showSwitch("m"); showSwitch("n"); showSwitch("eB"); }
        function sh_pass() { showSwitch("p"); hideSwitch("w"); hideSwitch("f"); hideSwitch("m"); hideSwitch("n"); hideSwitch("eB"); }
        function sh_warn() { hideSwitch("p"); showSwitch("w"); hideSwitch("f"); hideSwitch("m"); hideSwitch("n"); hideSwitch("eB"); }
        function sh_fail() { hideSwitch("p"); hideSwitch("w"); showSwitch("f"); hideSwitch("m"); hideSwitch("n"); hideSwitch("eB"); }
        function sh_manu() { hideSwitch("p"); hideSwitch("w"); hideSwitch("f"); showSwitch("m"); hideSwitch("n"); hideSwitch("eB"); }
        function sh_nota() { hideSwitch("p"); hideSwitch("w"); hideSwitch("f"); hideSwitch("m"); showSwitch("n"); hideSwitch("eB"); }
        function sh_erro() { hideSwitch("p"); hideSwitch("w"); hideSwitch("f"); hideSwitch("m"); hideSwitch("n"); showSwitch("eB"); }
    </script>
</head>
<body>
BODY_GOES_HERE
</body>
</html>
'@

[System.Text.StringBuilder]$body = @"
<a href="#BackToTop" title="Jump to top of page"><div class="btt"><div class="arrow"></div></div></a>
    <div id="BackToTop" class="header1">
        <span class="headerCompany">$reportCompanyName</span>
        <span class="headerQA"     >&nbsp;$($script:lang['QA-Results'])</span>
        <div class="headerDetails">
            <span class="item">$($script:lang['ScriptVersion']) <strong>$version      </strong></span>
            <span class="item">$($script:lang['Configuration']) <strong>$settingsFile </strong></span>
            <span class="item">$($script:lang['GeneratedOn']  ) <strong>$dt1          </strong></span>
            <span class="item">$($script:lang['GeneratedBy']  ) <strong>$un           </strong></span>
        </div>
    </div>

    <div class="header2">
        <div class="headerServer">
            <div class="serverName">$($server)</div>
            <div class="row"       >$($ResultsInput[0].Data.Split('|')[0])</div>
            <div class="row"><b    >$($ResultsInput[0].Data.Split('|')[1])</b></div>
            <div class="row"       >$($ResultsInput[0].Data.Split('|')[2]),&nbsp;&nbsp;&nbsp;&nbsp;$($ResultsInput[0].Data.Split('|')[3])</div>
        </div>
        <div class="summary">
            <a href="#" onclick="showall();"><div class="summaryBox x"><span class="code">$($script:lang['ShowAll']       )</span><span class="num">$($ResultsInput.Count - 1)</span></div></a>
            <a href="#" onclick="sh_pass();"><div class="summaryBox p"><span class="code">$($script:lang['Pass']          )</span><span class="num">$($resultsplit.p         )</span></div></a>
            <a href="#" onclick="sh_warn();"><div class="summaryBox w"><span class="code">$($script:lang['Warning']       )</span><span class="num">$($resultsplit.w         )</span></div></a>
            <a href="#" onclick="sh_fail();"><div class="summaryBox f"><span class="code">$($script:lang['Fail']          )</span><span class="num">$($resultsplit.f         )</span></div></a>
            <a href="#" onclick="sh_manu();"><div class="summaryBox m"><span class="code">$($script:lang['Manual']        )</span><span class="num">$($resultsplit.m         )</span></div></a>
            <a href="#" onclick="sh_nota();"><div class="summaryBox n"><span class="code">$($script:lang['Not-Applicable'])</span><span class="num">$($resultsplit.n         )</span></div></a>
            <a href="#" onclick="sh_erro();"><div class="summaryBox e"><span class="code">$($script:lang['Error']         )</span><span class="num">$($resultsplit.e         )</span></div></a>
        </div>
    </div>
    <div style="clear:both;"></div>

    <div class="sectionItem"><span class="sectionTitle">Jump Links To Sections</span></div>
    <div class="checkItem"><div style="text-align: center;"><br/>
SECTION_LINKS
    <br/><br/></div></div>
    <div style="clear:both;"></div>

"@


$body = get-service A* | ft

$body

$html = $html.Replace('BODY_GOES_HERE', $body.ToString())

#$htmlParams = @{
#  Title = "Windows Services: $computer"
#  Body = Get-Date
#  PreContent = "<P>Generated by $runby</P>"
#  PostContent = "For details, contact Corporate IT."
#}
#Get-Service A* |
#ConvertTo-Html @html |
#    Out-File Services.htm

$html | Out-File $path -Force -Encoding utf8

Invoke-Item output.htm
