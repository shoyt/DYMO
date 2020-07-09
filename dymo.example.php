<?php
if (isset($_REQUEST['viewSource']))
{
  $source = file_get_contents(__FILE__);
  $source = preg_replace('/<\?.*?\?\>/s', '', $source);
  echo '<pre style="font:11px courier new; text-align:left;">' . htmlentities($source) . '</pre>';  
  exit;
}
$pageTitle            = 'DYMO Printers';
$record               = []; // tie to DB and pull a recordset
$gageId               = $record['GAGE'            ];           
$serialNumber         = $record['SERIAL'          ];     
$model                = $record['MODEL'           ];      
$calibratedOn         = $record['CAL_DATE'        ];       
$inServiceOn          = $record['DATE_IN_SERVICE' ];
$nextDueOn            = $record['NEXT_DUE'        ];       
$boldFont             = isset($_REQUEST['boldFont'])            ? $_REQUEST['boldFont']             : true;
$includeSerialNumber  = isset($_REQUEST['includeSerialNumber']) ? $_REQUEST['includeSerialNumber']  : true;
$includeModel         = isset($_REQUEST['includeModel'])        ? $_REQUEST['includeModel']         : true;
?>
<!doctype html>
<head>
  <meta http-equiv="content-type" content="text/html; charset=utf-8" />
  <meta http-equiv="x-ua-compatible" content="ie=edge" />
  <title><?= $pageTitle ?></title>
  <style type="text/css">
    body
    {
      font        : 9pt Arial;
      margin      : 30px;
      margin-left : 50px;
    }
    .configuration
    {
      display     : inline-block;
      font-size   : 9pt;
    }
    .label        ,
    label     
    {
      display     : inline-block;
      white-space : nowrap;
      width       : 165px;
    }
  </style>
  <script type="text/javascript" charset="utf-8" src="javascript/dymo.label.framework.js"> </script>
  <script type="text/javascript" charset="utf-8">
    var configuration = {
                          backwardCompatibility : {
                                                    activeX   : { interface : ""                , isSupported : false } ,
                                                    dymoAddIn : { interface : "Dymo.DymoAddIn"  , isSupported : false } ,
                                                    dymoTape  : { interface : "Dymo.DymoTape"   , isSupported : false }
                                                  },
                          defaultPrinter        : "DYMO LabelWriter 450D",
                          defaultTapeSize       : "19mm",
                          depricatedPrinters    : [],
                          display               : function()
                                                  {
                                                    var compatibility = document.getElementById("compatibility");
                                                    var printer       = getCookie("printer");
                                                    var boldFont      = document.getElementById("boldFont");
                                                    boldFont.value    = getCookie("boldFont");
                                                    setIndexByValue("tapeSize", getCookie("tapeSize"));
                                                    setOptions();
                                                    try
                                                    {
                                                      dymo.label.framework.getPrintersAsync().then
                                                                                              (
                                                                                                function (printers)
                                                                                                {
                                                                                                  var printerList = document.getElementById("printer");
                                                                                                  for (var i = 0, length = printers.length; i < length; i++)
                                                                                                  {
                                                                                                    var printerName = printers[i].name;
                                                                                                    printerList.options.add(new Option(printerName, printerName));
                                                                                                  }
                                                                                                  printer                   = printerName ? printerName : printer;
                                                                                                  setIndexByValue("printer", printer);
                                                                                                  var environment           = dymo.label.framework.checkEnvironment();
                                                                                                  compatibility.innerHTML  += "<span class='label'>BROWSER SUPPORTED</span><span    style='font-weight:bold;'>" + (environment.isBrowserSupported    ? "<span style='color:green;'>&check;</span>" : "<span style='color:red;'>NO</span>") + "</span><br>" +
                                                                                                                              "<span class='label'>FRAMEWORK INSTALLED</span><span  style='font-weight:bold;'>" + (environment.isFrameworkInstalled  ? "<span style='color:green;'>&check;</span>" : "<span style='color:red;'>NO</span>") + "</span><br>" +
                                                                                                                              "<span class='label'>WEB SERVICE RUNNING</span><span  style='font-weight:bold;'>" + (environment.isWebServicePresent   ? "<span style='color:green;'>&check;</span>" : "<span style='color:red;'>NO</span>") + "</span><br>" +
                                                                                                                              "<span class='label'>ERRORS</span><span               style='font-weight:bold;'>" + (environment.errorDetails          ? "<span style='color:green;'>&check;</span>" : "<span style='color:red;'>NO</span>") + "</span><br>";
                                                                                                  compatibility.style.cursor = "";
                                                                                                }
                                                                                              );
                                                    } catch (exception) {
                                                      compatibility.innerHTML    += "<span class='label'>ERRORS</span><span style='color:black;'>" + exception.message + "</span><br>";
                                                      compatibility.style.cursor  = "";
                                                    }
                                                  },
                          initialize            : function()
                                                  {
                                                    for (var objectType in this.backwardCompatibility)
                                                    {
                                                      if (!this.backwardCompatibility.hasOwnProperty(objectType)){ continue; }
                                                      try
                                                      {
                                                        new ActiveXObject(this.backwardCompatibility[objectType].interface);
                                                        this.backwardCompatibility[objectType].isSupported = true;
                                                      } catch (e) {
                                                        this.backwardCompatibility[objectType].isSupported = objectType == "activeX" && e.name == "TypeError";
                                                      }
                                                    }
                                                    if (this.backwardCompatibility.dymoAddIn.isSupported)
                                                    {
                                                      var addIn       = new ActiveXObject(this.backwardCompatibility.dymoAddIn.interface);
                                                      var printers    = addIn.GetDymoPrinters().split("|");
                                                      for (var i = 0, length = printers.length; i < length; i++)
                                                      {
                                                        if (printers[i] == ""){ continue; }
                                                        this.depricatedPrinters[this.depricatedPrinters.length] = printers[i];
                                                      }
                                                      if (this.depricatedPrinters.length == 0){ this.depricatedPrinters[this.depricatedPrinters.length] = this.defaultPrinter; }
                                                    }
                                                  },
                          minimumTapeSize       : "12mm"
                        };
    if (typeof Array.includes !== 'function')
    {
      Array.prototype.includes = function(value, start)
                                 {
                                   start = !start || start < 0 ? 0 : start;
                                   for (var i = start, length = this.length; i < length; i++)
                                   {
                                     if (this[i] == value){ return true; }
                                   }
                                   return false;
                                 }
    }
    switch (true)
    {
      case window.addEventListener  : window.addEventListener ("load"   , initialize, false)  ; break;
      case window.attachEvent       : window.attachEvent      ("onload" , initialize)         ; break;
      default                       : window.onload = initialize                              ; break;
    }
    function display(node)
    {
      node                = document.getElementById(node);
      node.style.display  = node.style.display == "none" ? "" : "none";
    }
    function getCookie(name)
    {
      name               += "=";
      var decodedCookie   = decodeURIComponent(document.cookie);
      var crumbs          = decodedCookie.split(";");
      for (var i = 0; i < crumbs.length; i++)
      {
        var crumb = crumbs[i];
        while (crumb.charAt(0) == " "){ crumb = crumb.substring(1); }
        if (crumb.indexOf(name) == 0){ return crumb.substring(name.length, crumb.length); }
      }
      return "";
    }
    function getXmlFormat()
    {
      return '<' + '?xml version="1.0" encoding="utf-8"?' + '>                                                                    ' + "\r\n" + '\
<ContinuousLabel Version="8.0" Units="twips">                                                                                     ' + "\r\n" + '\
  <PaperOrientation>Landscape</PaperOrientation>                                                                                  ' + "\r\n" + '\
  <Id>[[[LABEL_NAME]]]</Id>                                                                                                       ' + "\r\n" + '\
  <PaperName>[[[PAPER_NAME]]]</PaperName>                                                                                         ' + "\r\n" + '\
  <LengthMode>Auto</LengthMode>                                                                                                   ' + "\r\n" + '\
  <LabelLength>1200</LabelLength>                                                                                                 ' + "\r\n" + '\
  <RootCell>                                                                                                                      ' + "\r\n" + '\
    <Length>7200</Length>                                                                                                         ' + "\r\n" + '\
    <LengthMode>Auto</LengthMode>                                                                                                 ' + "\r\n" + '\
    <BorderWidth>0</BorderWidth>                                                                                                  ' + "\r\n" + '\
    <BorderStyle>Solid</BorderStyle>                                                                                              ' + "\r\n" + '\
    <BorderColor Alpha="255" Red="0" Green="0" Blue="0" />                                                                        ' + "\r\n" + '\
    <SubcellsOrientation>Horizontal</SubcellsOrientation>                                                                         ' + "\r\n" + '\
    <Subcells>                                                                                                                    ' + "\r\n" + '\
      <Cell>                                                                                                                      ' + "\r\n" + '\
        <TextObject>                                                                                                              ' + "\r\n" + '\
          <Name>Text</Name>                                                                                                       ' + "\r\n" + '\
          <ForeColor Alpha="255" Red="0" Green="0" Blue="0" />                                                                    ' + "\r\n" + '\
          <BackColor Alpha="0" Red="255" Green="255" Blue="255" />                                                                ' + "\r\n" + '\
          <LinkedObjectName />                                                                                                    ' + "\r\n" + '\
          <Rotation>Rotation0</Rotation>                                                                                          ' + "\r\n" + '\
          <IsMirrored>False</IsMirrored>                                                                                          ' + "\r\n" + '\
          <IsVariable>True</IsVariable>                                                                                           ' + "\r\n" + '\
          <GroupID>-1</GroupID>                                                                                                   ' + "\r\n" + '\
          <IsOutlined>False</IsOutlined>                                                                                          ' + "\r\n" + '\
          <HorizontalAlignment>Left</HorizontalAlignment>                                                                         ' + "\r\n" + '\
          <VerticalAlignment>Middle</VerticalAlignment>                                                                           ' + "\r\n" + '\
          <TextFitMode>AlwaysFit</TextFitMode>                                                                                    ' + "\r\n" + '\
          <UseFullFontHeight>True</UseFullFontHeight>                                                                             ' + "\r\n" + '\
          <Verticalized>False</Verticalized>                                                                                      ' + "\r\n" + '\
          <StyledText>                                                                                                            ' + "\r\n" + '\
            <Element>                                                                                                             ' + "\r\n" + '\
              <String xml:space="preserve">[[[OUTPUT]]]</String>                                                                  ' + "\r\n" + '\
              <Attributes>                                                                                                        ' + "\r\n" + '\
                <Font Family="Courier New" Size="12" Bold="[[[BOLD_FONT]]]" Italic="False" Underline="False" Strikeout="False" /> ' + "\r\n" + '\
                <ForeColor Alpha="255" Red="0" Green="0" Blue="0" HueScale="100" />                                               ' + "\r\n" + '\
              </Attributes>                                                                                                       ' + "\r\n" + '\
            </Element>                                                                                                            ' + "\r\n" + '\
          </StyledText>                                                                                                           ' + "\r\n" + '\
        </TextObject>                                                                                                             ' + "\r\n" + '\
        <Length>6000</Length>                                                                                                     ' + "\r\n" + '\
        <LengthMode>Auto</LengthMode>                                                                                             ' + "\r\n" + '\
        <BorderWidth>0</BorderWidth>                                                                                              ' + "\r\n" + '\
        <BorderStyle>Solid</BorderStyle>                                                                                          ' + "\r\n" + '\
        <BorderColor Alpha="255" Red="0" Green="0" Blue="0" />                                                                    ' + "\r\n" + '\
      </Cell>                                                                                                                     ' + "\r\n" + '\
    </Subcells>                                                                                                                   ' + "\r\n" + '\
  </RootCell>                                                                                                                     ' + "\r\n" + '\
</ContinuousLabel>';
    }
    function initialize()
    {
      configuration.initialize();
      var compatibility         = document.getElementById("compatibility");
      compatibility.innerHTML   = "<span class='label'>ACTIVEX SUPPORTED</span><span    style='font-weight:bold;'>" + (configuration.backwardCompatibility.activeX.isSupported    ? "<span style='color:green;'>&check;</span>" : "<span style='color:red;'>NO</span>") + "</span><br>" +
                                  "<span class='label'>DYMOADDIN SUPPORTED</span><span  style='font-weight:bold;'>" + (configuration.backwardCompatibility.dymoAddIn.isSupported  ? "<span style='color:green;'>&check;</span>" : "<span style='color:red;'>NO</span>") + "</span><br>" +
                                  "<span class='label'>DYMOTAPE SUPPORTED</span><span   style='font-weight:bold;'>" + (configuration.backwardCompatibility.dymoTape.isSupported   ? "<span style='color:green;'>&check;</span>" : "<span style='color:red;'>NO</span>") + "</span><br>";
      if (configuration.backwardCompatibility.dymoAddIn.isSupported)
      {
        var printers    = configuration.depricatedPrinters;
        var printerList = document.getElementById("printer");
        for (var i = 0, length = printers.length; i < length; i++)
        {
          printerList.options.add(new Option(printers[i], printers[i]));
        }
        setIndexByValue("printer", getCookie("printer"));
      }
      dymo.label.framework.init(configuration.display);
    }
    function print()
    {
      var gageId              = document.getElementById("gageId"              );
      var serialNumber        = document.getElementById("serialNumber"        );
      var model               = document.getElementById("model"               );
      var calibratedOn        = document.getElementById("calibratedOn"        );
      var inServiceOn         = document.getElementById("inServiceOn"         );
      var nextDueOn           = document.getElementById("nextDueOn"           );
      var printer             = document.getElementById("printer"             );
      var tapeSize            = document.getElementById("tapeSize"            );
      var boldFont            = document.getElementById("boldFont"            );
      var includeSerialNumber = document.getElementById("includeSerialNumber" );
      var includeModel        = document.getElementById("includeModel"        );
      boldFont                = boldFont.checked;
      includeSerialNumber     = tapeSize.value != configuration.minimumTapeSize && includeSerialNumber.checked;
      includeModel            = tapeSize.value != configuration.minimumTapeSize && includeModel.checked;
      var printJob            = {
                                  BOLD_FONT     : (boldFont ? "True" : "False") ,
                                  LABEL_NAME    : "Tape" + tapeSize.value       ,
                                  PAPER_NAME    : tapeSize.value                ,
                                  OUTPUT        : "Gage ID     : " + gageId.value       + "\r\n" +
                                                  (includeSerialNumber  ? "Serial      : " + serialNumber.value  + "\r\n" : "")  +
                                                  (includeModel         ? "Model       : " + model.value         + "\r\n" : "")  +
                                                  "Calibrated  : " + calibratedOn.value + "\r\n" +
                                                  "In Service  : " + inServiceOn.value  + "\r\n" +
                                                  "Next Due    : " + nextDueOn.value
                                };
      var output              = printJob.OUTPUT;
      if (!configuration.depricatedPrinters.includes(printer.value))
      {
        output = getXmlFormat();
        for (var placeholder in printJob)
        {
          if (!printJob.hasOwnProperty(placeholder)){ continue; }
          output = output.replace("[[[" + placeholder + "]]]", printJob[placeholder]);
        }
        printJob = dymo.label.framework.openLabelXml(output.replace(/\s+$/g, ""));
        printJob.print(printer.value);
        return true;
      }
      output = output.replace(/("\r\n")+/g, "\r").replace(/\s+$/g, "");
      try
      {
        var dymoAddIn = new ActiveXObject(configuration.backwardCompatibility.dymoAddIn.interface );
        var dymoTape  = new ActiveXObject(configuration.backwardCompatibility.dymoTape.interface  );
        if (printer.value != configuration.defaultPrinter){ dymoAddIn.SelectPrinter(printer.value); }
        if (dymoTape.New())
        {
          dymoTape.AddText(output, "Courier New, 12" + (boldFont ? ", Bold" : ""), false, false);
          dymoTape.Print(1);
        } else {
          if (dymoAddIn.SmartPasteFromString(output))
          {
            dymoAddIn.Print(1, false);
          } else {
            throw configuration.backwardCompatibility.dymoAddIn.interface + " And " + configuration.backwardCompatibility.dymoTape.interface + " Interfaces Failed.";
          }
        }
      } catch (exception) {
        alert(exception + "\r\n\r\n" + "Please ensure the printer is connected and the DYMO DLS libraries are installed on this PC.");
        return false;
      }
      return true;
    }
    function setCookie(key, value, preserve)
    {
      var today   = new Date();
      var future  = new Date(today.getFullYear() + 10, today.getMonth(), today.getDate());
      var expires = preserve ? "expires=" + future.toUTCString() + ";" : "";
      document.cookie = key + "=" + value + ";" + expires + "path=/";
    }
    function setIndexByValue(select, value)
    {
      select = document.getElementById(select);
      for (var i = 0, length = select.options.length; i < length; i++)
      {
        if (select.options[i].value !== value){ continue; }
        select.selectedIndex = i;
        return true;
      }
      return false;
    }
    function setOptions()
    {
      var printer                   = document.getElementById("printer"             );
      var tapeSize                  = document.getElementById("tapeSize"            );
      var boldFont                  = document.getElementById("boldFont"            );
      var includeSerialNumber       = document.getElementById("includeSerialNumber" );
      var includeModel              = document.getElementById("includeModel"        );
      printer                       = printer.value     ? printer.value  : configuration.defaultPrinter;
      tapeSize                      = tapeSize.value    ? tapeSize.value : configuration.defaultTapeSize;
      boldFont                      = boldFont.checked  ? 1 : 0;
      setCookie("printer"   , printer   , true);
      setCookie("tapeSize"  , tapeSize  , true);
      setCookie("boldFont"  , boldFont  , true);
      includeSerialNumber.disabled  = includeModel.disabled = tapeSize == configuration.minimumTapeSize;
      if (tapeSize != configuration.minimumTapeSize){ return; }
      includeSerialNumber.checked   = includeModel.checked  = false;
    }
  </script>
</head>
<body>
  <div style="display:table-cell; vertical-align:top;">
  <span id="settings" style="display:inline-block;">
    <label for="printer" style="display:inline-block; white-space:nowrap; width:75px;">Printer</label>
    <select id="printer" name="printer" onchange="setOptions();"></select>
    <br>
    <label for="tapeSize" style="display:inline-block; white-space:nowrap; width:75px;">Tape Size</label>
    <select id="tapeSize" name="tapeSize" onchange="setOptions();">
      <option value="12mm">1/2&quot;</option>
      <option value="19mm" selected>3/4&quot;</option>
      <option value="24mm">1 &quot;</option>
    </select>
    <br>
    <label for="boldFont" style="display:inline-block; white-space:nowrap; width:75px;">Bold Font</label>
    <input id="boldFont" name="boldFont" type="checkbox" <?= ($boldFont ? 'checked' : '') ?> title="Bold Font" onchange="setOptions();">
    <br>
    <br>
    <label for="gageId" style="display:inline-block; white-space:nowrap; width:75px;">Gage ID</label>
    <input id="gageId" name="gageId" value="<?= $gageId ?>">
    <br>
    <label for="serialNumber" style="display:inline-block; white-space:nowrap; width:75px;">Serial Number</label>
    <input id="serialNumber" name="serialNumber" value="<?= $serialNumber ?>">
    <input id="includeSerialNumber" name="includeSerialNumber" type="checkbox" <?= ($includeSerialNumber ? 'checked' : '') ?> title="Include">
    <br>
    <label for="model" style="display:inline-block; white-space:nowrap; width:75px;">Model</label>
    <input id="model" name="model" value="<?= $model ?>">
    <input id="includeModel" name="includeModel" type="checkbox" <?= ($includeModel ? 'checked' : '') ?> title="Include">
    <br>
    <label for="calibratedOn" style="display:inline-block; white-space:nowrap; width:75px;">Calibrated</label>
    <input id="calibratedOn" name="calibratedOn" value="<?= $calibratedOn ?>">
    <br>
    <label for="inServiceOn" style="display:inline-block; white-space:nowrap; width:75px;">In Service</label>
    <input id="inServiceOn" name="inServiceOn" value="<?= $inServiceOn ?>">
    <br>
    <label for="nextDueOn" style="display:inline-block; white-space:nowrap; width:75px;">Next Due</label>
    <input id="nextDueOn" name="nextDueOn" value="<?= $nextDueOn ?>">
    <br>
    <br>
    <input type="button" onclick="print();" value="Print ...">
    <input type="button" onclick="self.close();" value="Close">
    <br>
    <br>
    <a href="?viewSource" target="viewSource" style="text-decoration:none;">View Source</a>
  </span>
  <span id="configuration" class="configuration" style="cursor:pointer;" onclick="display('compatibility');">
    <div style="padding-bottom:5px;">CONFIGURATION</div>
    <div id="compatibility" class="configuration" style="cursor:wait; display:none;"></div>
  </span>
  </div>
</body>