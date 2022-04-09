import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pdfkit
import re
from csv import writer
import pandas as pd
# for sending attachement
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


csvReportDir = "/Users/harshit/PycharmProjects/SendEmail/venv/G8_August_AA.csv"
emailID = yourEmail
password = passWord
#myHtml = r"""<html><head><meta content="text/html; charset=UTF-8" http-equiv="content-type"><style type="text/css">@import url('https://themes.googleusercontent.com/fonts/css?kit=GGSdX60RftK6aXtKOXIPyxyByNVkZsQI5tdlcGmdDrY');ol{margin:0;padding:0}table td,table th{padding:0}.c27{border-right-style:solid;padding:5pt 5pt 5pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;background-color:#93c47d;border-left-style:solid;border-bottom-width:1pt;width:102.8pt;border-top-color:#000000;border-bottom-style:solid}.c40{border-right-style:solid;padding:5pt 5pt 5pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;background-color:#38761d;border-left-style:solid;border-bottom-width:1pt;width:102.8pt;border-top-color:#000000;border-bottom-style:solid}.c28{border-right-style:solid;padding:5pt 5pt 5pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;background-color:#e06666;border-left-style:solid;border-bottom-width:1pt;width:102.8pt;border-top-color:#000000;border-bottom-style:solid}.c32{border-right-style:solid;padding:5pt 5pt 5pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;background-color:#edc448;border-left-style:solid;border-bottom-width:1pt;width:102.8pt;border-top-color:#000000;border-bottom-style:solid}.c50{border-right-style:solid;padding:5pt 5pt 5pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:434.2pt;border-top-color:#000000;border-bottom-style:solid}.c43{border-right-style:solid;padding:5pt 5pt 5pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:102.8pt;border-top-color:#000000;border-bottom-style:solid}.c14{border-right-style:solid;padding:5pt 5pt 5pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:210pt;border-top-color:#000000;border-bottom-style:solid}.c3{border-right-style:solid;padding:5pt 5pt 5pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:225pt;border-top-color:#000000;border-bottom-style:solid}.c18{border-right-style:solid;padding:5pt 5pt 5pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:113.2pt;border-top-color:#000000;border-bottom-style:solid}.c4{background-color:#ffffff;color:#000000;font-weight:700;text-decoration:none;vertical-align:baseline;font-size:12pt;font-family:"Comfortaa";font-style:normal}.c19{background-color:#38761d;color:#ffffff;font-weight:700;text-decoration:none;vertical-align:baseline;font-size:12pt;font-family:"Comfortaa";font-style:normal}.c0{background-color:#ffffff;color:#000000;font-weight:400;text-decoration:none;vertical-align:baseline;font-size:12pt;font-family:"Comfortaa";font-style:normal}.c25{color:#000000;font-weight:400;text-decoration:none;vertical-align:baseline;font-size:11pt;font-family:"Arial";font-style:normal}.c1{color:#000000;font-weight:400;text-decoration:none;vertical-align:baseline;font-size:12pt;font-family:"Comfortaa";font-style:normal}.c16{padding-top:0pt;padding-bottom:0pt;line-height:1.15;orphans:2;widows:2;text-align:left;height:11pt}.c12{padding-top:1.7pt;padding-bottom:0pt;line-height:1.0;text-align:left;margin-right:7.3pt;height:11pt}.c7{margin-left:5.9pt;padding-top:0pt;padding-bottom:0pt;line-height:1.0;text-align:left}.c21{margin-left:6.4pt;padding-top:0pt;padding-bottom:0pt;line-height:1.0;text-align:left}.c22{background-color:#38761d;color:#ffffff;text-decoration:none;vertical-align:baseline;font-style:normal}.c26{padding-top:0pt;padding-bottom:0pt;line-height:1.15;text-align:left;height:11pt}.c51{padding-top:1pt;padding-bottom:0pt;line-height:1.0;text-align:right;margin-right:182.7pt}.c24{padding-top:1pt;padding-bottom:0pt;line-height:1.0;text-align:center;margin-right:3pt}.c6{margin-left:39.2pt;padding-top:2.8pt;padding-bottom:0pt;line-height:1.0;text-align:left}.c34{margin-left:39.2pt;padding-top:2.8pt;padding-bottom:0pt;line-height:1.5;text-align:left}.c45{margin-left:5.9pt;padding-top:6.2pt;padding-bottom:0pt;line-height:1.0;text-align:left}.c30{padding-top:0pt;padding-bottom:0pt;line-height:1.0;text-align:left;height:11pt}.c47{margin-left:6pt;padding-top:6.2pt;padding-bottom:0pt;line-height:1.0;text-align:left}.c38{margin-left:40.5pt;border-spacing:0;border-collapse:collapse;margin-right:auto}.c5{padding-top:0pt;padding-bottom:0pt;line-height:1.0;text-align:center}.c17{margin-left:146.2pt;border-spacing:0;border-collapse:collapse;margin-right:auto}.c36{padding-top:0pt;padding-bottom:0pt;line-height:1.0;text-align:left}.c33{color:#ffffff;text-decoration:none;vertical-align:baseline;font-style:normal}.c46{margin-left:39.8pt;border-spacing:0;border-collapse:collapse;margin-right:auto}.c29{color:#000000;text-decoration:none;vertical-align:baseline;font-style:normal}.c23{font-size:12pt;font-family:"Comfortaa";font-weight:700}.c20{font-size:16pt;font-family:"Comfortaa";font-weight:700}.c13{font-size:12pt;font-family:"Comfortaa";font-weight:400}.c49{font-weight:700;font-size:10.5pt;font-family:"Comfortaa"}.c39{max-width:529.5pt;padding:13.5pt 49.5pt 102.8pt 33pt}.c41{font-weight:400;font-family:"Arial"}.c31{margin-left:6.5pt}.c35{margin-left:4.5pt}.c2{height:24.8pt}.c9{height:0pt}.c10{height:68.2pt}.c8{font-size:13pt}.c42{height:24pt}.c15{background-color:#ffffff}.c11{height:25.5pt}.c48{text-indent:36pt}.c44{height:11pt}.c37{height:27pt}.title{padding-top:24pt;color:#000000;font-weight:700;font-size:36pt;padding-bottom:6pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}.subtitle{padding-top:18pt;color:#666666;font-size:24pt;padding-bottom:4pt;font-family:"Georgia";line-height:1.15;page-break-after:avoid;font-style:italic;orphans:2;widows:2;text-align:left}li{color:#000000;font-size:11pt;font-family:"Arial"}p{margin:0;color:#000000;font-size:11pt;font-family:"Arial"}h1{padding-top:24pt;color:#000000;font-weight:700;font-size:24pt;padding-bottom:6pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}h2{padding-top:18pt;color:#000000;font-weight:700;font-size:18pt;padding-bottom:4pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}h3{padding-top:14pt;color:#000000;font-weight:700;font-size:14pt;padding-bottom:4pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}h4{padding-top:12pt;color:#000000;font-weight:700;font-size:12pt;padding-bottom:2pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}h5{padding-top:11pt;color:#000000;font-weight:700;font-size:11pt;padding-bottom:2pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}h6{padding-top:10pt;color:#000000;font-weight:700;font-size:10pt;padding-bottom:2pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}</style></head><body class="c15 c39"><p class="c36"><span style="overflow: hidden; display: inline-block; margin: 0.00px 0.00px; border: 0.00px solid #000000; transform: rotate(0.00rad) translateZ(0px); -webkit-transform: rotate(0.00rad) translateZ(0px); width: 187.00px; height: 60.00px;"><img alt="" src="/Users/harshit/PycharmProjects/SendEmail/venv/AATemplate/images/image1.png" style="width: 187.00px; height: 60.00px; margin-left: 0.00px; margin-top: 0.00px; transform: rotate(0.00rad) translateZ(0px); -webkit-transform: rotate(0.00rad) translateZ(0px);" title=""></span><span class="c29 c49">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></p><p class="c30"><span class="c29 c49"></span></p><p class="c5"><span class="c4">ACADEMIC SESSION: 202</span><span class="c23 c15">1</span><span class="c4">-2</span><span class="c23 c15">2</span><span class="c29 c23">&nbsp;</span></p><p class="c24 c44"><span class="c4"></span></p><p class="c24"><span class="c15 c20">AA Report - 1st</span><span class="c29 c20 c15">&nbsp;- 31st July</span></p><p class="c24 c44"><span class="c4"></span></p><p class="c51"><span class="c23 c29">&nbsp;</span></p><a id="t.844be22974995d54267db67c01786a11af3c6e02"></a><a id="t.0"></a><table class="c38"><tbody><tr class="c10"><td class="c50" colspan="1" rowspan="1"><p class="c7"><span class="c0">Student&rsquo;s Name:</span><span class="c13 c15">&nbsp; FirstName</span><span class="c0">&nbsp;</span><span class="c13 c15">LastName</span></p><p class="c45"><span class="c0">Sitare ID: &nbsp;</span><span class="c13 c15">SitareID</span></p><p class="c47"><span class="c0">Class: &nbsp;</span><span class="c13 c15">grade</span></p></td></tr></tbody></table><p class="c26"><span class="c25"></span></p><p class="c26 c35"><span class="c25"></span></p><a id="t.6365348f537921348ad8144f77491c834249e43e"></a><a id="t.1"></a><table class="c46"><tbody><tr class="c37"><td class="c14" colspan="1" rowspan="1"><p class="c7"><span class="c0">Subjects</span><span class="c1">&nbsp;</span></p></td><td class="c3" colspan="1" rowspan="1"><p class="c5 c31"><span class="c13 c15">Assignment </span><span class="c13">( % )</span><span class="c1">&nbsp;</span></p></td></tr><tr class="c2"><td class="c14" colspan="1" rowspan="1"><p class="c21"><span class="c0">English</span><span class="c1">&nbsp;</span></p></td><td class="c3" colspan="1" rowspan="1"><p class="c5"><span class="c13">EngMarksP</span></p></td></tr><tr class="c2"><td class="c14" colspan="1" rowspan="1"><p class="c21"><span class="c13 c15">Hindi</span><span class="c1">&nbsp;</span></p></td><td class="c3" colspan="1" rowspan="1"><p class="c5"><span class="c13">HinMarksP</span></p></td></tr><tr class="c2"><td class="c14" colspan="1" rowspan="1"><p class="c21"><span class="c13 c15">Mathematics</span><span class="c13">&nbsp;</span></p></td><td class="c3" colspan="1" rowspan="1"><p class="c5"><span class="c13">MatMarksP</span></p></td></tr><tr class="c42"><td class="c14" colspan="1" rowspan="1"><p class="c7"><span class="c13 c15">Science</span><span class="c13">&nbsp;</span></p></td><td class="c3" colspan="1" rowspan="1"><p class="c5"><span class="c13">SciMarksP</span></p></td></tr><tr class="c11"><td class="c14" colspan="1" rowspan="1"><p class="c7"><span class="c13 c15">Social Studies</span><span class="c13">&nbsp;</span></p></td><td class="c3" colspan="1" rowspan="1"><p class="c5"><span class="c13">SstMarksP</span></p></td></tr></tbody></table><p class="c26"><span class="c25"></span></p><p class="c26"><span class="c25"></span></p><p class="c34"><span class="c13 c15">Assignment a</span><span class="c0">verage (%): </span><span class="c23 c15">&nbsp;AssignmentP</span></p><p class="c6"><span class="c13 c15">Attendance average (%): </span><span class="c15 c23">&nbsp;AttendanceP</span></p><p class="c6"><span class="c13 c15">Buckets:</span></p><p class="c12 c48"><span class="c19"></span></p><p class="c12 c48"><span class="c19"></span></p><a id="t.67e10fae5f02522e513d39f00c975e9f209d6eeb"></a><a id="t.2"></a><table class="c17"><tbody><tr class="c9"><td class="c32" colspan="1" rowspan="1"><p class="c36"><span class="c1">Gold</span></p></td><td class="c18" colspan="1" rowspan="1"><p class="c5"><span class="c8">90% or more</span></p></td></tr><tr class="c9"><td class="c40" colspan="1" rowspan="1"><p class="c36"><span class="c13 c22">Dark Green </span></p></td><td class="c18" colspan="1" rowspan="1"><p class="c5"><span class="c8">89.99% - 80%</span></p></td></tr><tr class="c9"><td class="c27" colspan="1" rowspan="1"><p class="c36"><span class="c13 c33">Light Green</span></p></td><td class="c18" colspan="1" rowspan="1"><p class="c5"><span class="c29 c41 c8">79.99% - 70%</span></p></td></tr><tr class="c9"><td class="c43" colspan="1" rowspan="1"><p class="c36"><span class="c1">White</span></p></td><td class="c18" colspan="1" rowspan="1"><p class="c5"><span class="c29 c8 c41">69.99% - 60%</span></p></td></tr><tr class="c9"><td class="c28" colspan="1" rowspan="1"><p class="c36"><span class="c33 c13">Red</span></p></td><td class="c18" colspan="1" rowspan="1"><p class="c5"><span class="c8">Below 60%</span></p></td></tr></tbody></table><p class="c12"><span class="c19"></span></p><div><p class="c16"><span class="c25"></span></p></div></body></html>"""
myHtml8 = r"""<html><head><meta content="text/html; charset=UTF-8" http-equiv="content-type"><style type="text/css">@import url('https://themes.googleusercontent.com/fonts/css?kit=GGSdX60RftK6aXtKOXIPyxyByNVkZsQI5tdlcGmdDrY');ol{margin:0;padding:0}table td,table th{padding:0}.c34{border-right-style:solid;padding:5pt 5pt 5pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;background-color:#edc448;border-left-style:solid;border-bottom-width:1pt;width:102.8pt;border-top-color:#000000;border-bottom-style:solid}.c21{border-right-style:solid;padding:5pt 5pt 5pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;background-color:#93c47d;border-left-style:solid;border-bottom-width:1pt;width:102.8pt;border-top-color:#000000;border-bottom-style:solid}.c6{border-right-style:solid;padding:5pt 5pt 5pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:210pt;border-top-color:#000000;border-bottom-style:solid}.c41{border-right-style:solid;padding:5pt 5pt 5pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:102.8pt;border-top-color:#000000;border-bottom-style:solid}.c1{border-right-style:solid;padding:5pt 5pt 5pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:110.2pt;border-top-color:#000000;border-bottom-style:solid}.c37{border-right-style:solid;padding:5pt 5pt 5pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:434.2pt;border-top-color:#000000;border-bottom-style:solid}.c24{border-right-style:solid;padding:5pt 5pt 5pt 5pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:top;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:225pt;border-top-color:#000000;border-bottom-style:solid}.c4{background-color:#38761d;color:#ffffff;font-weight:700;text-decoration:none;vertical-align:baseline;font-size:12pt;font-family:"Comfortaa";font-style:normal}.c40{color:#000000;font-weight:700;text-decoration:none;vertical-align:baseline;font-size:16pt;font-family:"Comfortaa";font-style:normal}.c28{color:#000000;font-weight:400;text-decoration:none;vertical-align:baseline;font-size:11pt;font-family:"Arial";font-style:normal}.c15{color:#000000;font-weight:700;text-decoration:none;vertical-align:baseline;font-size:12pt;font-family:"Comfortaa";font-style:normal}.c2{padding-top:1.7pt;text-indent:36pt;padding-bottom:0pt;line-height:1.0;text-align:left;margin-right:7.3pt;height:11pt}.c0{color:#000000;font-weight:400;text-decoration:none;vertical-align:baseline;font-size:12pt;font-family:"Comfortaa";font-style:normal}.c18{color:#000000;font-weight:400;text-decoration:none;vertical-align:baseline;font-size:13pt;font-family:"Arial";font-style:normal}.c42{color:#000000;font-weight:700;text-decoration:none;vertical-align:baseline;font-size:10.5pt;font-family:"Comfortaa";font-style:normal}.c30{padding-top:1.7pt;padding-bottom:0pt;line-height:1.0;text-align:left;margin-right:7.3pt;height:11pt}.c3{padding-top:1pt;padding-bottom:0pt;line-height:1.0;text-align:center;margin-right:3pt;height:11pt}.c5{padding-top:0pt;padding-bottom:0pt;line-height:1.15;text-align:left;height:11pt}.c19{padding-top:1pt;padding-bottom:0pt;line-height:1.0;text-align:right;margin-right:182.7pt}.c12{margin-left:6pt;padding-top:6.2pt;padding-bottom:0pt;line-height:1.0;text-align:left}.c49{padding-top:1pt;padding-bottom:0pt;line-height:1.0;text-align:center;margin-right:3pt}.c33{margin-left:39.2pt;padding-top:2.8pt;padding-bottom:0pt;line-height:1.5;text-align:left}.c32{background-color:#38761d;color:#ffffff;text-decoration:none;vertical-align:baseline;font-style:normal}.c36{margin-left:39.2pt;padding-top:2.8pt;padding-bottom:0pt;line-height:1.0;text-align:left}.c16{padding-top:0pt;padding-bottom:0pt;line-height:1.0;text-align:center}.c38{color:#ffffff;text-decoration:none;vertical-align:baseline;font-style:normal}.c8{margin-left:39.8pt;border-spacing:0;border-collapse:collapse;margin-right:auto}.c11{padding-top:0pt;padding-bottom:0pt;line-height:1.0;text-align:left}.c48{padding-top:6.2pt;padding-bottom:0pt;line-height:1.0;text-align:left}.c9{background-color:#ffffff;font-size:12pt;font-family:"Comfortaa";font-weight:400}.c25{margin-left:146.2pt;border-spacing:0;border-collapse:collapse;margin-right:auto}.c39{margin-left:40.5pt;border-spacing:0;border-collapse:collapse;margin-right:auto}.c50{font-size:16pt;font-family:"Comfortaa";font-weight:700}.c23{font-size:12pt;font-family:"Comfortaa";font-weight:400}.c17{font-size:12pt;font-family:"Comfortaa";font-weight:700}.c29{max-width:529.5pt;padding:13.5pt 49.5pt 102.8pt 33pt}.c46{orphans:2;widows:2}.c20{font-size:13pt}.c22{margin-left:6.5pt}.c27{margin-left:5.9pt}.c10{background-color:#ffffff}.c31{height:11pt}.c51{height:25.5pt}.c35{height:68.2pt}.c26{margin-left:6.4pt}.c13{height:24.8pt}.c7{height:0pt}.c44{background-color:#38761d}.c43{margin-left:4.5pt}.c45{height:27pt}.c14{height:24pt}.c47{background-color:#e06666}.title{padding-top:24pt;color:#000000;font-weight:700;font-size:36pt;padding-bottom:6pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}.subtitle{padding-top:18pt;color:#666666;font-size:24pt;padding-bottom:4pt;font-family:"Georgia";line-height:1.15;page-break-after:avoid;font-style:italic;orphans:2;widows:2;text-align:left}li{color:#000000;font-size:11pt;font-family:"Arial"}p{margin:0;color:#000000;font-size:11pt;font-family:"Arial"}h1{padding-top:24pt;color:#000000;font-weight:700;font-size:24pt;padding-bottom:6pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}h2{padding-top:18pt;color:#000000;font-weight:700;font-size:18pt;padding-bottom:4pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}h3{padding-top:14pt;color:#000000;font-weight:700;font-size:14pt;padding-bottom:4pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}h4{padding-top:12pt;color:#000000;font-weight:700;font-size:12pt;padding-bottom:2pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}h5{padding-top:11pt;color:#000000;font-weight:700;font-size:11pt;padding-bottom:2pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}h6{padding-top:10pt;color:#000000;font-weight:700;font-size:10pt;padding-bottom:2pt;font-family:"Arial";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}</style></head><body class="c10 c29"><p class="c11"><span style="overflow: hidden; display: inline-block; margin: 0.00px 0.00px; border: 0.00px solid #000000; transform: rotate(0.00rad) translateZ(0px); -webkit-transform: rotate(0.00rad) translateZ(0px); width: 187.00px; height: 60.00px;"><img alt="" src="/Users/harshit/PycharmProjects/SendEmail/venv/AATemplate/images/image1.png" style="width: 187.00px; height: 60.00px; margin-left: 0.00px; margin-top: 0.00px; transform: rotate(0.00rad) translateZ(0px); -webkit-transform: rotate(0.00rad) translateZ(0px);" title=""></span><span class="c42">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></p><p class="c11 c31"><span class="c42"></span></p><p class="c16"><span class="c15 c10">ACADEMIC SESSION: 202</span><span class="c17 c10">1</span><span class="c15 c10">-2</span><span class="c17 c10">2</span><span class="c15">&nbsp;</span></p><p class="c3"><span class="c15 c10"></span></p><p class="c49"><span class="c10 c50">AA Report - MonthHead</span></p><p class="c3"><span class="c15 c10"></span></p><p class="c19"><span class="c15">&nbsp;</span></p><a id="t.e0780b7472ff08f62d699c29781365dc86b2f70b"></a><a id="t.0"></a><table class="c39"><tbody><tr class="c35"><td class="c37" colspan="1" rowspan="1"><p class="c11 c27"><span class="c0 c10">Student&rsquo;s Name:</span><span class="c9">&nbsp;FirstName LastName</span></p><p class="c27 c48"><span class="c0 c10">Sitare ID: &nbsp;</span><span class="c9">SitareID</span></p><p class="c12"><span class="c0 c10">Class: &nbsp;</span><span class="c9">grade</span></p></td></tr></tbody></table><p class="c5"><span class="c28"></span></p><p class="c5 c43"><span class="c28"></span></p><a id="t.6365348f537921348ad8144f77491c834249e43e"></a><a id="t.1"></a><table class="c8"><tbody><tr class="c45"><td class="c6" colspan="1" rowspan="1"><p class="c11 c27"><span class="c0 c10">Subjects</span><span class="c0">&nbsp;</span></p></td><td class="c24" colspan="1" rowspan="1"><p class="c16 c22"><span class="c9">Assignment </span><span class="c23">( % )</span><span class="c0">&nbsp;</span></p></td></tr><tr class="c13"><td class="c6" colspan="1" rowspan="1"><p class="c11 c26"><span class="c0 c10">English</span><span class="c0">&nbsp;</span></p></td><td class="c24" colspan="1" rowspan="1"><p class="c16"><span class="c9">EngMarksP</span></p></td></tr><tr class="c13"><td class="c6" colspan="1" rowspan="1"><p class="c11 c26"><span class="c9">Hindi</span><span class="c0">&nbsp;</span></p></td><td class="c24" colspan="1" rowspan="1"><p class="c16"><span class="c9">HinMarksP</span></p></td></tr><tr class="c13"><td class="c6" colspan="1" rowspan="1"><p class="c11 c26"><span class="c9">Mathematics</span><span class="c23">&nbsp;</span></p></td><td class="c24" colspan="1" rowspan="1"><p class="c16"><span class="c9">MatMarksP</span></p></td></tr><tr class="c14"><td class="c6" colspan="1" rowspan="1"><p class="c11 c27"><span class="c9">Science</span><span class="c23">&nbsp;</span></p></td><td class="c24" colspan="1" rowspan="1"><p class="c16"><span class="c9">SciMarksP</span></p></td></tr><tr class="c51"><td class="c6" colspan="1" rowspan="1"><p class="c11 c27"><span class="c9">Social Studies</span><span class="c23">&nbsp;</span></p></td><td class="c24" colspan="1" rowspan="1"><p class="c16"><span class="c9">SstMarksP</span></p></td></tr></tbody></table><p class="c5"><span class="c28"></span></p><p class="c5"><span class="c28"></span></p><p class="c33"><span class="c9">Assignment a</span><span class="c0 c10">verage (%): &nbsp;</span><span class="c10 c17">AssignmentP</span></p><p class="c36"><span class="c9">Attendance average (%): </span><span class="c10 c15">AttendanceP</span></p><p class="c36 c31"><span class="c0 c10"></span></p><p class="c36"><span class="c9">Buckets:</span></p><p class="c2"><span class="c4"></span></p><p class="c2"><span class="c4"></span></p><a id="t.0100c6157e7069cf2eca10eeb607fd288a529a76"></a><a id="t.2"></a><table class="c25"><tbody><tr class="c7"><td class="c34" colspan="1" rowspan="1"><p class="c11"><span class="c0">Gold</span></p></td><td class="c1" colspan="1" rowspan="1"><p class="c16"><span class="c20">90% or more</span></p></td></tr><tr class="c7"><td class="c41 c44" colspan="1" rowspan="1"><p class="c11"><span class="c23 c32">Dark Green </span></p></td><td class="c1" colspan="1" rowspan="1"><p class="c16"><span class="c20">89.99% - 70%</span></p></td></tr><tr class="c7"><td class="c21" colspan="1" rowspan="1"><p class="c11"><span class="c23 c38">Light Green</span></p></td><td class="c1" colspan="1" rowspan="1"><p class="c16"><span class="c18">69.99% - 60%</span></p></td></tr><tr class="c7"><td class="c41" colspan="1" rowspan="1"><p class="c11"><span class="c0">White</span></p></td><td class="c1" colspan="1" rowspan="1"><p class="c16"><span class="c18">59.99% - 50%</span></p></td></tr><tr class="c7"><td class="c41 c47" colspan="1" rowspan="1"><p class="c11"><span class="c38 c23">Red</span></p></td><td class="c1" colspan="1" rowspan="1"><p class="c16"><span class="c20">Below 50%</span></p></td></tr></tbody></table><p class="c30"><span class="c4"></span></p><div><p class="c5 c46"><span class="c28"></span></p></div></body></html>"""
readCsv = pd.read_csv(csvReportDir)
monthHead = "1st to 31st August 2021"
# htmlFile = open("/Users/harshit/PycharmProjects/SendEmail/venv/AATemplate/MonthlyReportAATemplateG9.html", "r")

session = smtplib.SMTP('smtp.gmail.com', 587)

# enable security
session.starttls()

# login with mail_id and password
session.login(emailID, password)


def pdfMail(sitareID, FirstName):
    print(f"Now sending mail to {sitareID}")
    body = f"""Hi {FirstName}
Hope you are safe and healthy.

Please find your assignment and attendance report for month of August '2021 (attached). 
If you are in the gold/green buckets, good job! If you are in the white or red buckets, you should aim to get to the green buckets. 

NOTE : Please talk to your city coordinator if you have any questions.

Best wishes,
Sitare Academic Team
"""
    receiver = f'{sitareID}@sitare.org'
    # sender = emailID
    # password = password
    message = MIMEMultipart()
    message['From'] = emailID
    message['To'] = receiver
    message['Subject'] = 'AA Report : August 2021'
    message.attach(MIMEText(body, 'plain'))  # add body as plain text
    pdfname = f'{sitareID}.pdf'
    # open the file in binary
    binary_pdf = open(pdfname, 'rb')

    payload = MIMEBase('application', 'octate-stream', Name=pdfname)
    payload.set_payload(binary_pdf.read())

    # enconding the binary into base64
    encoders.encode_base64(payload)

    # add header with pdf name
    payload.add_header('Content-Decomposition', 'attachment', filename=pdfname)
    message.attach(payload)
    # use gmail with port
    # session = smtplib.SMTP('smtp.gmail.com', 587)

    # enable security
    # session.starttls()

    # login with mail_id and password
    # session.login(emailID, password)

    text = message.as_string()
    session.sendmail(emailID, receiver, text)
    print(f'Mail Sent to {sitareID}')
    # session.quit()


def mailHtml():
    # me == my email address
    # you == recipient's email address
    me = emailID
    you = "harshit@sitare.org"

    # Create message container - the correct MIME type is multipart/alternative.
    msg = MIMEMultipart('alternative')
    msg['Subject'] = "Link"
    msg['From'] = me
    msg['To'] = you

    # Create the body of the message (a plain-text and an HTML version).
    text = "Hi!\nHow are you?\nHere is the link you wanted:\nhttp://www.python.org"
    html = ""

    # Record the MIME types of both parts - text/plain and text/html.
    part1 = MIMEText(text, 'plain')
    part2 = MIMEText(html, 'html')

    # Attach parts into message container.
    # According to RFC 2046, the last part of a multipart message, in this case
    # the HTML message, is best and preferred.
    msg.attach(part1)
    msg.attach(part2)
    # Send the message via local SMTP server.
    mail = smtplib.SMTP('smtp.gmail.com', 587)

    mail.ehlo()

    mail.starttls()

    mail.login(me, password)
    mail.sendmail(me, you, msg.as_string())
    mail.quit()
    print("Sucess")


def createPDF(sitareID, grade, firstName, lastName, engP, hinp, matP, sciP, sstP, assignmentP, attendanceP):
    # newHtml = str(htmlFile.read())
    newHtml = re.sub("MonthHead", monthHead, myHtml8)
    newHtml = re.sub("FirstName", firstName, newHtml)
    newHtml = re.sub("LastName", lastName, newHtml)
    newHtml = re.sub("SitareID", sitareID, newHtml)
    newHtml = re.sub("grade", grade, newHtml)
    newHtml = re.sub("EngMarksP", engP, newHtml)
    newHtml = re.sub("HinMarksP", hinp, newHtml)
    newHtml = re.sub("MatMarksP", matP, newHtml)
    newHtml = re.sub("SciMarksP", sciP, newHtml)
    newHtml = re.sub("SstMarksP", sstP, newHtml)
    newHtml = re.sub("AssignmentP", assignmentP, newHtml)
    newHtml = re.sub("AttendanceP", attendanceP, newHtml)
    options = {
        #'page-size': 'A4',
        'dpi': '400',
        'page-height': '14.0in',
        'page-width': '8.5in',
        'margin-top': '0.75in',
        'margin-right': '0.75in',
        'margin-bottom': '0.75in',
        'margin-left': '0.75in',
        'encoding': "UTF-8",
        'quiet': '',
        'disable-smart-shrinking': '',
        'lowquality': None,
        'custom-header': [
            ('Accept-Encoding', 'gzip')
        ],
        'allow': ["/Users/harshit/PycharmProjects/SendEmail/venv/AATemplate/images/"]
        ,
        'cookie': [
            ('cookie-empty-value', '""'),
            ('cookie-name1', 'cookie-value1'),
            ('cookie-name2', 'cookie-value2'),
        ],
        'no-outline': None
    }
    pdfkit.from_string(newHtml, f"/Users/harshit/PycharmProjects/SendEmail/{sitareID}.pdf", options=options)
    print(f"PDF Generated of {sitareID}")


def iterateCsv():
    i = 1
    try:
        for sitareID,  Grade, Firstname, Lastname, English, Hindi, Mathematics, Science, Sst, Assignment, Attendance in zip(
                readCsv["Sitare Id"], readCsv["Grade"], readCsv["Firstname"], readCsv["Lastname"],
                readCsv["English"], readCsv["Hindi"], readCsv["Mathematics"], readCsv["Science"], readCsv["Sst"],
                readCsv["Assignment"], readCsv["Attendance"]):
            createPDF(str(sitareID), str(Grade), str(Firstname), str(Lastname), str(English), str(Hindi),
                      str(Mathematics), str(Science), str(Sst),
                      str(Assignment), str(Attendance))
            pdfMail(sitareID, Firstname)
            i = i + 1
    except Exception:
        print(f"SomeThing Went wrong line {i}  with {Exception.with_traceback()}")


def readTest():
    htmlFile = open("/Users/harshit/PycharmProjects/SendEmail/venv/AATemplate/MonthlyReportAATemplateG9.html", "r")
    se = htmlFile.read()


if __name__ == '__main__':
    iterateCsv()
    session.quit()
    # readTest()
