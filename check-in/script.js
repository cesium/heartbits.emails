var EMAIL_SENT = 'EMAIL_SENT';

var EMAIL_SUBJECT = 'Hackathon Heartbits check-in';
var BODY_1 = 'Boa tarde caro(a) participante,\n\nEstá tudo pronto para o início da Hackathon HeartBits! Já podes aderir ao Discord, a nossa plataforma principal! Basta acederes a este link: https://discord.gg/6djqZbB e seguir as indicações que se encontram na plataforma. O teu código pessoal, que vais necessitar para fazer o check-in é: ';
var BODY_2 = '.\nDeixamos a ressalva de que as equipas podem vir a sofrer ligeiros ajustes até ao momento do início da atividade, em virtude de eventuais desistências. As equipas que se inscreveram já completas excetuam-se à ressalva anterior!\n\nEstamos à tua espera no Discord! Junta-te a nós!!\n\n\nAté já,\nA equipa organizadora da Hackathon HeartBits 2020';

var HTML_BODY_1 = '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional //EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"> <!--[if IE]><html xmlns="http://www.w3.org/1999/xhtml" class="ie"><![endif]--> <!--[if !IE]><!--><html style="margin: 0;padding: 0;" xmlns="http://www.w3.org/1999/xhtml"><!--<![endif]--><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"><title></title> <!--[if !mso]><!--><meta http-equiv="X-UA-Compatible" content="IE=edge"><!--<![endif]--><meta name="viewport" content="width=device-width"><style type="text/css">@media only screen and (min-width: 620px){.wrapper{min-width:600px !important}.wrapper h1{}.wrapper h1{font-size:32px !important;line-height:40px !important}.wrapper h2{}.wrapper h2{font-size:30px !important;line-height:38px !important}.wrapper h3{}.column{}.wrapper .size-8{font-size:8px !important;line-height:14px !important}.wrapper .size-9{font-size:9px !important;line-height:16px !important}.wrapper .size-10{font-size:10px !important;line-height:18px !important}.wrapper .size-11{font-size:11px !important;line-height:19px !important}.wrapper .size-12{font-size:12px !important;line-height:19px !important}.wrapper .size-13{font-size:13px !important;line-height:21px !important}.wrapper .size-14{font-size:14px !important;line-height:21px !important}.wrapper .size-15{font-size:15px !important;line-height:23px !important}.wrapper .size-16{font-size:16px !important;line-height:24px !important}.wrapper .size-17{font-size:17px !important;line-height:26px !important}.wrapper .size-18{font-size:18px !important;line-height:26px !important}.wrapper .size-20{font-size:20px !important;line-height:28px !important}.wrapper .size-22{font-size:22px !important;line-height:31px !important}.wrapper .size-24{font-size:24px !important;line-height:32px !important}.wrapper .size-26{font-size:26px !important;line-height:34px !important}.wrapper .size-28{font-size:28px !important;line-height:36px !important}.wrapper .size-30{font-size:30px !important;line-height:38px !important}.wrapper .size-32{font-size:32px !important;line-height:40px !important}.wrapper .size-34{font-size:34px !important;line-height:43px !important}.wrapper .size-36{font-size:36px !important;line-height:43px !important}.wrapper .size-40{font-size:40px !important;line-height:47px !important}.wrapper .size-44{font-size:44px !important;line-height:50px !important}.wrapper .size-48{font-size:48px !important;line-height:54px !important}.wrapper .size-56{font-size:56px !important;line-height:60px !important}.wrapper .size-64{font-size:64px !important;line-height:63px !important}}</style><meta name="x-apple-disable-message-reformatting"><style type="text/css">/*<![CDATA[*/body{margin:0;padding:0}table{border-collapse:collapse;table-layout:fixed}*{line-height:inherit}[x-apple-data-detectors]{color:inherit !important;text-decoration:none !important}.wrapper .footer__share-button a:hover, .wrapper .footer__share-button a:focus{color:#fff !important}.btn a:hover, .btn a:focus, .footer__share-button a:hover, .footer__share-button a:focus, .email-footer__links a:hover, .email-footer__links a:focus{opacity:0.8}.preheader,.header,.layout,.column{transition:width 0.25s ease-in-out, max-width 0.25s ease-in-out}.preheader td{padding-bottom:8px}.layout,div.header{max-width:400px !important;-fallback-width:95% !important;width:calc(100% - 20px) !important}div.preheader{max-width:360px !important;-fallback-width:90% !important;width:calc(100% - 60px) !important}.snippet,.webversion{Float:none !important}.stack .column{max-width:400px !important;width:100% !important}.fixed-width.has-border{max-width:402px !important}.fixed-width.has-border .layout__inner{box-sizing:border-box}.snippet,.webversion{width:50% !important}.ie .btn{width:100%}.ie .stack .column, .ie .stack .gutter{display:table-cell;float:none !important}.ie div.preheader, .ie .email-footer{max-width:560px !important;width:560px !important}.ie .snippet, .ie .webversion{width:280px !important}.ie div.header, .ie .layout{max-width:600px !important;width:600px !important}.ie .two-col .column{max-width:300px !important;width:300px !important}.ie .three-col .column, .ie .narrow{max-width:200px !important;width:200px !important}.ie .wide{width:400px !important}.ie .stack.fixed-width.has-border, .ie .stack.has-gutter.has-border{max-width:602px !important;width:602px !important}.ie .stack.two-col.has-gutter .column{max-width:290px !important;width:290px !important}.ie .stack.three-col.has-gutter .column, .ie .stack.has-gutter .narrow{max-width:188px !important;width:188px !important}.ie .stack.has-gutter .wide{max-width:394px !important;width:394px !important}.ie .stack.two-col.has-gutter.has-border .column{max-width:292px !important;width:292px !important}.ie .stack.three-col.has-gutter.has-border .column, .ie .stack.has-gutter.has-border .narrow{max-width:190px !important;width:190px !important}.ie .stack.has-gutter.has-border .wide{max-width:396px !important;width:396px !important}.ie .fixed-width .layout__inner{border-left:0 none white !important;border-right:0 none white !important}.ie .layout__edges{display:none}.mso .layout__edges{font-size:0}.layout-fixed-width, .mso .layout-full-width{background-color:#fff}@media only screen and (min-width: 620px){.column,.gutter{display:table-cell;Float:none !important;vertical-align:top}div.preheader,.email-footer{max-width:560px !important;width:560px !important}.snippet,.webversion{width:280px !important}div.header, .layout, .one-col .column{max-width:600px !important;width:600px !important}.fixed-width.has-border,.fixed-width.x_has-border,.has-gutter.has-border,.has-gutter.x_has-border{max-width:602px !important;width:602px !important}.two-col .column{max-width:300px !important;width:300px !important}.three-col .column,.column.narrow,.column.x_narrow{max-width:200px !important;width:200px !important}.column.wide,.column.x_wide{width:400px !important}.two-col.has-gutter .column, .two-col.x_has-gutter .column{max-width:290px !important;width:290px !important}.three-col.has-gutter .column, .three-col.x_has-gutter .column, .has-gutter .narrow{max-width:188px !important;width:188px !important}.has-gutter .wide{max-width:394px !important;width:394px !important}.two-col.has-gutter.has-border .column, .two-col.x_has-gutter.x_has-border .column{max-width:292px !important;width:292px !important}.three-col.has-gutter.has-border .column, .three-col.x_has-gutter.x_has-border .column, .has-gutter.has-border .narrow, .has-gutter.x_has-border .narrow{max-width:190px !important;width:190px !important}.has-gutter.has-border .wide, .has-gutter.x_has-border .wide{max-width:396px !important;width:396px !important}}@supports (display: flex){@media only screen and (min-width: 620px){.fixed-width.has-border .layout__inner{display:flex !important}}}@media only screen and (-webkit-min-device-pixel-ratio: 2), only screen and (min--moz-device-pixel-ratio: 2), only screen and (-o-min-device-pixel-ratio: 2/1), only screen and (min-device-pixel-ratio: 2), only screen and (min-resolution: 192dpi), only screen and (min-resolution: 2dppx){.fblike{background-image:url(https://i7.createsend1.com/static/eb/master/13-the-blueprint-3/images/fblike@2x.png) !important}.tweet{background-image:url(https://i8.createsend1.com/static/eb/master/13-the-blueprint-3/images/tweet@2x.png) !important}.linkedinshare{background-image:url(https://i9.createsend1.com/static/eb/master/13-the-blueprint-3/images/lishare@2x.png) !important}.forwardtoafriend{background-image:url(https://i10.createsend1.com/static/eb/master/13-the-blueprint-3/images/forward@2x.png) !important}}@media (max-width: 321px){.fixed-width.has-border .layout__inner{border-width:1px 0 !important}.layout, .stack .column{min-width:320px !important;width:320px !important}.border{display:none}.has-gutter .border{display:table-cell}}.mso div{border:0 none white !important}.mso .w560 .divider{Margin-left:260px !important;Margin-right:260px !important}.mso .w360 .divider{Margin-left:160px !important;Margin-right:160px !important}.mso .w260 .divider{Margin-left:110px !important;Margin-right:110px !important}.mso .w160 .divider{Margin-left:60px !important;Margin-right:60px !important}.mso .w354 .divider{Margin-left:157px !important;Margin-right:157px !important}.mso .w250 .divider{Margin-left:105px !important;Margin-right:105px !important}.mso .w148 .divider{Margin-left:54px !important;Margin-right:54px !important}.mso .size-8, .ie .size-8{font-size:8px !important;line-height:14px !important}.mso .size-9, .ie .size-9{font-size:9px !important;line-height:16px !important}.mso .size-10, .ie .size-10{font-size:10px !important;line-height:18px !important}.mso .size-11, .ie .size-11{font-size:11px !important;line-height:19px !important}.mso .size-12, .ie .size-12{font-size:12px !important;line-height:19px !important}.mso .size-13, .ie .size-13{font-size:13px !important;line-height:21px !important}.mso .size-14, .ie .size-14{font-size:14px !important;line-height:21px !important}.mso .size-15, .ie .size-15{font-size:15px !important;line-height:23px !important}.mso .size-16, .ie .size-16{font-size:16px !important;line-height:24px !important}.mso .size-17, .ie .size-17{font-size:17px !important;line-height:26px !important}.mso .size-18, .ie .size-18{font-size:18px !important;line-height:26px !important}.mso .size-20, .ie .size-20{font-size:20px !important;line-height:28px !important}.mso .size-22, .ie .size-22{font-size:22px !important;line-height:31px !important}.mso .size-24, .ie .size-24{font-size:24px !important;line-height:32px !important}.mso .size-26, .ie .size-26{font-size:26px !important;line-height:34px !important}.mso .size-28, .ie .size-28{font-size:28px !important;line-height:36px !important}.mso .size-30, .ie .size-30{font-size:30px !important;line-height:38px !important}.mso .size-32, .ie .size-32{font-size:32px !important;line-height:40px !important}.mso .size-34, .ie .size-34{font-size:34px !important;line-height:43px !important}.mso .size-36, .ie .size-36{font-size:36px !important;line-height:43px !important}.mso .size-40, .ie .size-40{font-size:40px !important;line-height:47px !important}.mso .size-44, .ie .size-44{font-size:44px !important;line-height:50px !important}.mso .size-48, .ie .size-48{font-size:48px !important;line-height:54px !important}.mso .size-56, .ie .size-56{font-size:56px !important;line-height:60px !important}.mso .size-64, .ie .size-64{font-size:64px !important;line-height:63px !important}/*]]>*/</style><style type="text/css">body{background-color:#fff}.logo a:hover,.logo a:focus{color:#1e2e3b !important}.mso .layout-has-border{border-top:1px solid #ccc;border-bottom:1px solid #ccc}.mso .layout-has-bottom-border{border-bottom:1px solid #ccc}.mso .border,.ie .border{background-color:#ccc}.mso h1,.ie h1{}.mso h1,.ie h1{font-size:32px !important;line-height:40px !important}.mso h2,.ie h2{}.mso h2,.ie h2{font-size:30px !important;line-height:38px !important}.mso h3,.ie h3{}.mso .layout__inner,.ie .layout__inner{}.mso .footer__share-button p{}.mso .footer__share-button p{font-family:Avenir,sans-serif}</style><meta name="robots" content="noindex,nofollow"><meta property="og:title" content="My First Campaign"></head> <!--[if mso]><body class="mso"> <![endif]--> <!--[if !mso]><!--><body class="no-padding" style="margin: 0;padding: 0;-webkit-text-size-adjust: 100%;"> <!--<![endif]--><table class="wrapper" style="border-collapse: collapse;table-layout: fixed;min-width: 320px;width: 100%;background-color: #fff;" role="presentation" cellspacing="0" cellpadding="0"><tbody><tr><td><div><div class="layout one-col fixed-width stack" style="Margin: 0 auto;max-width: 600px;min-width: 320px; width: 320px;width: calc(28000% - 167400px);overflow-wrap: break-word;word-wrap: break-word;word-break: break-word;"><div class="layout__inner" style="border-collapse: collapse;display: table;width: 100%;background-color: #ffffff;"> <!--[if (mso)|(IE)]><table align="center" cellpadding="0" cellspacing="0" role="presentation"><tr class="layout-fixed-width" style="background-color: #ffffff;"><td style="width: 600px" class="w560"><![endif]--><div class="column" style="text-align: left;color: #021a2c;font-size: 16px;line-height: 24px;font-family: Avenir,sans-serif;"><div style="Margin-left: 20px;Margin-right: 20px;"><div style="mso-line-height-rule: exactly;line-height: 23px;font-size: 1px;">&nbsp;</div></div><div style="Margin-left: 20px;Margin-right: 20px;"><div style="mso-line-height-rule: exactly;mso-text-raise: 11px;vertical-align: middle;"><h2 style="Margin-top: 0;Margin-bottom: 0;font-style: normal;font-weight: normal;color: #021a2c;font-size: 26px;line-height: 34px;"><strong>CHECK-IN</strong></h2><p style="Margin-top: 16px;Margin-bottom: 20px;">Boa tarde caro(a) participante,<br> <br> Está tudo pronto para o início da Hackathon HeartBits! Já podes aderir ao Discord, a nossa plataforma principal! Basta clicares no botão abaixo ou acederes a <a style="text-decoration: underline;transition: opacity 0.1s ease-in;color: #222d8f;" href="https://discord.gg/6djqZbB">este link</a> e seguir as indicações que se encontram na plataforma.</p></div></div><div style="Margin-left: 20px;Margin-right: 20px;"><div class="btn btn--flat btn--large" style="Margin-bottom: 20px;text-align: center;"> <!--[if !mso]--><a style="border-radius: 4px;display: inline-block;font-size: 14px;font-weight: bold;line-height: 24px;padding: 12px 24px;text-align: center;text-decoration: none !important;transition: opacity 0.1s ease-in;color: #ffffff !important;background-color: #225053;font-family: Avenir, sans-serif;" href="https://discord.gg/6djqZbB">Entrar no Discord</a><!--[endif]--> <!--[if mso]><p style="line-height:0;margin:0;">&nbsp;</p><v:roundrect xmlns:v="urn:schemas-microsoft-com:vml" href="https://discord.gg/6djqZbB" style="width:186px" arcsize="9%" fillcolor="#225053" stroke="f"><v:textbox style="mso-fit-shape-to-text:t" inset="0px,11px,0px,11px"><center style="font-size:14px;line-height:24px;color:#FFFFFF;font-family:Avenir,sans-serif;font-weight:bold;mso-line-height-rule:exactly;mso-text-raise:4px">Entrar no Discord</center></v:textbox></v:roundrect><![endif]--></div></div><div style="Margin-left: 20px;Margin-right: 20px;"><div style="mso-line-height-rule: exactly;mso-text-raise: 11px;vertical-align: middle;"><p class="size-10" style="Margin-top: 0;Margin-bottom: 20px;font-size: 10px;line-height: 18px;" lang="x-size-10"><span style="color:#707070">Copia este link para o teu browser se o botão não funcionar: https://discord.gg/6djqZbB</span></p></div></div><div style="Margin-left: 20px;Margin-right: 20px;"><div style="mso-line-height-rule: exactly;mso-text-raise: 11px;vertical-align: middle;"><p style="Margin-top: 0;Margin-bottom: 0;">O teu código pessoal, que vais necessitar para fazer o check-in, é:</p><p class="size-20" style="Margin-top: 20px;Margin-bottom: 0;font-size: 17px;line-height: 26px;text-align: center;" lang="x-size-20"><strong>'
var HTML_BODY_2 = '</strong></p><p style="Margin-top: 20px;Margin-bottom: 20px;">Deixamos a ressalva de que as equipas podem vir a sofrer ligeiros ajustes até ao momento do início da atividade, em virtude de eventuais desistências. As equipas que se inscreveram já completas&nbsp; excetuam-se à ressalva anterior!<br> <br> Estamos à tua espera no Discord! Junta-te a nós!!<br> <br> Até já,<br> A equipa organizadora da Hackathon HeartBits 2020</p></div></div><div style="Margin-left: 20px;Margin-right: 20px;"><div style="mso-line-height-rule: exactly;line-height: 9px;font-size: 1px;">&nbsp;</div></div></div> <!--[if (mso)|(IE)]></td></tr></table><![endif]--></div></div><div style="mso-line-height-rule: exactly;line-height: 20px;font-size: 20px;">&nbsp;</div><div style="background-color: #ebebeb;"><div class="layout one-col stack" style="Margin: 0 auto;max-width: 600px;min-width: 320px; width: 320px;width: calc(28000% - 167400px);overflow-wrap: break-word;word-wrap: break-word;word-break: break-word;"><div class="layout__inner" style="border-collapse: collapse;display: table;width: 100%;"> <!--[if (mso)|(IE)]><table width="100%" cellpadding="0" cellspacing="0" role="presentation"><tr class="layout-full-width" style="background-color: #ebebeb;"><td class="layout__edges">&nbsp;</td><td style="width: 600px" class="w560"><![endif]--><div class="column" style="text-align: left;color: #021a2c;font-size: 16px;line-height: 24px;font-family: Avenir,sans-serif;"><div style="Margin-left: 20px;Margin-right: 20px;"><div style="mso-line-height-rule: exactly;line-height: 20px;font-size: 1px;">&nbsp;</div></div><div style="Margin-left: 20px;Margin-right: 20px;"><div style="mso-line-height-rule: exactly;line-height: 3px;font-size: 1px;">&nbsp;</div></div><div style="font-size: 12px;font-style: normal;font-weight: normal;line-height: 19px;" align="center"> <a style="text-decoration: underline;transition: opacity 0.1s ease-in;color: #222d8f;" href="https://heartbits.pt/"><img style="border: 0;display: block;height: auto;width: 100%;max-width: 134px;" alt="Heartbits" src="cid:logo" width="134"></a></div><div style="Margin-left: 20px;Margin-right: 20px;Margin-top: 20px;"><div style="mso-line-height-rule: exactly;line-height: 1px;font-size: 1px;">&nbsp;</div></div></div> <!--[if (mso)|(IE)]></td><td class="layout__edges">&nbsp;</td></tr></table><![endif]--></div></div></div></div></td></tr></tbody></table></body></html>';

var LOGO = UrlFetchApp.fetch('https://heartbits.pt/assets/banner.png').getBlob().setName("LOGO");

/**
 * Sends non-duplicate emails with data from the current spreadsheet.
 */
function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2; // First row of data to process
  var numRows = 80; // Number of rows to process
  // Fetch the range of cells A2:B3
  var dataRange = sheet.getRange(startRow, 1, numRows, 4);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var name = row[0];
    var emailAddress = row[1];
    var randomCode = row[2];
    var emailSent = row[3];
    if (emailSent !== EMAIL_SENT) { // Prevents sending duplicates 
      var body = BODY_1 + randomCode + BODY_2;
      var html_body = HTML_BODY_1 + randomCode + HTML_BODY_2;
      var message = {
        to: emailAddress,
        subject: EMAIL_SUBJECT,
        body: body,
        htmlBody: html_body,
        inlineImages: {logo: LOGO}
      };
      MailApp.sendEmail(message);
      sheet.getRange(startRow + i, 4).setValue(EMAIL_SENT);
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  }
}
