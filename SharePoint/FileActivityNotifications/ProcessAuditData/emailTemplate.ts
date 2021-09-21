const template: string = `<!DOCTYPE html>
<html lang="en">

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <meta http-equiv="X-UA-Compatible" content="IE=11">
  <title>Microsoft</title>

  <style>
    table {
      border-collapse: collapse;
    }
    ul {
      padding: 0;
      margin: 0 0 0 24px;
      list-style-position: inside;
      font-family:'Segoe UI', Verdana, Geneva, sans-serif;
    }
    ul li {
      font-family:'Segoe UI', Verdana, Geneva, sans-serif;
      margin-bottom: 10px;
    }

    @media only screen and (max-width: 600px) {
      .column {
        width: 100% !important;
      }
    }
    @media only screen and (min-width: 0px) {
      ul {
        margin: 0 !important;
      }
    }
  </style>
</head>

<body link="#0078D4" vlink="#0078D4" alink="#0078D4"
  style="margin: 0; padding: 0; background-color: #E3E3E3; font-size: 11pt; font-family: 'Segoe UI', Verdana, Geneva, sans-serif; font-weight: normal; font-style: normal; line-height: normal; color: #000000;">
  <table role="presentation" width="100%" border="0" cellspacing="0" cellpadding="0"
    style="border: 0; border-collapse: collapse; width: 100%; background-color: #E3E3E3; font-size: 11pt; font-family: 'Segoe UI', Verdana, Geneva, sans-serif; font-weight: normal; font-style: normal; line-height: normal; color: #000000; text-align: left; mso-table-lspace: 0; mso-table-rspace: 0;">
    <tr>
      <td valign="top" align="center" style="border: 0">
        <table role="presentation" border="0" cellspacing="0" cellpadding="0"
          style="border: 0; border-collapse: collapse; margin: 0 auto; border-collapse: collapse; background-color: #E6E6E6; mso-table-lspace: 0; mso-table-rspace: 0;">

          <!-- row 1 -->
          <tr>
            <td width="643" align="center"
              style="border: 0; width: 643px; padding: 0; margin: 0; background-color: #FFFFFF;">
              <table role="presentation" width="100%" border="0" cellspacing="0" cellpadding="0"
                style="border: 0; border-collapse: collapse; width: 100%; padding: 0; margin: 0; background-color: #ffffff; font-family: 'Segoe UI', Verdana, Geneva, sans-serif; font-weight: normal; font-style: normal; line-height: 0; color: #000000; text-align: left; mso-table-lspace: 0; mso-table-rspace: 0;">
                <tr>
                  <td valign="top" style="border: 0; padding: 0; margin: 0;">

                    <!-- column 1 -->
                    <table role="presentation" class="column" width="100%" border="0" cellspacing="0" cellpadding="0"
                      style="border: 0; border-collapse: collapse; width: 100%; text-align: left; mso-table-rspace: 0; mso-table-lspace: 0; padding: 0; line-height: normal;">
                      <tr>
                        <td width="640" align="left" valign="top"
                          style="border: 0; width: 640px; background-color: #243a5e; margin: 0; padding: 20px 20px 40px 20px; color: #FFFFFF; border-collapse: collapse;">
                          <h1 style="font-family:'Segoe UI Semibold','Segoe UI', Verdana, Geneva, sans-serif; margin: 0px; padding:0px; font-weight: normal; mso-line-height: exactly; line-height: 150%; line-height: 35px;">SharePoint Audit Alert</h1>
                        </td>
                      </tr>
                      <tr>
                        <td width="640" valign="middle"
                          style="border: 0; width: 640px; height: 36px; vertical-align: middle; padding: 0 20px; margin: 0; background-color: #0078d4; font-size: 11.5pt; line-height: normal; font-family: 'Segoe UI Semibold', 'Segoe UI', Verdana, Geneva, sans-serif; color: #FFFFFF;">
                          File Notification
                        </td>
                      </tr>
                      <tr>
                        <td width="640" valign="top"
                          style="border: 0; font-size: 11pt; width: 100%; padding: 20px 20px 0 20px; margin: 0;">
                          <h2 style="font-family:'Segoe UI Semibold', 'Segoe UI', Verdana, Geneva, sans-serif; color:#000000; margin: 0; font-weight: normal;">Why am I receiving this notification?</h2>
                          <p style="font-family:'Segoe UI', Verdana, Geneva, sans-serif; margin: 20px 0 20px 0;">Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent vitae facilisis odio. Mauris rhoncus mauris vel diam imperdiet, rhoncus auctor metus pharetra. Sed porta risus ac mi lacinia sagittis. Etiam ornare libero sapien, et sagittis eros fringilla ut. Suspendisse imperdiet, neque nec condimentum facilisis, ante arcu vestibulum libero, eu suscipit velit orci eu purus. Aenean eu elit sit amet nulla ultricies sodales quis id ipsum. Sed quis sem sed nibh commodo finibus. Suspendisse auctor, urna in gravida rhoncus, mauris nulla luctus lorem, quis facilisis neque ante vel tellus.</p>

                          <h3 style="font-family:'Segoe UI', Verdana, Geneva, sans-serif; color:#000000; margin: 10px 0 0 0; padding: 0; font-size: 11pt;">Alert Details</h3>
                          <p style="font-family:'Segoe UI', Verdana, Geneva, sans-serif; margin-top: 10px;">
                            <table>
                              <tr>
                                <td style="font-weight: bold; padding-right: 15px;">Time</td>
                                <td>{{creationTime}}</td>
                              </tr>
                              <tr>
                                <td style="font-weight: bold; padding-right: 15px;">File</td>
                                <td>{{fileName}}</td>
                              </tr>
                              <tr>
                                <td style="font-weight: bold; padding-right: 15px;">User</td>
                                <td>{{Username}}</td>
                              </tr>
                              <tr>
                                <td style="font-weight: bold; padding-right: 15px;">Operation</td>
                                <td>{{Operation}}</td>
                              </tr>
                            </table>
                          </p>
                          
                        </td>
                      </tr>
                    </table>
                    <!-- /column 1 -->

                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <!-- /row 1 -->


          <!-- footer row -->
          <tr>
            <td width="640" align="center"
              style="border: 0; width: 640px; padding: 0; margin: 0; background-color: #FFFFFF; padding: 20px 0 0 0;">
              <table role="presentation" width="100%" border="0" cellspacing="0" cellpadding="0"
                style="border: 0; border-collapse: collapse; width: 100%; padding: 0; margin: 0; background-color: #F2F2F2; font-family: 'Segoe UI', Verdana, Geneva, sans-serif; font-weight: normal; font-style: normal; color: #000000; text-align: left; mso-table-lspace: 0; mso-table-rspace: 0;">
                <tr>
                  <td width="100%" valign="middle" align="left"
                    style="width: 100%; vertical-align:middle; text-align:left; padding: 0px;" height="60">
                    <img border="0" height="60" src="http://img-prod-cms-rt-microsoft-com.akamaized.net/cms/api/am/imageFileData/RE4FTEm?ver=314f"
                    alt="Microsoft footer logo" style="display:block; border-width:0;" />
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <!-- /footer row -->
        </table>
      </td>
    </tr>
  </table>
</body>

</html>`;

export default template;