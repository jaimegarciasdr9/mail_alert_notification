# -*- coding: utf-8 -*-
"""
Created on Thu Sep 14 12:35:47 2023
@author: Jaime
"""


def mail_alert_notification():
    import pandas as pd
    import os
    import database
    db = database.DatabaseLoad()
    conn = db.connect_to_database('informe_interno')

    # =============================================================================
    # Define email
    # =============================================================================
    def automatic_email_send(receiver, subject, message, file_attach):
        """
     Parameters
     ----------
     receiver: TYPE
        List of emails of the people to recieve this communication
     subject : TYPE
        Text that will be used as tittle for the email.
     message : TYPE
         String that will be shown in the body of the email.
     if_exists : TYPE
        If the table already exists "append".
        If the table has to be deleted "replace".

    Important information
    ---------------------
    .to_sql will create the table in the SQL server, therefore the column names of
     the dataframe will be the same column names you will have in the database.

     """
        import smtplib
        from email.mime.base import MIMEBase
        from email import encoders
        from email.mime.multipart import MIMEMultipart
        from email.mime.text import MIMEText

        server = '------'
        user = '-----'
        password = '-----'

        msg = MIMEMultipart()
        msg['Subject'] = subject
        msg['From'] = user
        msg['To'] = ', '.join(receiver)

        text_msg = MIMEText("""\
      <html>
        <body><span style="font-family:Arial, Helvetica, sans-serif">     
               <p>{}</p>
               <p></p>
               <p></p>
      </font>
      <head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
      <style>
      @import url('https://fonts.googleapis.com/css?family=Exo&display=swap');
      @import url('https://fonts.googleapis.com/css?family=Roboto+Condensed&display=swap');
      </style>
      </head>
      <body>
      <table cellpadding="0" cellspacing="0" border="0" style="background: none; border-width: 0px; border: 0px; margin: 0; padding: 0;">
      <tr>
      <td colspan="2" style=" color: #0f2b4c; font-size: 13px; font-weight: 600; font-family: Arial, Helvetica, sans-serif; padding-left: 5px;">Botty | <strong style="color: #eb0146;">Operational Bot Service</strong> </td>
      <tr>
      <td colspan="2" style=""><hr></td>
      </tr>
      <tr>
      <td style="color: #757474; font-size: 12px; font-family: Arial, Helvetica, sans-serif;  padding-left: 5px;">
      <a style=" color: #757474; text-decoration: none!important;" href="https://goo.gl/maps/E3Gggc59AbQzWFCE7" target="_blank">C/ Serrano, 18 bajo Izquierda, 28020 Madrid</a>
      </td>
      </tr>
      <tr>
      <td style="color: #757474; font-size: 12px; font-family: Arial, Helvetica, sans-serif;  padding-left: 5px;">
      <a style="color: #757474; text-decoration: none!important;" href="https://goo.gl/maps/2duVpdhUEJ2S2cpp7" target="_blank">C/ Serrano, 18 bajo Izquierda, 28020 Madrid</a>
      </td>
      </tr>
      <!-- Linkedin  -->
      <td rowspan="2" >
      <a href="https://es.linkedin.com/company/" nosend="1" target="_blank" title="LinkedIn"> <img src="https://..." alt="LinkedIn" style="display: inline-block; vertical-align: top; width: 30px;height: 30px;"></a>
      </td>
      </tr>
      <tr>
      <td style="color: #757474; font-size: 12px; font-family: Arial, Helvetica, sans-serif;  padding-left: 5px;">
      <a href="http://www.company.com" style=" color: #757474; text-decoration: none; font-weight: normal; font-size: 12px;">www.company.com</a>
      </td>
      </tr>
      </table>
      <!-- Footer Animado  -->
      <a href="http://www.company.com">
      <img src="https://..." nosend="1" alt="Logo " style="vertical-align: middle; width:310px; margin-top: 10px;">
      </a>
      <p style="max-width: 300px; font-size: 11px; padding-left: 5px; text-align: justify; color: grey;" >
      The content of this email is confidential and intended for the recipient specified in message only. It is strictly forbidden to share any part of this message with any third party, without a written consent of the sender. If you received this message by mistake, please reply to this message and follow with its deletion, so that we can ensure such a mistake does not occur in the future.
      </p>
      </body>
      </html>
      """.format(message), 'html')

        msg.attach(text_msg)

        part = MIMEBase('application', "octet-stream")
        part.set_payload(open("{}".format(file_attach), "rb").read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename="{}"'.format(file_attach))
        msg.attach(part)

        with smtplib.SMTP_SSL(server) as smtp:
            smtp.login(user, password)
            smtp.send_message(msg)
            smtp.quit()

    # Comprobar los usuarios que no tienen dni
    df_bbdd = pd.read_sql(
        "SELECT nombre_completo AS 'NOMBRE COMPLETO', email_empleado AS CORREO FROM dim_empleados_sesame WHERE dni_empleado = 'none'",
        conn)

    os.chdir(r'C:\Users\Administrator\...')

    if len(df_bbdd) > 0:
        writer = pd.ExcelWriter('usuarios_sin_dni.xlsx', engine='xlsxwriter')

        df_bbdd.to_excel(writer, sheet_name='Sin DNI', index=False)

        worksheet = writer.sheets['Sin DNI']

        worksheet.set_tab_color('red')
        worksheet.set_header('Sin DNI')

        writer.save()

        message = "<p>Buenos días Manuel,</p>\
                    <p></p>\
                    <p> Adjunto encontrará un excel con un listado de los usuarios que no tienen DNI.</p>\
                    <p></p>\
                    <p> Por favor asignale un DNI.</p>\
                    <p></p>\
                    <p> Por favor no responda a este mensaje, para dudas o aclaraciones dirijase a:</p>\
                    <p> marzo@gmail.com</p>"

        file_attach = 'usuarios_sin_dni.xlsx'

        automatic_email_send(["correoelectronico@gmail.com"], '#Google usuarios sin dni', message, file_attach)


mail_alert_notification()
