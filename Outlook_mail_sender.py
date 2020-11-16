import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = '<to_mail_id>'
mail.Subject = 'Message subject'
mail.Body = 'Message body'
mail.HTMLBody = '<h2>HTML Message body</h2>'
attachment  = "Path to the attachment"
mail.Attachments.Add(attachment)

mail.Send()