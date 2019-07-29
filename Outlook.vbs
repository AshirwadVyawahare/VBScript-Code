Set conn = CreateObject("outlook.application")
set smail= conn.createitem(0)

with smail
.to = "Ashirwad.Vyawahare@synechron.com"
.subject = "Test mail with attachment"
.body = "Testing if attachment can be sent ;)"
.attachments.add "C:\Documents and Settings\Ashirwad.Vyawahare\Desktop\Test.vbs"
.display
end with

'smail.send


