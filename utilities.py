
# import general libraries
import os
import logging

# import specific libraries
import win32com.client as client

def send_mail_via_com(html_body, subject, recipient, on_behalf_of,                      
                      outlook_object=client.Dispatch("Outlook.Application"),
                      cc="", attachments=[], have_replies_sent_to=[],
                      read_receipt=False, delivery_receipt=False,
                      autosend=False, save_path=None, **keyword_parameters):
    """Send email using com object for Microsoft Outlook
	
	Keyword arguments:
	Required:
	html_body - string which contains the body of the email in valid HTML code
	subject - string containing the subject of the email
	recipient - string containing the email address(es) of the desired recipients, delimited by semi-colin
	on_behalf_of - string containing the email address of the sender of the email
	outlook_object (default client.Dispatch) - the com object which is used to send the email
	
	Optional:
	cc - string containing email address(es) of desired people to be carbon copies to the email, delimited by semi-colin
	attachments - list containing full paths to files to attach to the email
	have_replies_sent_to - list of email addresses which will be send replies to the email
	delivery_receipt - boolean indicating if the originator should recieve a delivery report
	autosend - boolean indicating if the email should be automatically sent or not
	save_path - string containing full path to where the email should be saved as a .msg file
    """
    msg = outlook_object.CreateItem(0)
    msg.SentOnBehalfOfName = on_behalf_of
    msg.To = recipient
    msg.Subject = subject
    msg.CC = cc
    msg.ReadReceiptRequested = read_receipt
    msg.OriginatorDeliveryReportRequested = delivery_receipt
    msg.HtmlBody = html_body

    for item in have_replies_sent_to:
        msg.ReplyRecipients.Add(have_replies_sent_to)
        
    for attachment in attachments:
        msg.Attachments.Add(attachment)

    if save_path is not None:
        msg.SaveAs(save_path)
    else:
        if not(autosend):
            raise ValueError("!!Please specifiy save_path or turn on\
                             autosend!!")
    # send emails if autosend is True
    if autosend:
        msg.send


