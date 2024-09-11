#!/usr/bin/python3

try:
	# The only non-standard package we need is the 'requests' package.
	import requests
except:
	print("We need the 'requests' package: pip install requests")
	raise SystemExit
import json, argparse, sys, os, base64, re

# This Python script is designed to send emails using the Microsoft Graph API.
# It uses an Azure AD App Registration to authenticate and authorize email
# sending. The script requires the 'requests' package to function.

# The script defines a class MS365GraphMailer with methods to get an access
# token from Microsoft Graph API and to send an email. The email data is
# passed as a dictionary and can include fields such as 'from', 'to', 'cc',
# 'bcc', 'replyTo', 'subject', 'body', 'contentType', and 'saveToSentItems'.

# The script also includes a main function that sets up command line arguments
# for the email data, creates an instance of the MS365GraphMailer class, and
# sends the email.

# The script can be run from the command line with the appropriate arguments.
# It also supports reading the email body from stdin if the message argument
# is 'STDIN'.

# Steps you need to complete:
# 1. Set up an app registration in the Azure portal (documented below).
# 2. Set the CLIENT_ID, CLIENT_SECRET, and TENANT_ID values in the script.
# 3. Run the script from the command line with the appropriate arguments:
#    Run the script with -h or --help to see the available arguments.

# Steps to set up an app registration in the Azure portal:
# 1. Go to the Azure portal (https://portal.azure.com).
# 2. Go into the "Microsoft Entra ID" area (used to be Azure Active Directory).
# 3. Go to "App registrations".
# 4. Click on "+ New registration".
# 5. Fill in the required fields:
#	 - Name: Enter a name for your app registration.
#	 - Supported account types: Choose the appropriate account type (single if only one tenant, or multi if multitenant).
#    - Redirect URI: Not needed, because we are going to do "Application permissions", not "Delegated permissions".
#    NOTE: "Application permissions" are for when the app is acting on its own behalf (server side, no user interaction), while "Delegated permissions" are for when the app is acting on behalf of a user (user must allow app the first time it is used).
# 6. Click "Register".
# 7. On the "Overview" page for the app, jot down these two values:
#    - Application (client) ID
#    - Directory (tenant) ID
# 8. Go to "API permissions" and click on "+ Add a permission".
#    - Under "Microsoft APIs", choose "Microsoft Graph".
#    - Choose "Application permissions".
#    - Search for "Mail.Send" and select it.
#    - Search for "Mail.ReadWrite" and select it. [Only needed if you want sent emails to be saved in the Sent Items folder.]
#      The above two permissions are the only ones needed, if any "User.*" permissions are listed, you can remove them.
#    - Click the "Add permissions" button.
#    - Above the list of permissions, you must click the "Grant admin consent for <your tenant>" button.
#      When prompted to confirm, click "Yes".
# 9. Go to "Certificates & secrets" and under "Client secrets", click on "+ New client secret".
#	 - Add a description.
#	 - Choose an expiration period. [I recommend 24 months]
#	 - Click "Add".
# 10. Jot down the value of the client secret.
#     WARNING: Make sure you copy the "value" NOT the "Secret ID", the "value" is your client secret.
# 11. You can now use the "Application (client) ID", "Directory (tenant) ID", and "Client secret" in your script.
# 12. WARNING: You should set yourself up a calendar event to remind you to renew the client secret before it expires.

# When renewing the client secret, simply create a new client secret, and then update this script with that new value.
# You do not need to update the "Application (client) ID" or "Directory (tenant) ID" values, only the client secret.
# You can then delete the old (expiring/expired) client secret from the Azure portal.


# Set up the Azure AD App Registration values
CLIENT_ID = "CLIENT ID HERE"
CLIENT_SECRET = "CLIENT SECRET VALUE HERE"
TENANT_ID = "TENANT ID HERE"

class MS365GraphMailer:
	def __init__(self, client_id, client_secret, tenant_id):
		"""
		Initializes the MS365GraphMailer class.

		Args:
			client_id (str): The client ID of the Azure AD App Registration.
			client_secret (str): The client secret of the Azure AD App Registration.
			tenant_id (str): The tenant ID of the Azure AD App Registration.
		"""
		self.client_id = client_id
		self.client_secret = client_secret
		self.tenant_id = tenant_id
		self.access_token = None

	def get_access_token(self):
		"""
		Retrieves the access token from Microsoft Graph API using client credentials.

		Raises:
			requests.exceptions.HTTPError: If the request to retrieve the access token fails.
		"""
		data = {
			'grant_type': 'client_credentials',
			'client_id': self.client_id,
			'client_secret': self.client_secret,
			'scope': 'https://graph.microsoft.com/.default'
		}

		response = requests.post("https://login.microsoftonline.com/%s/oauth2/v2.0/token" % (self.tenant_id), data=data)
		response.raise_for_status()  # raise exception if invalid response
		self.access_token = response.json()['access_token']

	def send_email(self, data):
		"""
		Sends an email using the Microsoft Graph API.

		Args:
			data (dict): The email data to be sent:
				from (str): The email address of the sender.
				to (str or list): The email address(es) of the To recipient(s).
				cc (str or list): The email address(es) of the Cc recipient(s). (optional)
				bcc (str or list): The email address(es) of the Bcc recipient(s). (optional)
				replyTo (str): The email address to reply to. (optional)
				subject (str): The subject of the email.
				body (str): The body of the email.
				headers (dict): Additional headers to include in the email. (optional)
				attachments (list): List of file paths to attach to the email. (optional)
				attachments_inline (list): List of file paths to attach to the email inline at the bottom of the email. (optional)
				contentType (str): The content type of the email body 'Text' or 'HTML'. (optional, default: 'Text')
				saveToSentItems (bool): Whether to save the email in the Sent Items folder. (optional, default: True)

		Raises:
			requests.exceptions.HTTPError: If the request to send the email fails.
		"""

		# Convert all keys in data to lowercase key names
		data = {key.lower(): value for key, value in data.items()}

		# Let's make sure our data has atleast these fields: from, to, subject, body
		if not all(key in data for key in ['from', 'to', 'subject', 'body']):
			raise ValueError("Email data must contain 'from', 'to', 'subject', and 'body' fields.")

		# If contenttype not in data, then set to 'Text'
		if 'contenttype' not in data:
			data['contenttype'] = 'Text'

		# Let's make sure our contenttype is either 'Text' or 'HTML'
		data['contenttype'] = data['contenttype'].upper()
		if data['contenttype'] not in ['TEXT', 'HTML']:
			raise ValueError("Content type must be 'Text' or 'HTML'.")
		if data['contenttype']  == 'TEXT': data['contenttype'] = 'Text'

		# If saveToSentItems not in data, then set to True
		if 'savetosentitems' not in data:
			data['savetosentitems'] = True
		if data['savetosentitems']:
			data['savetosentitems'] = "true"
		else:
			data['savetosentitems'] = "false"
		
		# Set up the SendMail API endpoint
		# Documentation: https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0&tabs=http
		endpoint = "https://graph.microsoft.com/v1.0/users/%s/sendMail" % (data['from'])

		# Let's get the access token
		self.get_access_token()

		# Setup the headers
		headers = {
			'Authorization': 'Bearer ' + self.access_token,
			'Content-Type': 'application/json'
		}

		# Set up the basic message payload
		message = {
			"message": {
				"subject": data['subject'],
				"body": {
					"contentType": data['contenttype'],
					"content": data['body']
				},
			},
			"saveToSentItems": data['savetosentitems']
		}

		# Clean up the from and remove any spaces
		data['from'] = data['from'].replace(' ', '')

		# Let's setup the To, CC, BBCC, and ReplyTo fields
		if 'to' in data:
			if isinstance(data['to'], str): data['to'] = data['to'].replace(' ', '').split(',') # Convert to list if string
			message['message']['toRecipients'] = [{"emailAddress": {"address": recipient}} for recipient in data['to']]
		if 'cc' in data:
			if isinstance(data['cc'], str): data['cc'] = data['cc'].replace(' ', '').split(',') # Convert to list if string
			message['message']['ccRecipients'] = [{"emailAddress": {"address": recipient}} for recipient in data['cc']]
		if 'bcc' in data:
			if isinstance(data['bcc'], str): data['bcc'] = data['bcc'].replace(' ', '').split(',') # Convert to list if string
			message['message']['bccRecipients'] = [{"emailAddress": {"address": recipient}} for recipient in data['bcc']]
		if 'replyto' in data:
			if isinstance(data['replyto'], str): data['replyto'] = data['replyto'].replace(' ', '').split(',') # Convert to list if string
			message['message']['replyTo'] = {"emailAddress": {"address": data['replyto']}}

		# Add headers if they are set
		if 'headers' in data:
			message['message']['internetMessageHeaders'] = [{"name": key, "value": value} for key, value in data['headers'].items()]

		# Add attachments if they are set
		if 'attachments' in data:
			message['message']['attachments'] = []
			for attachment in data['attachments']:
				# Check if the file exists
				if not os.path.exists(attachment):
					raise FileNotFoundError("Attachment file '%s' does not exist." % (attachment))
				
				# Get the filename and file content
				filename = os.path.basename(attachment)
				with open(attachment, 'rb') as f:
					content = base64.b64encode(f.read()).decode('utf-8')

				# Add the attachment to the message
				contentType = "application/octet-stream"
				message['message']['attachments'].append({
					"@odata.type": "#microsoft.graph.fileAttachment",
					"name": filename,
					"contentBytes": content,
					"isInline": False,
					"contentType": contentType
				})

		# Add inline attachments if they are set
		if 'attachments_inline' in data:
			if 'attachments' not in message['message']:
				message['message']['attachments'] = []
			for attachment in data['attachments_inline']:
				# Check if the file exists
				if not os.path.exists(attachment):
					raise FileNotFoundError("Attachment file '%s' does not exist." % (attachment))
				
				# Get the filename and file content
				filename = os.path.basename(attachment)
				with open(attachment, 'rb') as f:
					content = base64.b64encode(f.read()).decode('utf-8')

				# Add the attachment to the message
				contentType = "application/octet-stream"
				# we assume file extension is an image extension like .jpg, .png, .gif, etc...
				contentType = "image/%s" % (os.path.splitext(filename)[1][1:].lower())
				contentId = re.sub(r'\W+', '', filename)
				message['message']['attachments'].append({
					"@odata.type": "#microsoft.graph.fileAttachment",
					"name": filename,
					"contentBytes": content,
					"isInline": True,
					"contentId": contentId,
					"contentType": contentType
				})
				# We assume that the body of the HTML attachment already has the image tag with the same contentId which is the file name with all non-alphanumeric characters removed, this can be uncommented if you just want to add the image to the end of the email body
				#message['message']['body']['content'] = "%s<img src=\"cid:%s\" alt=\"%s\" >" % (message['message']['body']['content'], contentId, contentId)


		# Make the API call
		response = requests.post(endpoint, headers=headers, data=json.dumps(message))

		# Print the results
		if response.status_code == requests.codes.accepted:
			print("MESSAGE SENT")
		else:
			print("Message sending failed with error code %s, %s" % (response.status_code, response.text))

def main():
	# Setup command line arguments
	parser = argparse.ArgumentParser(description='MS365 Graph Mailer', epilog="Send an email using Microsoft Graph API.\nBy default, the sent email is saved in the 'Sent Items' folder.")
	parser.add_argument('-f', '--fromaddr', type=str, help='The From address for the email', required=True)  # Can't use 'from', it is a reserved keyword
	parser.add_argument('-t', '--to', type=str, help='Comma separated list of To addresses', required=True)
	parser.add_argument('-s', '--subject', type=str, help='The subject of the email', required=True)
	parser.add_argument('-m', '--message', type=str, help='The message/body of the email (STDIN will use stdin)', required=True)
	parser.add_argument('-c', '--cc', type=str, help='Comma separated list of Cc addresses', required=False)
	parser.add_argument('-b', '--bcc', type=str, help='Comma separated list of Bcc addresses', required=False)
	parser.add_argument('-r', '--replyto', type=str, help='The address to set for the ReplyTo field', required=False)
	parser.add_argument('-H', '--header', action='append', type=str, help='Header in the format "Header1:Value1" (use of this argument for multiple headers)', required=False)
	parser.add_argument('-o', '--contenttype', type=str, help='Content type of message (default: Text)', required=False, default='Text', choices=['Text', 'HTML'])
	parser.add_argument('-a', '--attach', action='append', help='Path to the file to attach (use of this argument for multiple attachments)', required=False)
	parser.add_argument('-i', '--attach-inline', action='append', help='Path to the file to attach inline (use of this argument for multiple attachments)', required=False)
	parser.add_argument('-n', '--nosavetosent', action='store_true', help='Do not save sent message to "Sent Items" folder', required=False)

	# Parse command line arguments
	args = parser.parse_args()

	# Setup default values for contenttype and savetosentitems if they aren't set
	if not args.contenttype: args.contenttype = 'Text'
	if not args.nosavetosent: args.nosavetosent = False

	# If the message argument is 'STDIN', read the message from stdin
	if args.message.upper() == 'STDIN':
		args.message = sys.stdin.read()

	# Setup the email data
	email_data = {
		'from': args.fromaddr,
		'to': args.to,
		'subject': args.subject,
		'body': args.message.replace('\\n', '\n'),  # Replace \n with newline
		'contenttype': args.contenttype,
		'savetosentitems': not args.nosavetosent
	}

	# Check if cc, bcc, and replyto are set, and if so, add them to the email data
	if args.cc: email_data['cc'] = args.cc
	if args.bcc: email_data['bcc'] = args.bcc
	if args.replyto: email_data['replyto'] = args.replyto

	# Parse headers if they are set, and add them to the email data
	headers = {}
	if args.header:
		for h in args.header:
			temp_header = h.split(':', 1)
			if len(temp_header) == 2:
				temp_header[1] = temp_header[1].strip()
				headers[temp_header[0]] = temp_header[1]
			else:
				print("Invalid header format: %s" % (h))
				sys.exit(0)
			email_data['headers'] = headers

	# If attachments are set, add them to the email data
	if args.attach:
		email_data['attachments'] = args.attach
	if args.attach_inline:
		email_data['attachments_inline'] = args.attach_inline

	# Create an instance of the MS365GraphMailer class
	mailer = MS365GraphMailer(CLIENT_ID, CLIENT_SECRET, TENANT_ID)

	# Send the email
	mailer.send_email(email_data)

if __name__ == '__main__':
	main()
