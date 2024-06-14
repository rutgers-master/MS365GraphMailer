# MS365GraphMailer

MS365GraphMailer is a Python script designed to send emails using the Microsoft Graph API. It uses an Entra ID (Azure AD) App Registration to authenticate and authorize email sending.

## Requirements
* [Python](https://www.python.org/)
  * [Requests](https://pypi.org/project/requests/)

On Windows after installing Python:
```console
python pip install requests
```

On RHEL 8.x Linux:
```console
yum install python36 python3-requests
```
On Debian/Ubuntu Linux:
```console
apt install python3 python3-requests
```

## Getting Started
Ready to get started? Clone the repo, or just grab the `MS365GraphMailer.py` program.

Read the instructions at the top of the program which will walk you through setting up the Entra ID (Azure AD) App Registartion.  There are only 3 variables in the code that need to be modified:

* **CLIENT_ID**: The Application (client) ID from the App Registration.
* **CLIENT_SECRET**: The client secret from the App Registration.
* **TENANT_ID**: The Directory (tenant) ID from the App Registration.

You can then send a test message with something like:
```console
python MS365GraphMailer.py -f myfrom@address.here -t myto@address.here -s "Test Subject" -m "Test Body"
```
If you run into any errors, please double check the App Registration setup and the variables in the code.