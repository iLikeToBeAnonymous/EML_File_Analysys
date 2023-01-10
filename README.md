# EML File Analyzer

## About
### _Purpose_
This is a small collection of scripts created in the attempt to analyze a large quantity of `.eml` files downloaded from one account on Thunderbird.
The majority of the `.eml` files are assumed to be responses to messages originating from the email account which is being accessed via Thunderbird.
Some messages, however, did not originate from the client email, but were from either an email address directly, a response from the mail server of the address which received the emails,
or from the mail server of the outbound emails.

## Goals
Analyze the many `.eml` files (test data set was on approximately 50k files) and extract and summarize the following information:
- Frequency in which any given email address appeared in the data set
- Nature of the message received (e.g., bounced, delayed, failed, etc.)
- If an error/warning message has the sender listed as a mail server, look elsewhere in the message contents to identify the original email address having the problem.

### _Method_
