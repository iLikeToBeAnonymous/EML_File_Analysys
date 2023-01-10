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

## Misc (needs organized)
At this stage, the primary script used is `Eml_error_and_email_compilation.py`.

### _Multi-part Messages_
I found out that warning and error messages are oftentimes what are called "multipart" messages. This can be seen by looking at
the topmost `Content-Type` field in a `.eml` file. If a message is multipart, it seems that the Python email module does not look
at fields that appear after the first occurrence of `Content-Type`. In my testing, the only way to extract all additonal fields is by using
the `.walk()` method.
