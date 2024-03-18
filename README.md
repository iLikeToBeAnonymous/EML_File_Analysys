# EML File Analyzer

## About
### _Purpose_
This is a small collection of scripts created in the attempt to analyze a large quantity of `.eml` files downloaded from one account on Thunderbird.
The majority of the `.eml` files are assumed to be responses to messages originating from the email account which is being accessed via Thunderbird.
Some messages, however, did not originate from the client email, but were from either an email address directly, a response from the mail server of the address which received the emails,
or from the mail server of the outbound emails.

### _How to Use_
- Using Thunderbird, export all questionable emails as `.eml` files to the "ThunderbirdExports" folder.
- Run the Python file `Eml_error_and_email_compilation.py`

### _Points of Interest_
Because my Python was rusty when I began this project, I tried using ChatGPT as a helper. However, most of the code has been modified by me after-the-fact.

### _Goals_
Analyze the many `.eml` files (test data set was on approximately 50k files) and extract and summarize the following information:
- Frequency in which any given email address appeared in the data set
- Nature of the message received (e.g., bounced, delayed, failed, etc.)
- If an error/warning message has the sender listed as a mail server, look elsewhere in the message contents to identify the original email address having the problem.

---

## Misc (needs organized)
At this stage, the primary script used is `Eml_error_and_email_compilation.py`.

My Python scripts (so far) rely heavily on the `email` module, specifically [`email.message.Message`](https://docs.python.org/3/library/email.compat32-message.html).  
The newer form is [`email.message`](https://docs.python.org/3/library/email.message.html).

ChatGPT recommended the regex `r'[\w\.-]+@[\w\.-]+'`, but I found that in some cases, it would match to
non-addresses that happened to contain an "`@`" symbol. Therefore I fell back to my own regex and compiled
it as a global variable in the main script:

```Python
emailRegEx = re.compile("[a-zA-Z0-9_\.\-\+]{1,}@[a-zA-Z0-9\-]+\.[a-zA-Z0-9]{2,4}")
```
This has the (supposed) added advantage of being more efficient, but I don't know if this is true with modern Python versions.

### _Difference between `json.dump()` and `json.dumps()`_
While this is obvious in hindsight, I think this is worth putting here in case I (or anyone else) needs a reminder.
When I asked this of ChatGPT (Jan 9, 2023 version), I got the following answer:

> _`json.dump()` and `json.dumps()` are two methods provided by the `json` module in Python for working with JSON data. The main difference between them is in how they handle the data:_

> - `json.dump(obj, fp, *, skipkeys=False, ensure_ascii=True, check_circular=True, allow_nan=True, cls=None, indent=None, separators=None, default=None, sort_keys=False, **kw)`  
>_This method writes a Python object obj to a file-like object fp (such as a file opened in write mode) in JSON format._
>
> - `json.dumps(obj, *, skipkeys=False, ensure_ascii=True, check_circular=True, allow_nan=True, cls=None, indent=None, separators=None, default=None, sort_keys=False, **kw)`  
> _This method converts a Python object `obj` to a JSON formatted string._

> _In summary, `json.dumps()` is used to convert a Python object to a JSON string, while `json.dump()` is used to write a Python object to a file-like object in JSON format._

> _The other arguments in `json.dump()` and `json.dumps()` are used for more specific purposes such as setting the indentation for pretty-printing of json, allow or disallow certain types of data in json and more, you can check the documentation for more details._

### _Multi-part Messages_
I found out that warning and error messages are oftentimes what are called "multipart" messages. This can be seen by looking at
the topmost `Content-Type` field in a `.eml` file. If a message is multipart, it seems that the Python email module does not look
at fields that appear after the first occurrence of `Content-Type`. In my testing, the only way to extract all additonal fields is by using
the `.walk()` method.
