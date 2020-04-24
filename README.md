## CovEd AutoMailing

This repository contains the code for auto informing and mailing mentors and mentees of their matchings.

### Usage

Run the following pip command to get the Google OAuth library.
```
pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
```

Run main.py after installing the OAuth Google libraries. A window will open where you will have to login to the coved account. After that it will auto update and mail the mentors.
Please do not share the pickle file generated after logging in as it will cause breach of security.

You can change the email text by simply changing email.txt and using the appropriate template tags.


