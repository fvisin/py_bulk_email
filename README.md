A python script to send emails to multiple addresses from an 'xls' or '.xlsx'
list.

THE SOFTWARE IS PROVIDED AS IS AND **NO SUPPORT** WILL BE GIVEN.

## Installation
This script depends on `pyexcel_xls`. 
On Mac or Linux, from a terminal execute (you might need to install pip 
before):

```bash
pip install --user pyexcel_xls
```

## Use
See the html in py_bulk_email.xlsx for an example of the usage of this tool.
The script allows you to add inline images just by:
    * Copying the images into the inline_images directory (take note of their
      alphabetical order!) 
    * Writing the usual html code for images <img src="cid:image#"> replacing
      the # with an increasing image number that follows the alphabetical order
      of the images.

You can easily put placeholders in the html with curly braces, that will be
replaced by the corresponding field in the Contacts sheet.

Finally, you can specify some information that will be used to build the email
in the Email content sheet:
    * **subject:** the subject of the emails 
    * **primary email field:** the email that will be used as first option
    * **secondary email field":** the email that will be used if the primary
      email field is empty

Of course you can also add attachments! To do so you just need to put them in
the attachments directory.
