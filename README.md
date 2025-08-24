# InvoiceSender
Gets a list of people's email addressed and a PDF with a collection of their respective invoices and sends the bills accordingly.

Program plan:
* 2 inputs:
    * **xls** file with house name, person name, email address
    * **PDF** file with multiple invoices grouped together, each corresponding to a certain person
* Solution:
    * Read the xls file
    * Get the person's email address
        * Create a local database and later update it
        * There might be 2 email addresses, send to both
    * Read the PDF file
    * Extract each file in the PDF file
    * Access my email tool from this program
    * Send the correct file to the correct email address (local database)
        * Local database can be a dict at first, later a better database


Notes:
* The address in the PDF must match the one provided in the client's table