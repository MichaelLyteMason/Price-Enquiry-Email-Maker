# Price-Enquiry-Email-Maker
Generate bespoke emails for multiple contacts at multiple suppliers enquiring about products you intend to find the lowest price for and use in your project or sell onto your customer.

## Program Use Cases
If you intend to do a project for a customer and need to quote them on cost of materials e.g you're a plumbing business and have a list of supplies you need for a project. To get the best possible price you will likely have to contact multiple suppliers bespokely. This program takes the list of items you need, what quantities and even what prices you last paid (to get your supplier to match or beat them), allows you to select a list of contacts from an editable list of contacts at various suppliers then produces several emails (as seen in the code) like:

Hi [contact_name],

I hope you're well. I'm emailing to enquire on prices for the following products:
                        

[product_list]

                        
Many thanks,
YOUR NAME

With each email being altered for each supplier contact, then at the end giving the list of supplier contact details so you can copy the emails individually and send emails to each supplier.

## Dependencies
- tkinter
- csv
- os
- pyperclip

## Set Up
1. Run the python file once in its own virtual environment.
2. In the VE it will automatically create the suppliers and email template files.
3. In the email template file you can alter the format to your liking.
4. Now you can use the program freely.
