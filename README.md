# woocomerse_product_variable_import
This python script can be used to create the correct format to match the import file for woocommerce.

This is a python script for taking a list of products to be imported into woocommerce.

The best way to use this script is first to add a few products into Woocommerce as you want them to be - and then use the default export product button in products to get the layout you need to follow for the import.

This will provide you will a CSV file to use as your base. When you import all your products into Woocommerse, this will be the format to use for your import.

##WHAT DOES THIS SCRIPT DO?##
I noticed there were many expensive plugins for WordPress that did an 'ok' job at importing but struggled with more complicated product setups, like with variable products.

I created this script to solve this problem for myself, and I wanted to share it with others who may benefit.

To get the best results from this script you will have to know some Python code, and, most likey, edit some of the code to adapt to your specific needs.

This script does not modify all woocommerce columns. The column structure I used was as follows and then I copied and pasted what was needed to fit the exact format of the import file for woocommerce.

Each title is a column starting at "Name" which is column 1 or in excel column A

Name   Name_attribute_To_append_to_name   Quantity   Units  Price 1   Quantity   Units  Price 2   Quantity   Units  Price 3  Regular price  Attribute 1 value(s)   Meta: metadata_name    Meta: metadata_name_2  Meta: metadata_name_3  Type   Meta: _wp_page_template    Tax class  Attribute 1 visible Image SKU    Parent

If there is interest, I could rebuild it to line up perfectly with a standard woocommerce export file - but I built this for a custom project, and most likey you may have a client product file which does not line up with the export file from woocommerse.
