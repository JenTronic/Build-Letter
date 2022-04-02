# Build Letter
Tool for generating labels for letters - using LibreOffice Writer and customer data from Prestashop.

![Build Letter Screenshot](/Screenshots/Build-Letter.png)

Requirements for running the Build Letter tool:

[PrestaShop 1.7](https://www.prestashop.com) <span style="color:gray;font-size:0.8em">*(Running with MySQL / MariaDB)*</span> <br />
[Windows PowerShell 7.0](https://aka.ms/powershell-release?tag=stable) <br />
[LibreOffice Writer 7.2](https://www.libreoffice.org/download/download) <br />
[MySQL Connector Net 8.0](https://dev.mysql.com/downloads/connector/net) <br /><br />

# How to use Build Letter?

First start is all about gaining access to your PrestaShop database, and preparing needed templates:<br /><br />
<ol>
<li>Run "Build Letter.ps1" using PowerShell.</li>
<li>Fill in the database information - for your PrestaShop database, and hit the "Save Settings" button.</li>
<li>Modify the LibreOffice Writer templates, located within the Templates folder.</li>
</ol>
<br />
Now Build Letter is ready for use, so for each label needed, all you do is:<br /><br />
<ol>
<li>Run "Build Letter.ps1" using PowerShell.</li>
<li>Enter a valid order ID, from your PrestaShop webshop.</li>
<li>Select which type of address to retrieve, and which template to use - and click the "Generate" button.</li>
<li>LibreOffice will now open the new label for final modification - i.e entering of the postage code.</li>
<li>Save the new label, and click the "Print Label" button.
</ol><br />

# Running code before and after printing 

Some printers may need special configuration, in order to print the labels - or perhaps you just need to run some code when printing.

If running additional code is needed, you may edit the "Functions.ps1" file, and create new functions, that will run automatically, when a specific printer is used:

Before-[printer_name]: <span style="color:gray;font-size:0.8em">*(Code is running before printing)*</span><br />
After-[printer_name]: <span style="color:gray;font-size:0.8em">*(Code is running after printing)*</span>

(Note that spaces " " in [printer_name] must be replaced with underscores "_". )