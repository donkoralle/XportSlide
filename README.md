# XportSlide
A PowerPoint Add-In (Office 2010 and up) to export the current slide in a high resolution (150, 300 &amp; 600 dpi) PNG-image.

Sometimes you want to share the stuff you produced in PowerPoint with others. Be it for a book our journal article, to hand on a figure to a collegue etc. To do so, taking a screenshot or using PowerPoints iternal export functionality might not give you the control over the quality of the exported slide you want. 

Since I ran quite often into such situations, I wrote a primitve Add-In to export the current slide to an PNG-image (to avoid JPG-like artefacts) with 150, 300, 600 dpi or any other resolution you like.

## Before you start

This Add-In is developed on a Win10 machine running MS Office 10 and 16. For this setup the Add-Ins should work without problems (hopefully).

And no: I did not test it on Mac Office etc. - good luck there ;-)

## How to use it

After [installing the Add-In](#how-to-install-it) the "Slide 2 PNG" button appears under the "Add-Ins" ribbon:

![How to use the Add-In](documentation/add_in_6.png)

Clicking this button activates an input box for defining the width of the exported slide. You can enter any width you like, widths for 150 (default), 300 and 600 dpi are shown ontop the inputbox. Hitting "OK" starts the export process. 

![How to use the Add-In](documentation/add_in_7.png)

After successfully writing the export file a confirmation dialog is shown:

![How to use the Add-In](documentation/add_in_8.png)

The exported image of the current slide is stored in the same folder as the presentation. To identifiy multiple exported slides the slide number is included in the filename. In this example I exported slides 1 and 2 from the presentation "Test.pptx":

![How to use the Add-In](documentation/add_in_9.png)

## How to install it

First of all, download the Add-In (.ppam file) and store it somewhere on your harddrive. Technically, an Add-In can be executed from nearly every place on your computer. To have an oversight on your Add-Ins I suggest to copy/move the Add-In (.ppam file) in the Microsoft Office "AddIns" folder, which you should find here:

C:\Users\YOUR USERNAME\AppData\Roaming\Microsoft\AddIns

In case you can not see the "AppData" folder, you have to activate "View - Hidden items" in your explorer:

![How to install the Add-In](documentation/show_hidden.png)

After storing the Add-In (.ppam) file in the "AddIns" folder you can close your explorer, fire up PowerPoint and choose "File - Options - Add-Ins" (Sorry for the german interface):

![How to install the Add-In](documentation/add_in_1.png)

At the bottom of the options-window you will find the "Manage" ("Verwalten") list, where you select "PowerPoint Add-ins" and click "Go" ("Los").

![How to install the Add-In](documentation/add_in_2.png)

In the Add-Ins dialog box, click "Add new" ("Neu hinzufügen ...").

![How to install the Add-In](documentation/add_in_3.png)

The "Add New PowerPoint Add-In" dialog box should directly drop you in your personal Office "Addins" folder, so you should directly see the Add-In. Otherwise you can navigate to the location of the Add-In, select it and then click "OK".

![How to install the Add-In](documentation/add_in_4.png)

After that a security notice appears asking you wheater you want to activate the Add-In's Macros or not. Click "Enable Macros", and then "Close". Now the Add-In should show up in the list of activated Add-Ins. Click "Close" to finish the installation process.

![How to install the Add-In](documentation/add_in_5.png)




