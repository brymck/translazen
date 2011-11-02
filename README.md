translazen
==========

Add functionality for translators to PowerPoint. Designed to increase in
usefulness as the design and wording choices of the origin authors increase in
awfulness.

Installation 
------------

If you haven't already, sigh as you boot up Windows, then get
[Git](http://help.github.com/win-set-up-git/). Make sure PowerPoint is closed.

For Windows, open the command prompt and copy-and-paste the following:

    git clone git://github.com/brymck/translazen.git
    copy /y "translazen\translazen.ppa" "%AppData%\Microsoft\Addins"

For Office for Mac, use this:

    git clone git://github.com/brymck/translazen.git
    translazen/install

If it gives you a "permission denied" message, most likely you still have an
instance of PowerPoint running. If you can verify that's not the case, there's
something wacky going on with user permissions.

Open PowerPoint and select the add-in:

* Windows 2007+: Office Button > PowerPoint Options > Add-Ins > Select
  "PowerPoint Add-ins" in the Manage drop-down > Go... > Add New... > Select
  translazen > OK > Close
* Earlier/Mac: Tools > Add-ins > Add New... > Select translazen > OK > Close

Updating
--------

Go to the original directory called `translazen` where you first downloaded
this repository and run the following commands separately:

    git pull origin master
    copy /y "translazen.ppa" "%AppData%\Microsoft\Addins"

Or for Macs:

    ./update

Happy (or at least less miserable) editing!

- Bryan McKelvey
