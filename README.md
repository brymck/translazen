pp\_yaku\_zen
=============

Add functionality for translators to PowerPoint. Designed to increase in
usefulness as the design and wording choices of the origin authors increase in
awfulness.

Installation 
------------

If you haven't already, sigh as you boot up Windows, then get
[Git](http://help.github.com/win-set-up-git/). Make sure PowerPoint is closed.

For Windows, open the command prompt and copy-and-paste the following:

    git clone git://github.com/brymck/pp_yaku_zen.git
    copy /y "pp_yaku_zen\pp_yaku_zen.ppa" "%AppData%\Microsoft\Addins"

For Office for Mac, use this:

    git clone git://github.com/brymck/pp_yaku_zen.git
    pp_yaku_zen/install

If it gives you a "permission denied" message, most likely you still have an
instance of PowerPoint running. If you can verify that's not the case, there's
something wacky going on with user permissions.

Open PowerPoint and select the add-in:

* Windows 2007+: Office Button > PowerPoint Options > Add-Ins > Select
  "PowerPoint Add-ins" in the Manage drop-down > Go... > Add New... > Select
  pp\_yaku\_zen > OK > Close
* Earlier/Mac: Tools > Add-ins > Add New... > Select pp\_yaku\_zen > OK > Close

Updating
--------

Go to the original directory called `pp_yaku_zen` where you first downloaded
this repository and run:

    git pull
    copy /y "pp_yaku_zen\pp_yaku_zen.ppa" "%AppData%\Microsoft\Addins"

Or for Macs:

    git clone git://github.com/brymck/pp_yaku_zen.git
    pp_yaku_zen/install

Happy (or at least less miserable) editing!

- Bryan McKelvey
