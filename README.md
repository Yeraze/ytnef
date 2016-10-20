Yerase's TNEF Stream Reader
===========================

[![Build Status](https://travis-ci.org/Yeraze/ytnef.svg?branch=master)](https://travis-ci.org/Yeraze/ytnef)

Installation
------------

Simply run make, and make install.
On some systems you may need to edit the Makefile to change the CC= &
LINK= lines to use cc instead of gcc.

Usage
-----

`ytnef filename.dat` - parse the file & print some basic information.

`ytnef -v filename.dat` - parse the file & print EVERYTHING.  good for
        debugging.

`ytnef -f . filename.dat` - parse the file and save all of the attachments
        to the local directory.

`ytnef -h` - See these and more options and examples.

Automatic stripping of TNEF streams using procmail:
---------------------------------------------------

If you want to use ytnefprocess.pl, then make sure to do the following:

1. Make a directory `~/.ytnef`
2. Make a directory `~/.ytnef/output`
3. Install the MIME::Parser perl module.  You can do this with:

    ```
    perl -MCPAN -e shell
    (let it load)
    install MIME::Parser
    ```

4. Copy ytnefprocess.pl & ytnef to a place like /usr/bin.
5. Create a .procmailrc something like this:

    ```
    PATH=/bin:/usr/bin
    HOME=/home/randall
    :0fw
    * > 0
    | /usr/bin/ytnefprocess.pl
    ```

and Voilà! you should be ready to go.

Help
----

If you want additional information exported (to the vCalendar or vCard
information), then use the -v option to try & find out where the information
is.  Then send the winmail.dat file, along with the -v listing to me 
(randall.hand@gmail.com) along with some information & justifications.  It works 
best if you send several data files, showing the option in various states 
to back up your claim.  For example, if you want me to parse recurrence
patterns encoded in the TNEF stream, then send yourself several meeting
requests with almost identical information, and several different recurrence
patterns, and (hopefully) the only fields that change will the the UID & the
recurrence information.


Support
-------

Don't hesitate to send support requests to me at randall.hand@gmail.com. I'll do 
what I can to help.  Also don't hesitate to post bug reports or
support/feature requests to https://github.com/Yeraze/ytnef/issues/new

