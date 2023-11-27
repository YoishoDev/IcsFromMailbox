*Read this in other languages: [Deutsch](README.de.md)*

# IcsFromMailbox
Load a calendar (events) from a Microsoft Exchange system and convert it into an iCalendar file (PowerShell).

## Idea and background
For a WordPress installation, a calendar from Microsoft Exchange 2019 (on-premises) should be displayed using a [plug-in](https://de.wordpress.org/plugins/ics-calendar/).
For this, the plug-in requires a URL to an iCalendar file. Since Exchange 2019 doesn't offer the option to publish a calendar via a URL, I thought about
Export events from a calendar in the Exchange system into an ICS file and simply store them in the directory of the WordPress installation.
Since the WordPress installation runs on a Microsoft Windows system, it should, if possible, be a PowerShell script so that additional environments do not have to be installed.
Although I found a PHP variant, PHP is installed to create an ICS file from a CSV file. However, I would still have needed something else to capture the events from that
Load Exchange system.
Since I didn't find anything suitable (as a whole), here's a variant with combinations of different things that I found on the internet.

- Export the events to a CSV file using a PS cmdlet from [here](https://github.com/David-Barrett-MS/PowerShell-EWS-Scripts).
- Convert to an ICS file
- Save the file in the WordPress installation

I spent a long time ensuring that the generated ICS file passed the validity check. The contents were always ok, only the file "had something" at the beginning.
To this day I don't have a proper explanation for the behavior :confused:.

Perhaps the scripts will serve as a suggestion for someone to solve similar "problems" and thus save them from long and possibly unsuccessful searches.

## Wishes, criticism, errors, comments, suggestions for improvement...

Bring it on! :+1: Please write everything to me.
Either you use the same [page](https://github.com/YoishoDev/IcsFromMailbox/issues) or\
You write me an <a href="mailto:development@yoisho.de">E-mail<a>.
