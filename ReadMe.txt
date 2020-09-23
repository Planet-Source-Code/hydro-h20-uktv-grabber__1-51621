UKTVGrabber By CornStopper
~~~~~~~~~~~~~~~~~~~~~~~~~~

This program grabs ut tv listing information from
the net. The speed it grabs the information is dependant
on your connection speed, the number of channels
you are grabbing and the number of days.

For example, it took 5.5 mins to collected 3 days worth of
information for one channel at 48kbps.

The output the grabbed data is in XML format.

USAGE
~~~~~

UKTVGrabber is command line driven. If you are using a
dial up modem, then add to then of each line MODEM. If this
parameter if given then the application will automatically
connect to then net and then dis-connect.


The first thing you have to do is configure what
channels you require. This is done by the following commmand:-

	UKTVGrabber -configure
		or
	UKTVGrabber -configure MODEM

This will connect to then net and retrieve the available channels.
Then just select which channels you require, and apply the changes.
You can alter which channels you want to collect by using the following
command:-

	UKTVGrabber -channels

After you have configured which channels to collect then lets get them.
This is done by using the '-grab' command.

You could just use the following:-

	UKTVGrabber -grab
		or
	UKTVGrabber -grab MODEM


This would use the default settings, which are to grab two days worth
of listings and output the result to 'listings.xml' at the application
path.

You can use the following to overide this:-

	UKTVGrabber -grab days=3
	UKTVGrabber -grab xml=c:\xmltv\uklist.xml
	UKTVGrabber -grab days=3 xml=c:\xmltv\uklist.xml

Please note the xml path must not contain spaces, 
e.g. xml=c:\program files\xmltv\uklist.xml
Also the days parameter must be between 1 and 7.

Any problems just PM at modshack