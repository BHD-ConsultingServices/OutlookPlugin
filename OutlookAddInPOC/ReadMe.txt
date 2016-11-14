Outlook Plugin Spike

This spike demonstrates how to create a simple plugin into Microsoft Outlook. In the plugin, we added two buttons. One that reads the email with all its contents. The second button is just a sample. 

How to launch and operate:
1.	Ensure your Outlook is closed.
2.	Open the solution and start it.
3.	It will launch Outlook with the plugin already installed.
4.	Open an email.
5.	You will see a new ribbon on top named “Add Tab Name Here”.
6.	Once you click on it you will see the two buttons. 
7.	The first one will show you how to access the email contents any attachment. You can also access attachments inside an attached email with recursion.
8.	Have fun!

Known Problems
1.	Only works on Outlook 2013 & 2016. 
2.	We also encountered issues in 2013. 
3.	The certificate sometimes gives issues. Just install it under Trusted. Any work arounds here are welcome.

Release Strategy
1.	We deployed this via click-once. Just run the setup on the desired PC and every time Outlook is opened it check for a new update. 
2.	Update are seamless then and the user don’t even see it. 

Security
1.	Because this application might live on remote machines, we don’t save any connection strings on anything on it. 
2.	All data is passed to a web service.

Best Regards 
Pretor Team
