# Extract-PowerApps-Sharing
Series of scripts and PBI dashboard to extract and view which users have had Power Apps shared with them

There are several steps to get this all working. I broke down the work required into steps (scripts), so steps can be repeated if necessary. 

1.	Step 1 – get a CSV file with the list of apps you want the user info for. I previously shared a KQL query for Azure Resource Graph Explorer, which gives the data we need. There is one issue with this though, it doesn’t return if a canvas app is premium or not. Therefore, we need to get this from the CoE toolkit Dataverse table which holds this information. Either we replace this KQL CSV totally and just use Dataverse table, OR, we use Dataverse table and join it to the KQL data. I suspect we can probably get away with just using Dataverse table (export to CSV/Excel).
2.	Step 2 – run PowerShell script (Get-App-Shared-List.ps1) against CSV from Step 1. This goes through each row (app) in the CSV and gets a list of users and groups the app is shared with. It writes this output to another CSV file, which is a list of users and groups each app is shared with. This script relies on the source CSV file having at least the following columns appId,modelDrivenAppId,appType,environmentId. The script also accepts these parameters:
a.	InputCsvPath – name of source CSV file
b.	OutputCsvPath – name of output CSV file
c.	EnvironmentId – Optional, if included it will only process apps for a specific environmentid. I think this is useful for testing and also if the script just takes too long to run, we can break it down by environment.
d.	IncludeModelApps – Optional, if excluded its default is true. i.e. do you want it to extract who model driven apps are shared with? If you include model apps, it takes longer because its more complex.
e.	IncludeManagedApps – Optional, and it must be set to false for now. So would just leave this param out.
f.	IncludePause – Optional, if excluded the default is true. This includes a random pause between 0 and 2 seconds between each API request, to prevent throttling.
3.	Step 3 – run PowerShell script (Expand-Groups.ps1) against CSV file from above. This extracts users out of Entra groups and Dataverse teams. This will then produce a new CSV file which is a list of all users extracted from groups and teams. It has smarts so it doesn’t expand the same group twice. It has the following inputs.
a.	InputCsvPath – name of source CSV file
b.	OutputCsvPath – name of output CSV file
c.	ExpandTeams – optional, and if excluded default is false. This will look at a Team type group and expand it, by using Dataverse and potentially Entra ID if needed too. This is for model driven apps.
d.	IncludePause – Optional, if excluded the default is true. This includes a random pause between 0 and 2 seconds between each API request, to prevent throttling.

We will use the output from all three of these as input files to a Power BI dashboard.

As you can imagine these scripts can take very long to execute across 20k apps. Several hours. A few suggestions:

1.	When running step 2, test with a few small CSV files first to check its working. You can also test with specific environment id’s. You could even just take the CSV file, and split into 4 or 5 CSVs and run the whole process 4 or 5 times. We can combine CSV’s later through simple script, or copy and paste.
2.	When running step 3, you could also trim out to have a small CSV to do some testing with first.
3.	Keep the outputs of these steps secure, they have some level of PII in them.
4.	When you run the scripts, they will need to run as you, and will need permissions to access resources e.g. Entra ID in Step 3, hence it’s a good idea to test first with small set of apps, to see if you have permissions. When you run each one, if you dont have correct PS module installed it will tell you, and if you haven't authenticated against the appropriate service it will tell you. 
