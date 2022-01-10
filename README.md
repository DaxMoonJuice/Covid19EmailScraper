# Covid19EmailScraper
A python script to process NHS Covid Test Emails and output results to a CSV

Built in a day to help out a friend

<h1>Instructions For Use </h1>

<ul>
<li>Download and extract the files from the github repository</li>
<li>Double click the "CovidEmailProcessor.exe" file to run the script</li>
</ul>


<h1> Configuring Script To Work With Your Outlook Folder Structure </h1>

The site_list_prod.json file is used to tell the script which outlook accounts + folders it needs to search through for Covid Test Emails

JSON is a type of text format for transferring and storing data. The site_list_prod.json file contains a list of configurations in the a JSON format

Each of the sets of { } brackets in the list (the [ ] brackets) contains the configuration information for a single site.

<h2>Example Configuration</h2>

<i>[  {  
      "Site_Name": "Reading",  
    "Email_Account": "test@testington.com",  
    "Folder_Path": "Inbox/Factory Workers/Reading Central"  
  }  ]</i>



The <b>Site_Name</b> is the name the Site will appear as in the output results

<b>Email_Account</b> is the logged in email_account containing the folder + emails

<b>Folder_Path</b> is the path to the location of the folder. The folder path "Inbox/Factory Workers/Reading Central" is saying to the script. The emails to process for this site are in the "Reading Central" folder inside the "Factory Workers" folder inside the "Inbox" folder.
