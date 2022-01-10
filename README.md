# Covid19 Email Scraper
A python script to process NHS Covid Test Emails and output results to a CSV. Built in a weekend to help out a friend

The script works as follows

<ul>

<li>Connects to an instance of Microsoft Outlook running locally via the win32com library</li> 
<li>Looks through email folders as specified in the configuation file "site_list.json"</li>
<li>Identifies email templates (Positive + Negative Results For PCR + Lateral Flow Tests, Kit Registration Emails)</li>
<li>Extracts key fields per template</li>
<li>Exports results to csv file</li>
      
</ul>


<h1>Instructions For Use </h1>

<ul>
<li>Download and extract the files from the github repository</li>
<li>Double click the "CovidEmailProcessor.exe" file to run the script</li>
</ul>


<h1> Configuring Script To Work With Your Outlook Folder Structure </h1>

The site_list_prod.json file is used to tell the script which outlook accounts + folders it needs to search through for Covid Test Emails.

JSON is a type of text format for transferring and storing data. The site_list.json file contains a list of configurations in the JSON format.

Each of the sets of { } brackets in the list (the [ ] brackets) contains the configuration information for a single site. 

An example of a configuration file can be found below.

<h2>Example Configuration</h2>

<i>[  {  
      "Site_Name": "Reading",  
    "Email_Account": "test@testington.com",  
    "Folder_Path": "Inbox/Factory Workers/Reading Central"  
  }  ]</i>



The <b>Site_Name</b> is the name of the test site the email folder corresponds to.

<b>Email_Account</b> is the logged in email_account containing the folder + emails.

<b>Folder_Path</b> is the path to the location of the folder. The folder path "Inbox/Factory Workers/Reading Central" is saying to the script. The emails to process for this site are in the "Reading Central" folder inside the "Factory Workers" folder inside the "Inbox" folder.
