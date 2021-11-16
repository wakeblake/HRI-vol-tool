# HRI-vol-tool
Volunteer Hours Reporting Tool built on Google Apps Script

The Volunteer Hours Reporting Tool was designed to make volunteer reporting requirements at a small nonprofit with a large volunteer workforce much less onerous. Built on the free Google Apps Script platform, the tool permits volunteers to submit requested, properly formatted data through an authenticated client-side web app, which is then saved and aggregated on an internal Google tracking spreadsheet.  When the volunteer is authenticated, the web app simplifies the reporting process by loading a template of her data from a central repository "protected" sheet on the tracking spreadsheet.  


The Tracking Google Spreadsheet

The program runs on the designated tracking spreadsheet as a container-bound script (i.e. unique to a specific spreadsheet instance). When the tracking spreadsheet is opened, the Extensions menu populates with the tool as a custom add-on "Volunteer Reporting Tool":


Importing Volunteer Data

Volunteer data is imported as a .csv file using the "Import a File" menu item.  Volunteer data must be imported in this way, and should not be created manually or else formatting issues may create client-side errors.


Activating Sheets

The resulting uploaded sheet must be activated before volunteers can make client-side data requests and report their hours. The "Activate Sheets" menu item creates the central repository "protected" sheet from which the web app pulls volunteer data, as well as a linked "report summary" sheet that aggregates volunteers' submitted responses. Active sheets have the following characteristics:
<ul>
  <li>Have edit permissions restricted to the administrator</li>
  <li>Have data validation rules set to prevent editing <i>restricted</i> rows and columns (Primary Keys, Clients, and the header row)</li>
  <li>Have an invalid email filter trigger that notifies the administrator when an email is improperly formatted (columns with email data only)</li>
  <li>Have blue-colored tabs</li>
  <li>Only two active sheets exist at any time (a "protected" sheet and a "report summary" sheet)</li>
</ul>


Deactivating Sheets

Running the tool's "Activate Sheets" process while active sheets exist will permit the administrator to deactivate those sheets and set a new protected and report summary sheet. This might be done, for instance, to make edits to restricted cells or to activate a new uploaded set of data. 
