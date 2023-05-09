# GradesAutomater

This script's purpose is to help my high school teachers automate sending out midterm and final marks (rough versions) rather then having to manually email each student their mark. This script has been made with Google Apps Script.

## How to use:

-> You can view the spreadsheet and the code via this link as well:
  - https://docs.google.com/spreadsheets/d/14ZLlK3m54Wstukh2_04En0PDVpMU-lzycE-Wx6LKbGU/edit?usp=sharing

1) First of all what you will want to do is open up a new google sheets template and head over to the **Extensions** button which resides on the Menu. Then you will click on the **Apps Script*** button.


<img width="850" alt="image" src="https://github.com/Rush-Shaw/GradesAutomater/assets/69100536/a2e458a6-babd-4d2c-a09b-f5bdc1de81fc">


2) Next you'll want to copy the code from the file titled _Code.gs_ and paste it into your now opened google apps script workflow.


3) You can now save that code by _CTRL + S_ or _CMD + S_ and then refresh your google sheet. You will notice now that a new button labelled **Deploy Grades** has now appeared in the menu. 

<img width="738" alt="image" src="https://github.com/Rush-Shaw/GradesAutomater/assets/69100536/0e8b5e99-3032-48e3-be73-61765a7a9c5f">


4) Before sending out your grades, you must make an email draft (template) so that the script can find the said draft and use that as a base. Wherever you want that column to be shown, you will have to write the name of that column heading in the email surrounded with {{}}. 

Shown below is an example template I've made that user's can use:

<img width="595" alt="image" src="https://github.com/Rush-Shaw/GradesAutomater/assets/69100536/05ac4b7e-24fb-4285-8714-517dd9a4c75b">


5) Now, you can finally press the newly added dropdown and click on **Send Emails**. You will then be prompted to enter the heading of your email draft. This is very neccessary as the script will find the draft using the heading name so it must be existant.

<img width="749" alt="image" src="https://github.com/Rush-Shaw/GradesAutomater/assets/69100536/0e0f5b1e-7422-44df-9dc3-c7ed6a41f45d">


6) Finally, you can click on **Ok** and watch the magic unfold as your sent mails will show you the date they were sent on and that cell will turn green.

<img width="794" alt="image" src="https://github.com/Rush-Shaw/GradesAutomater/assets/69100536/e91dc67c-40da-498e-bc83-edbac5ae8dea">


*Note, if a email is non-existant or does not seem to work, then you will be emailed about it so don't worry!
