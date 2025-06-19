Description:-
    This is a addin for microsoft outlook. The name of the addin is "Mail manager"
    The addin provides the following features:-
                  >Whenever you send an email, a json file containing information about the mail is stored in your local device.
                  >Whenever you send an email, a copy of all the attachments are stored locally as well.

Extra features:-
    Some extra features are also provided in the form of mail modes and rule. There are 3 
    rules that your email must follow in order for the addin to approve it and send.
    The three rules are:-
                  >Attachments names cannot be suspicious (malware.txt,virus.js etc).
                  >Attachments cannot exceed 5mb.
                  >Attachments cannot contain confidential information (checks for keywords like "secret" or "confidential" in the txt files).

Mail modes:-
    These conditions might be too constricting for some users who might want to bypass them, 
    for example some users would not want their attachment content to be analysed or they might 
    want to send attachments larger than 5mb. So for these reasons users are provided with 3 
    mail modes (Private,Protected,Public) to bypass them.
    ![image](https://github.com/user-attachments/assets/433c26f3-56be-4dd1-89d5-94d5c86d9d77)

    Private mode: In private mode, all 3 conditions are enforced normally.
    Protected mode: In protected mode, the attachment content is not checked, only the size and name are checked.
    Public mode: In public mode, only the attachment name is checked. The size and content are not checked


How to run:-
    All the necessary js,html files are deployed on render so you only need the manifest file to add the addin to your account and the python exe file so that the files can be saved to the local storagE.
    However if your account already has the addin then you need only the python exe file. 
    In order to add the addin to your account for the first time, 
    1.download the manifest.xml file, 
    2.then go to microsft outlook.
    3.click on the application icon and then click on get addins
    ![image](https://github.com/user-attachments/assets/83a309a4-df68-4b7c-b54b-b9e7309c6d47)
    ![image](https://github.com/user-attachments/assets/e5103cb6-ee85-48a2-a560-2d5a55a5b185)

    4.click on my addins and then add the addin by clicking on the manifest.xml file
    ![image](https://github.com/user-attachments/assets/9c3318e0-eead-4b23-b724-2d676d3941cf)
    ![image](https://github.com/user-attachments/assets/9a039128-b00b-41eb-afb8-f248586c8159)

    5. the addin has now been successfully added.
    

    
