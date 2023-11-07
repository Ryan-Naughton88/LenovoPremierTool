import customtkinter as ctk
import win32com.client
import datetime

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")
ctk.set_widget_scaling(0.8)  # widget dimensions and text size
ctk.set_window_scaling(1.0)  # window geometry dimensions


def copy_WO_template():
    template_text = """V.WO.Onsite.ProbDesc.003
Labor Only--
Problem Verification Description--
Tech Direction--
Part(s) for onsite service--
Previous Failed Repair-- 
Parts Sent Directly to Customer-- 
Customer Educated on BitLocker-- 
Repeat Repair for Same Issue--
RAID Configured--
Additional Comments--"""

    root.clipboard_clear()  # Clear the clipboard
    root.clipboard_append(template_text)  # Copy the template text to the clipboard
    root.update()  # Required to ensure the clipboard is updated

def copy_depot_template():
    # Define your template text here
    template_text = """V.WO.Depot.ProbDesc.001
Problem Description-- 
Suggested Tech Direction--
Customer Educated on BitLocker-- 
Informed customer to take notice of the Depot Drop in the Box document- Y
"""

    root.clipboard_clear()  # Clear the clipboard
    root.clipboard_append(template_text)  # Copy the template text to the clipboard
    root.update()  # Required to ensure the clipboard is updated

def copy_notes_template():
    #Define the current date in a (M,D,Y) format
    current_date = datetime.datetime.now()
    formatted_date = current_date.strftime("%m-%d-%Y %I:%M %p")
    #Define input variables
    name = name_entry.get()
    case = case_entry.get()
    model = model_entry.get()
    serial = serial_entry.get()
    mtm = mtm_entry.get()
    email = email_entry.get()
    phone = phone_entry.get()
    address = address_entry.get()
    # Define your template text here
    template_text = f"""V.Case.Timeline.Note.002
rnaughton {formatted_date}
 
Caller Name: {name}
 
Problem Description:
- 

 Action/Troubleshooting Done
- 
- 
- 
- 
- 
- 

 Resolution Plan/Next Steps:
- 
 
Reminder: Approval required in Internal Note for repeated SB-SSD-LCD-RAM

PARTS: 

EXTRAS:
-
-
-

CASE#: {case}

MODEL: {model}
SERIAL: {serial}
MTM: {mtm}

CUST EMAIL: {email}
CUST PHONE: {phone}
CUST ADDRESS: {address}
---------------------------------------------------------------------------------------------------------------------------------------
"""

    root.clipboard_clear()  # Clear the clipboard
    root.clipboard_append(template_text)  # Copy the template text to the clipboard
    root.update()  # Required to ensure the clipboard is updated

def clear_all_fields():
    for entry in [name_entry, email_entry, phone_entry, address_entry, model_entry, serial_entry, mtm_entry, case_entry, wonum_entry, cuwonum_entry]:
        entry.delete(0, 'end')

def create_bitlocker_email():
    subject = f'Lenovo Case# {case_entry.get()}'
    recipient_email = email_entry.get()
    name = name_entry.get()
    model = model_entry.get()
    wo = wonum_entry.get()
    mtm = mtm_entry.get()
    serial = serial_entry.get()
    address = address_entry.get()
    phone = phone_entry.get()
    service_type = dropdown.get()

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = recipient_email
    mail.Subject = subject
    html_source = f"""
    <!DOCTYPE html>
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:dt="uuid:C2F41010-65B3-11d1-A29F-00AA00C14882" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=us-ascii">
    <meta name="Generator" content="Microsoft Word 15 (filtered medium)">
</head>
<body lang="EN-US" link="#0563C1" vlink="#954F72" style="word-wrap: break-word">
    <div class="WordSection1">
        <p style="mso-margin-top-alt: 0in; margin-right: 0in; margin-bottom: 8.0pt; margin-left: 0in">Hello {name},<o:p></o:p></p>
        <p style="margin: 0in">&nbsp;<o:p></o:p></p>
        <p style="margin: 0in">Thank you for contacting Lenovo Premier Technical Support. In order to resolve the issue on your <span style="color: red">{model}</span>, I have created <span style="color: red">{service_type} </span>Work Order# <span style="color: red">{wo}</span><o:p></o:p></p>
        <p style="margin: 0in"><span style="color: red">&nbsp;<o:p></o:p></span></p>
        <p style="margin: 0in"><b>Machine Type: {mtm}</b><o:p></o:p></p>
        <p style="margin: 0in"><b>Serial #: {serial}</b><o:p></o:p></p>
        <p style="margin: 0in"><b>Address: {address}</b><o:p></o:p></p>
        <p style="margin: 0in"><b>Phone# {phone}</b><o:p></o:p></p>
        <p style="margin: 0in">&nbsp;<o:p></o:p></p>
        <p style="margin: 0in">&nbsp;<o:p></o:p></p>
        <p style="margin: 0in">Next Business Day service is subject to availability of service parts. Parts in stock will ship next business day. Once the Lenovo onsite technician receives the part(s), he/she will call to coordinate the onsite repair time.<o:p></o:p></p>
        <p style="margin: 0in"> <o:p></o:p></p>
        <p style="margin: 0in">Throughout the service delivery process, we will monitor your case and contact you periodically to provide updates. If you have any questions, please feel free to contact me directly. You can also call the Premier Technical Support team directly at 1-855-669-3600.<o:p></o:p></p>
        <p style="margin: 0in">&nbsp;<o:p></o:p></p>
        <p style="margin: 0in"><b><span style="font-size: 14.0pt">Microsoft BitLocker</span></b><span style="font-size: 14.0pt"><o:p></o:p></span></p>
        <p style="margin: 0in"><span style="color: black">We have requested a service for a systemboard replacement. If your hard drive is protected by BitLocker encryption, switching the systemboard will lock your drive requiring you to enter a recovery key. If you do not know your recovery key, you should disable BitLocker before swapping out the systemboard.<o:p></o:p></span></p>
        <p style="margin: 0in"><span style="color: black">&nbsp;<o:p></o:p></span></p>
        <p style="margin: 0in"><b><span style="color: red">Important: If you have a company IT Department, please verify with them the BitLocker policy before following these directions.&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></b><span style="color: red"><o:p></o:p></span></p>
        <p style="margin: 0in"><span style="color: black">&nbsp;<o:p></o:p></span></p>
        <p style="margin: 0in"><span style="color: black">To check for the Recovery Key, follow the directions here:&nbsp; <a href="https://support.lenovo.com/us/en/solutions/HT506878">https://support.lenovo.com/us/en/solutions/HT506878</a><o:p></o:p></p>
        <p style="margin: 0in"><span style="color: black">&nbsp;<o:p></o:p></span></p>
        <p style="margin: 0in"><span style="color: black">If you do not have a Recovery Key, follow these directions to see if it needs to be disabled:<o:p></o:p></span></p>
        <p style="margin: 0in"><span style="color: black">&nbsp;<o:p></o:p></span></p>
        <ol style="margin-top: 0in" start="1" type="1">
            <li class="MsoNormal" style="color: black; mso-list: l0 level1 lfo1; vertical-align: middle">Click on the Start menu<o:p></o:p></li>
            <li class="MsoNormal" style="color: black; mso-list: l0 level1 lfo1; vertical-align: middle">Type ‘<b>manage bitlocker</b>’ and open the ‘Manage Bitlocker’ Control Panel<o:p></o:p></li>
            <li class="MsoNormal" style="color: black; mso-list: l0 level1 lfo1; vertical-align: middle">The Control Panel will show the status of BitLocker on your drive.<o:p></o:p></li>
            <li class="MsoNormal" style="color: black; mso-list: l0 level1 lfo1; vertical-align: middle">If BitLocker is enabled, click “<b>Turn off BitLocker</b>”<o:p></o:p></li>
        </ol>
        <ol style="margin-top: 0in" start="4" type="1">
            <ol style="margin-top: 0in" start="1" type="a">
                <li class="MsoNormal" style="color: black; mso-list: l0 level2 lfo2; vertical-align: middle">Decryption of your drive can take several hours, so it needs to be done the day before the tech is scheduled.<o:p></o:p></li>
            </ol>
        </ol>
        <p style="mso-margin-top-alt: 0in; margin-right: 0in; margin-bottom: 0in; margin-left: 81.0pt"><span style="color: black">&nbsp;<o:p></o:p></span></p>
        <p style="margin: 0in"><b><span style="font-size: 14.0pt">Windows System Login</span></b><span style="font-size: 14.0pt"><o:p></o:p></span></p>
        <p style="margin: 0in">If you currently use a PIN, Fingerprint, or Face Recognition, known as “Windows Hello”, to log into your computer, the systemboard replacement will clear this data and not allow login via these methods. <b><span style="color: red">You will need to know your <u>password</u> to access your system after the systemboard replacement.</span></b> If you do not know your password, please follow the directions from Microsoft for your version of Windows: <a href="https://support.microsoft.com/en-us/windows/change-or-reset-your-windows-password-8271d17c-9f9e-443f-835a-8318c8f68b9c">https://support.microsoft.com/en-us/windows/change-or-reset-your-windows-password-8271d17c-9f9e-443f-835a-8318c8f68b9c</a><o:p></o:p></p>
        <p style="margin: 0in">&nbsp;<o:p></o:p></p>
        <p style="margin: 0in">&nbsp;<o:p></o:p></p>
        <p style="margin: 0in">If you were satisfied with your Premier Support service, please consider letting my manager Tiffany Harrell know at <a href="mailto:tharrell4@lenovo.com">tharrell4@lenovo.com</a>.<o:p></o:p></p>
        <p style="margin: 0in">&nbsp;<o:p></o:p></p>
        <p style="margin: 0in"><b><i><span style="font-size: 14.0pt; color: red">Note: Please use "Reply All" to respond to this message. </span></i></b><span style="font-size: 14.0pt; color: red"><o:p></o:p></span></p>
        <p style="margin: 0in">&nbsp;<o:p></o:p></p>
        <p style="margin: 0in">Thank you for choosing Lenovo Premier Technical Support,<o:p></o:p></p>
        <p class="MsoNormal"><o:p>&nbsp;</o:p></p>
    </div>
</body>
</html>
"""
    mail.HTMLBody = html_source
    mail.Display()

def create_WO_email():
    subject = f'Lenovo Case# {case_entry.get()}'
    recipient_email = email_entry.get()
    name = name_entry.get()
    model = model_entry.get()
    wo = wonum_entry.get()
    mtm = mtm_entry.get()
    serial = serial_entry.get()
    address = address_entry.get()
    phone = phone_entry.get()
    service_type = dropdown.get()

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = recipient_email
    mail.Subject = subject
    html_source = f"""
    <!DOCTYPE html>
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:dt="uuid:C2F41010-65B3-11d1-A29F-00AA00C14882" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=us-ascii">
    <meta name="Generator" content="Microsoft Word 15 (filtered medium)">
</head>
<body lang="EN-US" link="#0563C1" vlink="#954F72" style="word-wrap: break-word">
    <div class="WordSection1">
        <p style="mso-margin-top-alt: 0in; margin-right: 0in; margin-bottom: 8.0pt; margin-left: 0in">Hello {name},<o:p></o:p></p>
        <p style="margin: 0in">&nbsp;<o:p></o:p></p>
        <p style="margin: 0in">Thank you for contacting Lenovo Premier Technical Support. In order to resolve the issue on your <span style="color: red">{model}</span>, I have created <span style="color: red">{service_type} </span>Work Order# <span style="color: red">{wo}</span><o:p></o:p></p>
        <p style="margin: 0in"><span style="color: red">&nbsp;<o:p></o:p></span></p>
        <p style="margin: 0in"><b>Machine Type: {mtm}</b><o:p></o:p></p>
        <p style="margin: 0in"><b>Serial #: {serial}</b><o:p></o:p></p>
        <p style="margin: 0in"><b>Address: {address}</b><o:p></o:p></p>
        <p style="margin: 0in"><b>Phone# {phone}</b><o:p></o:p></p>
        <p style="margin: 0in">&nbsp;<o:p></o:p></p>
        <p style="margin: 0in">&nbsp;<o:p></o:p></p>
        <p style="margin: 0in">Next Business Day service is subject to availability of service parts. Parts in stock will ship next business day. Once the Lenovo onsite technician receives the part(s), he/she will call to coordinate the onsite repair time.<o:p></o:p></p>
        <p style="margin: 0in"> <o:p></o:p></p>
        <p style="margin: 0in">Throughout the service delivery process, we will monitor your case and contact you periodically to provide updates. If you have any questions, please feel free to contact me directly. You can also call the Premier Technical Support team directly at 1-855-669-3600.<o:p></o:p></p>
        <p style="margin: 0in">&nbsp;<o:p></o:p></p>
        <p style="margin: 0in">If you were satisfied with your Premier Support service, please consider letting my manager Tiffany Harrell know at <a href="mailto:tharrell4@lenovo.com">tharrell4@lenovo.com</a>.<o:p></o:p></p>
        <p style="margin: 0in">&nbsp;<o:p></o:p></p>
        <p style="margin: 0in"><b><i><span style="font-size: 14.0pt; color: red">Note: Please use "Reply All" to respond to this message. </span></i></b><span style="font-size: 14.0pt; color: red"><o:p></o:p></span></p>
        <p style="margin: 0in">&nbsp;<o:p></o:p></p>
        <p style="margin: 0in">Thank you for choosing Lenovo Premier Technical Support,<o:p></o:p></p>
        <p class="MsoNormal"><o:p>&nbsp;</o:p></p>
    </div>
</body>
</html>"""
    mail.HTMLBody = html_source
    mail.Display()

root = ctk.CTk()
root.geometry("500x500")
root.title("Lenovo Premier Tool")

# Create top, bottom and button frames
top_frame=ctk.CTkFrame(root,  width=200,  height=  400)
top_frame.grid(row=0,  column=0,  padx=10,  pady=5, sticky="EWNS")

bottom_frame=ctk.CTkFrame(root, width=400, height=50)
bottom_frame.grid(row=1, column=0, sticky="EWNS")

button_frame=ctk.CTkFrame(root)
button_frame.grid(row=2, column=0, sticky="EWNS")

root.columnconfigure(0, weight=1)
root.rowconfigure(1, weight=1)

top_frame.columnconfigure(0, weight=1)
top_frame.rowconfigure(1, weight=1)

bottom_frame.columnconfigure(0, weight=1)
bottom_frame.rowconfigure(1, weight=1)


#create label for Customer Info
label = ctk.CTkLabel(top_frame, text='Customer Info', font=("Arial", 18), justify="center")
label.grid(row=0, columnspan=2, padx=10, pady=10)

#label and entry field for Customer Name
cust_name = ctk.CTkLabel(top_frame, text="Cust Name")
cust_name.grid(row=2, column=0, padx=10, pady=10)
name_entry = ctk.CTkEntry(top_frame, placeholder_text="Name", width=300)
name_entry.grid(row=2, column=1)

#label and entry field for Customer Email
cust_email = ctk.CTkLabel(top_frame, text="Cust Email")
cust_email.grid(row=3, column=0, padx=10, pady=10)
email_entry = ctk.CTkEntry(top_frame, placeholder_text="Email", width=300)
email_entry.grid(row=3, column=1, padx=15)

#label and entry field for Customer Phone
cust_phone = ctk.CTkLabel(top_frame, text="Cust Phone")
cust_phone.grid(row=4, column=0, padx=10, pady=10)
phone_entry = ctk.CTkEntry(top_frame, placeholder_text="Phone", width=300)
phone_entry.grid(row=4, column=1)

#label and entry field for Customer Address
cust_address = ctk.CTkLabel(top_frame, text="Cust Address")
cust_address.grid(row=5, column=0, padx=10, pady=10)
address_entry = ctk.CTkEntry(top_frame, placeholder_text="Address", width=300)
address_entry.grid(row=5, column=1)

#Create Label for Machine Info
label = ctk.CTkLabel(top_frame, text='Machine Info', font=("Arial", 18), justify="center")
label.grid(row=6, columnspan=2, padx=10, pady=10)

#Create Label and entry for Model
model = ctk.CTkLabel(top_frame, text="Model")
model.grid(row=7, column=0, pady=10)
model_entry = ctk.CTkEntry(top_frame, placeholder_text="Model", width=300)
model_entry.grid(row=7, column=1)

#Create Label and entry for Serial Number
serial = ctk.CTkLabel(top_frame, text="Serial")
serial.grid(row=8, column=0, pady=10)
serial_entry = ctk.CTkEntry(top_frame, placeholder_text="Serial", width=300)
serial_entry.grid(row=8, column=1)

#Create Label and entry for MTM
mtm = ctk.CTkLabel(top_frame, text="MTM")
mtm.grid(row=9, column=0, pady=10)
mtm_entry = ctk.CTkEntry(top_frame, placeholder_text="MTM", width=300)
mtm_entry.grid(row=9, column=1, padx=15)

#Create Label for Case Info
label = ctk.CTkLabel(top_frame, text='Case Info', font=("Arial", 18), justify="center")
label.grid(row=10, columnspan=2, padx=10, pady=10)

#Create Label and Entry for Case Number
case = ctk.CTkLabel(top_frame, text="Case Number")
case.grid(row=11, column=0, pady=10, sticky="EW")
case_entry = ctk.CTkEntry(top_frame, placeholder_text="Case#")
case_entry.grid(row=11, column=1, padx=15, sticky="EW")

#Create Label and Entry for WO Number
wonum = ctk.CTkLabel(top_frame, text="WO Number")
wonum.grid(row=12, column=0, pady=10, sticky="EW")
wonum_entry = ctk.CTkEntry(top_frame, placeholder_text="WO#")
wonum_entry.grid(row=12, column=1, padx=15, sticky="EW")

#Create Label and Entry for Current WO Number
cuwonum = ctk.CTkLabel(top_frame, text="Current WO")
cuwonum.grid(row=13, column=0, pady=10, sticky="EW")
cuwonum_entry = ctk.CTkEntry(top_frame, placeholder_text="Current WO #")
cuwonum_entry.grid(row=13, column=1, padx=15, sticky="EW")

#Create Label for Service Type Dropdown and dropdown list with options
options = ['Onsite', 'Depot', 'Parts Only']
selected_option = ctk.StringVar()

service_type = ctk.CTkLabel(top_frame, text="Service Type")
service_type.grid(row=14, column=0, pady=10, sticky="EW")
#label = ctk.CTkLabel(top_frame, text="Selected Option:")
#label.grid(row=14, column=1, pady=10, sticky="EW")
dropdown = ctk.CTkComboBox(top_frame, values=options)
dropdown.grid(row=14, column=1, sticky="EW")

#Button to clear fields
reset = ctk.CTkButton(top_frame, text="Reset Fields", command=clear_all_fields)
reset.grid(row=15, columnspan=2, stick="EW")

#Create Label for Textbox
notes = ctk.CTkLabel(bottom_frame, text="Premier Notes Template", font=("Arial", 18), justify="center")
notes.grid(row=0, column=0, columnspan=2)

#Create Button to copy notes template
copy_notes_button = ctk.CTkButton(bottom_frame, text="Notes Template", command=copy_notes_template)
copy_notes_button.grid(row=1, column=0, columnspan=2, sticky="EW")

#Create Textbox
textbox = ctk.CTkTextbox(bottom_frame, width=600, height=550, pady=4, font=("Arial", 15))
textbox.grid(row=2, column=0, columnspan=2,sticky="NS")

#Create BitLocker Button
bitlocker_button = ctk.CTkButton(button_frame, text="BitLocker Email", command=create_bitlocker_email)
bitlocker_button.grid(row=0, column=0, sticky="EW")

#Create WO Email Button
email_button = ctk.CTkButton(button_frame, text="WO Email", command=create_WO_email)
email_button.grid(row=0, column=1, sticky="EW")

#Create WO Template Button
wotemplate_button = ctk.CTkButton(button_frame, text="WO Template", command=copy_WO_template)
wotemplate_button.grid(row=0, column=2, sticky="EW")

#Create Depot Template Button
depottemplate_button = ctk.CTkButton(button_frame, text="Depot Template", command=copy_depot_template)
depottemplate_button.grid(row=0, column=3, sticky="EW")

root.mainloop()