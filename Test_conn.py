import subprocess
import os
import win32com.client as win32

os.chdir("C:\\New-Project\\Power-Shell")

commands = [
    "tnc google.com -DiagnoseRouting",
    "Test-NetConnection google.com -DiagnoseRouting",
    'Test-NetConnection -ComputerName pac.zscaler.net -Port 80',
    "ipconfig /all",
    'ping -n 1 "Default Gateway"',
    #'$ip_zscaler_gateway = (nslookup gateway.zscalar.net | Select-String "Address: " | ForEach-Object { $_.ToString().Trim("Address: ") })[0]',
    'Test-NetConnection -ComputerName gateway.zscaler.net -Port 9400',
    "nslookup google.com",
    "route print"
]

output_text = ""
for idx, command in enumerate(commands):
    if command.startswith('ping -n 1 "Default Gateway"'):
        process = subprocess.Popen(['ping', '-n', '1', 'Default Gateway'], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        output, errors = process.communicate()
    elif command.startswith('$ip_zscaler_gateway'):
        ip_zscaler_gateway = subprocess.check_output(['powershell', '-Command', command], text=True).strip()
        output_text += f"Session {idx + 1}"
        output_text += f"Command: {command}"
        output_text += f"IP Address of gateway.zscalar.net: {ip_zscaler_gateway}\n"
    else:
        process = subprocess.Popen(["powershell", "-Command", command], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        output, errors = process.communicate()

        if process.returncode == 0:
            output_text += f"Session {idx + 1}\n"
            output_text += f"Command: {command}\n"
            output_text += f"Output:\n{output}\n"
            separator = "=" * 30
        else:
            output_text += f"Session {idx + 1}\n"
            output_text += f"Error occurred while executing command: {command}\n"
            output_text += f"Errors:\n{errors}\n"
        

print("All commands executed.")
separator = "=" * 30
message = "Hi team,\nPlease find the below test result from my PC for further investigation.\n\n" + output_text + "\nThanks"

# Send the email using Outlook
outlook = win32.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)  # 0 represents a new mail item
mail.To = "Anil.Mchandran@ust.com"
mail.Subject = "Initial Connectivity Test Data"
mail.Body = message
mail.Send()
print("Email sent successfully.")
