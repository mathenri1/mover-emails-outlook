import win32com.client

# Sincronizar conta do Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI").Folders['Conta']

# Vincular pastas que vão ser salvas
inbox = outlook.Folders('Inbox')
pasta1_folder = inbox.Folders('Pasta1')
past2_folder = pasta1_folder.Folders['Archive Monitoring']
pasta3_subfolder = pasta1_folder.Folders['Job Monitoring']
pasta4_folder = pasta1_folder.Folders['TIPS Monitoring']

# Mover e-mails
for message in inbox.Items:
    if 'Título1' in message.Subject:
        message.Move(pasta2_folder)
    elif 'Título2' in message.Subject:
        message.Move(pasta3_subfolder)
    elif 'Titulo3' or 'Titulo3.1' in message.Subject:
        message.Move(pasta4_folder)
