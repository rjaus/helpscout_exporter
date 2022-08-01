"""
Download conversations from a Help Scout mailbox and save to CSV file.
Be sure to replace the App ID and Secret in the authorization call as well
as the Mailbox ID in the conversations API call.
Python: 3.9.0
"""

import csv
import datetime
import requests
import getopt,sys
import os
import base64
import json
from decouple import config
from pprint import pprint

CLIENT_ID = config('HELPSCOUT_CLIENT_ID')
CLIENT_SECRET = config('HELPSCOUT_CLIENT_SECRET')

def authenticate():
    # The token endpoint.
    auth_endpoint = 'https://api.helpscout.net/v2/oauth2/token'
    # Preparing our POST data.
    post_data = ({
        'grant_type': 'client_credentials',
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET
    })

    # Send the data.
    r = requests.post(auth_endpoint, data=post_data)
    
    # Save our token.
    token = r.json()['access_token']

    return token


def list_mailboxes(endpoint_headers):
    r = requests.get("https://api.helpscout.net/v2/mailboxes", headers=endpoint_headers)

    for mailbox in r.json()["_embedded"]["mailboxes"]:
        print(str(mailbox["id"])+" : " + mailbox["name"])

def export_mailboxes(mailboxes_id, endpoint_headers, start_page, end_page, include_attachments):
    print(f"Starting on page {start_page}")
    print(f"Ending on page {end_page}")

    for mailbox_id in mailboxes_id.split(','): 
        print(mailbox_id)

    for mailbox_id in mailboxes_id.split(','): 
        all_conversations = False
        page = start_page if start_page else 1
        if not os.path.exists(mailbox_id+'/'):
            os.makedirs(mailbox_id+'/')

        # Read if there is an existing export file
        if os.path.exists(f"{mailbox_id}.csv"):
            # If so, open it, read the first entry.
            with open(f"{mailbox_id}.csv", mode="r", newline='', encoding='utf-8') as orig_f:
                csv_reader = csv.reader(orig_f)
                next(csv_reader)
                first_row = next(csv_reader)
                last_conversation_id = int(first_row[0])

                # Then save the whole thing to a new file
                # on completion, add the existing entries to the new file.
                # we will rewrite over the original file when creating the export
                with open(f"{mailbox_id}-tmp.csv", mode="w", newline='', encoding='utf-8') as tmp_f:
                    orig_f.seek(0)
                    tmp_f.write(orig_f.read())

        else:
            last_conversation_id = 0

        print(f"Last conversation ID: {last_conversation_id}")

        # Creates our file, or opens existing to prepend entries
        with open(str(mailbox_id)+'.csv', mode="w", newline='', encoding='utf-8') as fh:

            # Define our columns.
            columns = ['ID', 'Customer Name', 'Customer email addresses', 'Assignee', 'Status', 'Subject', 'Preview', 'Tags', 'Custom Fields', 'Created At', 'Closed At', 'Closed By', 'Resolution Time (seconds)']  
            csv_writer = csv.DictWriter(fh, fieldnames=columns) # Create our writer object.
            csv_writer.writeheader() # Write our header row.
            
            while not all_conversations:
                # Prepare conversations endpoint with status of conversations we want and the mailbox.
                conversations_endpoint = 'https://api.helpscout.net/v2/conversations?status=all&mailbox={}&page={}&embed=threads'.format(
                    mailbox_id,
                    page
                )

                r = requests.get(conversations_endpoint, headers=endpoint_headers)

                conversations = r.json()

                print(f"Total Pages {conversations['page']['totalPages']}")

                # Cycle over conversations in response.
                for conversation in conversations['_embedded']['conversations']:
                    print(conversation['id'])

                    # If we've reached the conversation from the start of our existing export file, stop exporting
                    if last_conversation_id == conversation['id']:
                        print(f"Reached already exported conversations.  Finishing... {last_conversation_id}")
                        all_conversations = True

                        # Save previous entries (from tmp file) to end of file
                        with open(f"{mailbox_id}-tmp.csv", mode="r", newline='', encoding='utf-8') as tmp_f:
                            csv_reader = csv.reader(tmp_f)
                            # Skip header
                            next(csv_reader)
                            for row in csv_reader:
                                csv_writer.writerow(dict(zip(columns, row)))

                        break

                    # If the email is missing, we won't keep this conversation.
                    # Depending on what you will be using this data for,
                    # You might omit this.
                    if 'email' not in conversation['primaryCustomer']:
                        print('Missing email for {}'.format(customer_name))
                        continue

                    ## CSV File Column & Row Assignment
                    # Prepare customer name.
                    customer_name = '{} {}'.format(
                        conversation['primaryCustomer']['first'],
                        conversation['primaryCustomer']['last']
                    )

                    # Prepare assignee, subject, and closed date if they exist.
                    assignee = '{} {}'.format(conversation['assignee']['first'], conversation['assignee']['last']) \
                        if 'assignee' in conversation else ''
                    subject = conversation['subject'] if 'subject' in conversation else 'No subject'
                    preview = conversation['preview'] if 'preview' in conversation else ''
                    closed_at = conversation['closedAt'] if 'closedAt' in conversation else ''
                    tags = json.dumps(conversation['tags']) if 'tags' in conversation else ''
                    custom_fields = json.dumps(conversation['customFields']) if 'customFields' in conversation else ''

                    # If the conversation has been closed, let's get the resolution time and who closed it.
                    closed_by = ''
                    resolution_time = 0
                    if 'closedByUser' in conversation and conversation['closedByUser']['id'] != 0:
                        closed_by = '{} {}'.format(
                            conversation['closedByUser']['first'], conversation['closedByUser']['last']
                        )
                        createdDateTime = datetime.datetime.strptime(conversation['createdAt'], "%Y-%m-%dT%H:%M:%S%z")
                        closedDateTime = datetime.datetime.strptime(conversation['closedAt'], "%Y-%m-%dT%H:%M:%S%z")
                        resolution_time = (closedDateTime - createdDateTime).total_seconds()

                    csv_writer.writerow({
                        'ID': conversation['id'],
                        'Customer Name': customer_name,
                        'Customer email addresses': conversation['primaryCustomer']['email'],
                        'Assignee': assignee,
                        'Status': conversation['status'],
                        'Subject': subject,
                        'Preview': preview,
                        'Tags': tags,
                        'Custom Fields': custom_fields,
                        'Created At': conversation['createdAt'],
                        'Closed At': closed_at,
                        'Closed By': closed_by,
                        'Resolution Time (seconds)': resolution_time
                    })

                    ## Save threads as HTML files in directory
                    threads = conversation['_embedded']['threads']
                    body = '<html><body>'
                    for thread in threads:
                        ## Function for handling attachments
                        ## Suggest using the new attachment export function instead of this
                        if include_attachments==True:  
                            # Handle attachments
                            if 'attachments' in thread["_embedded"].keys():
                                for attachment in thread["_embedded"]["attachments"]:
                                    # Attachment data
                                    r = requests.get(attachment["_links"]['data']['href'], headers=endpoint_headers)
                                    data = r.json()['data']
                                    with open(f"{mailbox_id}/{conversation['id']}-{attachment['filename']}", "wb") as f:
                                        f.write(base64.b64decode(data))

                        createdBy = thread['createdBy']['email']
                        if 'body' in thread.keys():
                            body += "\n<br><b>From:"+createdBy+"</b><br/>"+thread['body']+'<br /><tr>'
                    body += "</body></html>"
                    with open(mailbox_id+'/'+str(conversation['id'])+'.html', 'w+') as f:
                        f.write(body)

                if page == conversations['page']['totalPages']:
                    all_conversations = True
                    continue
                else:
                    page += 1
                    print(f"Page {page} of {conversations['page']['totalPages']}")
                    if end_page:
                        if page >= end_page:
                            exit()




def export_attachments_only(mailboxes_id, endpoint_headers, start_page, end_page):
    print(f"Starting on page {start_page}")
    print(f"Ending on page {end_page}")

    for mailbox_id in mailboxes_id.split(','): 
        print(mailbox_id)

    for mailbox_id in mailboxes_id.split(','): 
        all_conversations = False
        page = start_page if start_page else 1
        if not os.path.exists('attachments/'):
            os.makedirs('attachments/')
            
        while not all_conversations:
            # Prepare conversations endpoint with status of conversations we want and the mailbox.
            conversations_endpoint = 'https://api.helpscout.net/v2/conversations?status=all&mailbox={}&page={}&embed=threads'.format(
                mailbox_id,
                page
            )

            r = requests.get(conversations_endpoint, headers=endpoint_headers)

            conversations = r.json()

            print(f"Total Pages {conversations['page']['totalPages']}")

            # Cycle over conversations in response.
            for conversation in conversations['_embedded']['conversations']:
                print(conversation['id'])
                # If the email is missing, we won't keep this conversation.
                # Depending on what you will be using this data for,
                # You might omit this.
                if 'email' not in conversation['primaryCustomer']:
                    print('Missing email for {}'.format(conversation['id']))
                    continue

                threads = conversation['_embedded']['threads']

                for thread in threads:
                    if thread['_embedded']['attachments']:
                        print(f"Attachment(s) found for {conversation['primaryCustomer']['email']}")
                        # Download the attachment(s)
                        for attachment in thread["_embedded"]["attachments"]:
                            # If it doesn't exist, download it
                            if not os.path.exists(f"attachments/{conversation['primaryCustomer']['email']}"):
                                os.makedirs(f"attachments/{conversation['primaryCustomer']['email']}")

                            if not os.path.exists(f"attachments/{conversation['primaryCustomer']['email']}/{attachment['filename']}"):
                                # Attachment data
                                print(f"Download... size: {attachment['size']}")
                                r = requests.get(attachment["_links"]['data']['href'], headers=endpoint_headers)
                                data = r.json()['data']
                                with open(f"attachments/{conversation['primaryCustomer']['email']}/{attachment['filename']}", "wb") as f:
                                    f.write(base64.b64decode(data))
                            else:
                                print("attachment exists")
                        
            if page == conversations['page']['totalPages']:
                all_conversations = True
                continue
            else:
                page += 1
                print(f"Page {page} of {conversations['page']['totalPages']}")
                if end_page:
                    if page >= end_page:
                        exit()       

                        



def main(argv):
    opt_list_mailboxes = False
    opt_export_mailboxes = None
    opt_start_page = False
    opt_end_page = False
    include_attachments = False
    attachments_only = False
    try:
        opts, args = getopt.getopt(argv,"le:", ['start=','end=','attachments', 'attachments-only'])
    except getopt.GetoptError:
        print('-l : list all messageries')
        print('-e id : export id')
        sys.exit(2)

    for (opt,arg) in opts:
        if (opt == '-l'):
            opt_list_mailboxes = True
        if (opt == '-e'):
            opt_export_mailboxes = arg
        if (opt == '--start'):
            opt_start_page = int(arg)
        if (opt == '--end'):
            opt_end_page = int(arg)
        if (opt == '--attachments'):
            include_attachments = True
        if (opt == '--attachments-only'):
            attachments_only = True

    token = authenticate()


    # Prepare our headers for all endpoints using token.
    endpoint_headers = {
        'Authorization': 'Bearer {}'.format(token)
    }


    if opt_list_mailboxes:
        list_mailboxes(endpoint_headers)


    if opt_export_mailboxes != None:
        if attachments_only == True:
            export_attachments_only(opt_export_mailboxes,endpoint_headers, opt_start_page, opt_end_page)
        else:
            export_mailboxes(opt_export_mailboxes,endpoint_headers, opt_start_page, opt_end_page, include_attachments) 
        

if __name__ == "__main__":
   main(sys.argv[1:])