import os
from gmail import get_gmail_service, fetch_unread_emails
from llm import analyze_email
from excel import append_to_excel, message_id_exists
from notifier import show_notification

def process_emails():
    print("Starting email processing...")
    service = get_gmail_service()
    
    # Use "is:unread" to process everything new since the last run
    emails = fetch_unread_emails(service,5)
    
    if not emails:
        print("No new relevant emails found.")
        return

    placement_updates = []
    academic_count = 0
    other_count = 0
    
    for email in emails:
        message_id = email['id']
        subject = email['subject']
        sender = email['sender']
        body = email['body']
        
        # Prevent double-processing if it's already in the Excel sheet
        if message_id_exists(message_id):
            print(f"Skipping duplicate: {subject}")
            continue
            
        print(f"Analyzing with LLM: {subject}")
        result = analyze_email(subject, body)
        
        if not result:
            print(f"Failed to parse email: {subject}")
            continue
            
        category = result.get("category", "OTHER")
        subtype = result.get("subtype", "OTHER")
        extracted_data = result.get("extracted_data", {})
        digest = result.get("digest", subject)
        
        try:
            email_metadata = {
                'id': message_id,
                'date': email['date'],
                'subtype': subtype,
                'link': email['link'],
                'summary': digest
            }
            
            if category == "ACADEMIC":
                academic_count += 1
                append_to_excel({}, email_metadata)
                
            elif category == "OTHER" or category == "OTHER / REVIEW":
                other_count += 1
                append_to_excel({}, email_metadata)
                
            elif category == "PLACEMENT":
                append_to_excel(extracted_data, email_metadata)
                placement_updates.append(digest)
                
        except PermissionError:
            print("\n❌ ERROR: The Excel file is currently open!")
            print("Please close 'College_Placement_2026.xlsx' in Excel and run the script again.")
            return # Stop processing further to avoid data loss
            
        # Optional: Mark as read after successful processing
        # mark_as_read(service, message_id)
            
    # Send Notification
    if placement_updates or academic_count > 0:
        title = "College Mail — " + emails[-1]['date']
        lines = []
        
        if placement_updates:
            lines.append(f"🔴 {len(placement_updates)} Placement Updates")
            for u in placement_updates[:5]:
                lines.append(u)
            if len(placement_updates) > 5:
                lines.append(f"...and {len(placement_updates)-5} more.")
                
        if academic_count > 0:
            lines.append(f"📚 {academic_count} academic updates")
            
        if other_count > 0:
            lines.append(f"❓ {other_count} other/review required")
            
        message = "\n".join(lines)
        print("\n--- Notification Triggered ---")
        
        # Handle Windows console emoji encoding issues
        try:
            print(title)
            print(message)
        except UnicodeEncodeError:
            print("College Mail Update (Emojis hidden due to terminal encoding)")
            print(message.encode('ascii', 'ignore').decode('ascii'))
            
        show_notification(title, message)

if __name__ == "__main__":
    process_emails()
