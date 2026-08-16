from win11toast import toast

def show_notification(title, message):
    try:
        toast(title, message)
    except Exception as e:
        print(f"Failed to show notification: {e}")
