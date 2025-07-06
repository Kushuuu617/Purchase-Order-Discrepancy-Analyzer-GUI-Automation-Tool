def update_status(status_text, msg):
    status_text.insert("end", msg + "\n")
    status_text.see("end")