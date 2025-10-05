try parse(dob):
    pass
except:
    scrape_info_label.set("\u2740 Error in DOB\nTry again.")
        root.update_idletasks()
        return

if not mrn.isdigit():
        scrape_info_label.set("\u2740 Error in MRN\nTry again.")
        root.update_idletasks()
        return
 
elif (not is_email(email)) or (email in {"", "na"}):
        scrape_info_label.set("\u2740 Problem with email")
        root.update_idletasks()
        return
else:
    scrape_info_label.set("\u2705")
    root.update_idletasks()
    return