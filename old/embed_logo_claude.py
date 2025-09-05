import win32com.client as win32

def detailed_embed_example():
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    
    mail.To = "test@example.com"
    mail.Subject = "Understanding Image Embedding"
    
    logo_path = r"C:\path\to\logo.png"
    
    # STEP 1: Add image as regular attachment
    print("Step 1: Adding image as attachment...")
    attachment = mail.Attachments.Add(logo_path)
    print(f"Attachment added: {attachment.FileName}")
    
    # STEP 2: Convert attachment to embedded resource
    print("Step 2: Setting Content-ID property...")
    
    # The MAPI property for Content-ID
    CONTENT_ID_PROPERTY = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
    
    # Our unique identifier for this image
    our_content_id = "my_logo_123"
    
    # This is the magic line that converts regular attachment to embedded image
    attachment.PropertyAccessor.SetProperty(CONTENT_ID_PROPERTY, our_content_id)
    print(f"Content-ID set to: {our_content_id}")
    
    # STEP 3: Reference it in HTML
    html_body = f"""
    <html>
    <body>
        <h2>Embedded Image Example</h2>
        <!-- This cid: must match the Content-ID we set above -->
        <img src="cid:{our_content_id}" alt="Logo" width="150">
        <p>The image above is embedded, not attached!</p>
    </body>
    </html>
    """
    
    mail.HTMLBody = html_body
    
    # Display instead of sending for testing
    mail.Display()  # Use mail.Send() when ready

detailed_embed_example()
