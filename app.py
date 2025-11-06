import streamlit as st
import pandas as pd
import os
import time
import sys
import subprocess
import pyautogui
import logging
from pathlib import Path

# Attempt to install missing dependencies
def install_dependencies():
    dependencies = [
        'pywhatkit',
        'pyautogui',
        'python-xlib',  # for Linux/Mac
        'pywin32',      # for Windows clipboard support
        'openpyxl>=3.1.0',  # Compatible with pandas and fixes extLst error
        'xlrd>=2.0.1'        # Add xlrd for .xls file support
    ]
    
    for dep in dependencies:
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", dep])
            st.info(f"Installed {dep} successfully")
        except Exception as e:
            st.error(f"Could not install {dep}: {e}")

# Check and install dependencies
try:
    import pywhatkit
    import pyautogui
except ImportError:
    st.warning("Missing dependencies. Attempting to install...")
    install_dependencies()
    
    # Retry imports
    try:
        import pywhatkit
        import pyautogui
    except ImportError:
        st.error("Failed to install required dependencies. Please install manually.")
        st.stop()

def send_whatsapp_message(phone_number, message, media_path=None, media_type='image'):
    """
    Send WhatsApp message with optional media (image/video) using PyWhatKit

    Args:
        phone_number: Recipient's phone number
        message: Message text or caption
        media_path: Path to media file (image or video)
        media_type: Type of media ('image' or 'video')

    Returns:
        tuple: (success: bool, message: str)
    """
    try:
        # Clean phone number (remove non-digit characters)
        cleaned_number = ''.join(filter(str.isdigit, str(phone_number)))

        # Ensure the number is in the correct format (with country code)
        # Remove leading 91 if already present to avoid duplication
        if cleaned_number.startswith('91') and len(cleaned_number) > 10:
            cleaned_number = cleaned_number[2:]

        formatted_number = f"+91{cleaned_number}"

        # Validate media file if provided
        if media_path:
            if not os.path.exists(media_path):
                return False, f"Media file not found: {media_path}"

            # Get absolute path to ensure it's accessible
            media_path = os.path.abspath(media_path)
            file_size = os.path.getsize(media_path)

            # Check file size (WhatsApp has limits)
            max_size_mb = 100 if media_type == 'video' else 16
            if file_size > max_size_mb * 1024 * 1024:
                return False, f"Media file too large. Max size: {max_size_mb}MB"

            # Log media details
            st.write(f"📎 Attaching {media_type}: {os.path.basename(media_path)} ({file_size / 1024 / 1024:.2f}MB)")

        # Send message with media
        if media_path and os.path.exists(media_path):
            try:
                if media_type == 'image':
                    # Send message with image
                    pywhatkit.sendwhats_image(
                        receiver=formatted_number,
                        img_path=media_path,
                        caption=message,
                        wait_time=15  # Increased wait time for image processing
                    )
                    return True, f"Message with image sent to {phone_number}"

                elif media_type == 'video':
                    # Send message with video
                    pywhatkit.sendwhats_video(
                        receiver=formatted_number,
                        video_path=media_path,
                        caption=message,
                        wait_time=15  # Increased wait time for video processing
                    )
                    return True, f"Message with video sent to {phone_number}"

            except AttributeError as ae:
                # If sendwhats_video doesn't exist, fall back to image method
                if 'sendwhats_video' in str(ae):
                    st.warning(f"Video sending not supported in this pywhatkit version. Sending as image instead.")
                    pywhatkit.sendwhats_image(
                        receiver=formatted_number,
                        img_path=media_path,
                        caption=message,
                        wait_time=15
                    )
                    return True, f"Message with media sent to {phone_number}"
                raise

        else:
            # Send text message only
            pywhatkit.sendwhatmsg_instantly(
                phone_no=formatted_number,
                message=message,
                wait_time=10
            )
            return True, f"Text message sent to {phone_number}"

    except Exception as e:
        error_msg = f"Error sending message to {phone_number}: {str(e)}"
        return False, error_msg

def main():
    st.title("WhatsApp Bulk Message Sender (PyWhatKit)")

    # Sidebar for configuration
    st.sidebar.header("Configuration")

    # Excel/CSV file upload
    uploaded_file = st.sidebar.file_uploader(
        "Upload Excel/CSV File", 
        type=['xlsx', 'xls', 'csv'],
        help="File must contain 'Name' and 'Phone Number' columns"
    )

    # Media file upload (images and videos)
    uploaded_media = st.sidebar.file_uploader(
        "Upload Media (Optional)",
        type=['png', 'jpg', 'jpeg', 'gif', 'mp4', 'avi', 'mov', 'mkv'],
        help="Image or video to send along with the message"
    )

    # Message input
    message_template = st.sidebar.text_area(
        "Message Template", 
        "Hello {{Name}}, this is a test message!",
        help="Use {{Name}} for personalization"
    )

    # Send button
    if st.sidebar.button("Send Messages"):
        # Validate file upload
        if uploaded_file is not None:
            try:
                # Save uploaded media temporarily if provided
                media_path = None
                media_type = 'image'

                if uploaded_media:
                    # Create temp directory if it doesn't exist
                    os.makedirs('temp', exist_ok=True)
                    media_path = os.path.join('temp', uploaded_media.name)

                    # Determine media type based on file extension
                    file_ext = uploaded_media.name.lower().split('.')[-1]
                    if file_ext in ['mp4', 'avi', 'mov', 'mkv']:
                        media_type = 'video'
                    else:
                        media_type = 'image'

                    # Save media file
                    with open(media_path, 'wb') as f:
                        f.write(uploaded_media.getbuffer())

                    st.success(f"✅ Media file uploaded: {uploaded_media.name} ({media_type})")

                # Read file based on extension with improved error handling
                try:
                    # Try different engines based on file extension
                    file_extension = uploaded_file.name.lower().split('.')[-1]
                    
                    if file_extension == 'csv':
                        df = pd.read_csv(uploaded_file)
                    elif file_extension == 'xlsx':
                        df = pd.read_excel(uploaded_file, engine='openpyxl')
                    elif file_extension == 'xls':
                        df = pd.read_excel(uploaded_file, engine='xlrd')
                    else:
                        # Default to CSV then openpyxl
                        try:
                            df = pd.read_csv(uploaded_file)
                        except:
                            df = pd.read_excel(uploaded_file, engine='openpyxl')
                        
                except Exception as excel_error:
                    st.error(f"Error reading Excel file: {excel_error}")

                    # Show specific error messages and solutions
                    if "extLst" in str(excel_error):
                        st.error("⚠️ openpyxl version compatibility issue detected!")
                        st.info("💡 Solution: Run this command in terminal: pip install openpyxl>=3.1.0")
                    elif "xlrd" in str(excel_error) or "Missing optional dependency 'xlrd'" in str(excel_error):
                        st.error("⚠️ Missing xlrd dependency for .xls files!")
                        st.info("💡 Solution: Run this command in terminal: pip install xlrd>=2.0.1")

                    # Try alternative approach
                    try:
                        st.info("Attempting alternative file reading method...")
                        if file_extension == 'xls':
                            df = pd.read_excel(uploaded_file, engine='xlrd')
                        else:
                            df = pd.read_excel(uploaded_file, engine='openpyxl')
                    except Exception as fallback_error:
                        st.error(f"All file reading methods failed: {fallback_error}")
                        st.error("Please ensure you have the correct dependencies installed:")
                        st.code("pip install -r requirements.txt")
                        return

                # Validate columns
                required_columns = ['Name', 'Phone Number']
                if not all(col in df.columns for col in required_columns):
                    st.error("Excel sheet must contain 'Name' and 'Phone Number' columns.")
                    return

                # Print dataframe for debugging
                st.write("Contacts DataFrame:")
                st.write(df)

                # Progress tracking
                progress_bar = st.progress(0)
                status_text = st.empty()

                # Flag to check if WhatsApp Web is already open
                whatsapp_open = False

                # Send messages
                total_contacts = len(df)
                for idx, row in df.iterrows():
                    try:
                        # Personalize message
                        personalized_message = message_template.replace("{{Name}}", str(row['Name']))
                        
                        # Open WhatsApp Web only if it's not already open
                        if not whatsapp_open:
                            # Open WhatsApp Web in the first tab
                            pyautogui.hotkey('ctrl', 't')  # Open new tab
                            pyautogui.typewrite("https://web.whatsapp.com")  # Navigate to WhatsApp Web
                            pyautogui.press('enter')
                            time.sleep(20)  # Wait for WhatsApp Web to load
                            whatsapp_open = True

                        # Send message
                        success, send_status = send_whatsapp_message(
                            str(row['Phone Number']),
                            personalized_message,
                            media_path,
                            media_type
                        )

                        # Display status
                        if success:
                            st.success(f"✅ {send_status}")
                        else:
                            st.error(f"❌ {send_status}")

                        # Update progress
                        progress = (idx + 1) / total_contacts
                        progress_bar.progress(progress)
                        status_text.text(f"Sending message {idx + 1}/{total_contacts}")

                        # Wait between messages to avoid rate limiting
                        time.sleep(10)  # Adjusted sleep time for smoother operation

                        # Close tab after each message is sent
                        if success:
                            pyautogui.hotkey('ctrl', 'w')  # Close current tab
                            time.sleep(5)  # Allow time for tab to close
                            # Keep WhatsApp Web open in the same tab for next message

                    except Exception as inner_e:
                        st.error(f"Error processing contact {row['Name']}: {inner_e}")

                # Final status
                st.success("Message sending completed!")

            except Exception as e:
                st.error(f"An error occurred: {e}")
            
            finally:
                # Remove temporary media file
                if 'media_path' in locals() and media_path and os.path.exists(media_path):
                    try:
                        os.remove(media_path)
                        st.info(f"Cleaned up temporary media file: {media_path}")
                    except Exception as cleanup_error:
                        st.warning(f"Could not delete temporary file: {cleanup_error}")

        else:
            st.error("Please upload an Excel file before sending messages.")

    # Additional instructions
    st.sidebar.info("""
    ## 📋 Instructions
    1. **Prepare an Excel file** with columns:
       - Name
       - Phone Number
    2. **Optional: Upload Media** (Image or Video)
       - Supported formats: PNG, JPG, JPEG, GIF, MP4, AVI, MOV, MKV
       - Max size: 16MB for images, 100MB for videos
    3. **Enter your message template**
       - Use {{Name}} for personalization
    4. **Click 'Send Messages'**
    5. **IMPORTANT: Keep WhatsApp Web open** in your default browser

    ## ⚠️ Important Notes
    - Messages will be sent automatically using PyWhatKit
    - Ensure phone numbers are for Indian numbers (starting with +91)
    - Media files are temporarily stored and deleted after sending
    - Each message has a 10-second delay to avoid rate limiting
    """)

if __name__ == "__main__":
    main()
