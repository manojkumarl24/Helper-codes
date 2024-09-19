import streamlit as st
import os
import time
from sider import open_sidebar
import win32com.client  # Assuming you're using win32com.client to handle PST parsing

# Function to parse PST file
def parse_pst_file(pst_path):
    st.write("Parsing the PST file...")  # Inform the user parsing has started
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    
    # Simulate long PST parsing (Replace this with your actual parsing logic)
    try:
        # Parsing logic
        outlook_stores = outlook.Folders
        for store in outlook_stores:
            print(store.Name)
            time.sleep(0.5)  # Simulating processing
        st.session_state['parsed'] = True  # Mark parsing as complete
    except Exception as e:
        st.error(f"Failed to parse PST file: {e}")
        st.session_state['parsed'] = False

# Main interface
st.title("PST File Parser")

pst_file_path = st.text_input("Enter the path of the PST file:")

# Button to parse PST file
if st.button("Parse PST File"):
    if pst_file_path:
        st.session_state['parsed'] = False  # Reset the parsed state
        parse_pst_file(pst_file_path)
    else:
        st.error("Please provide a valid PST file path.")

# Display success message only after parsing is done
if 'parsed' in st.session_state and st.session_state['parsed']:
    st.success("PST file parsed successfully!")
    open_sidebar()  # Call sidebar only after parsing is complete

#FILE2- SIDEBAR.PY
import streamlit as st
import os

def open_sidebar():
    # Sidebar title
    with st.sidebar:
        st.header("Output Settings")

        # Get output folder path from user
        output_folder = st.text_input("Enter the output folder path:")

        # Check if it's a valid folder path
        if output_folder and os.path.isdir(output_folder):
            st.success("Valid output folder path")
            st.session_state['valid_folder'] = True  # Mark folder path as valid
        else:
            st.error("Please provide a valid folder path.")
            st.session_state['valid_folder'] = False

        # Only show filter button if the folder path is valid
        if 'valid_folder' in st.session_state and st.session_state['valid_folder']:
            if st.button("Show Filter Options"):
                # Display filter options
                st.write("Here are the filter options.")


