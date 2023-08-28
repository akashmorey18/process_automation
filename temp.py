import streamlit as st
from PIL import Image
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file_creation_information import FileCreationInformation

# SharePoint authentication using client ID and client secret

site_url = "https://tatadigitalltd.sharepoint.com/sites/AutomationProject"
client_id = '96026163-d873-4cca-abd4-7acf4fa3e1a3'
client_secret = 'DolnVP4E9IIoLNf/86GHbABu+djbFE679RerpV5Z9Hs='


ctx_auth = AuthenticationContext(url=site_url)
ctx_auth.acquire_token_for_app(client_id, client_secret)
ctx = ClientContext(site_url, ctx_auth)

def upload_to_sharepoint(file_content, file_name, folder_url):
    file_info = FileCreationInformation()
    file_info.content = file_content
    file_info.url = file_name
    target_folder = ctx.web.get_folder_by_server_relative_url(folder_url)
    uploaded_file = target_folder.files.add(file_info)
    ctx.execute_query()

def main():
    st.title("Image Upload to SharePoint")
    
    uploaded_image = st.file_uploader("Choose an image", type=["jpg", "png", "jpeg"])
    
    if uploaded_image is not None:
        image = Image.open(uploaded_image)
        st.image(image, caption="Uploaded Image", use_column_width=True)

        if st.button("Upload to SharePoint"):
            image_content = uploaded_image.read()
            folder_url = "/sites/AutomationProject/Shared Documents/Fashion Merchandising"
            upload_to_sharepoint(image_content, uploaded_image.name, folder_url)
            st.success("Image uploaded to SharePoint!")

if __name__ == "__main__":
    main()
