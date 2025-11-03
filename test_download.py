import requests

def test_download():
    url = "http://127.0.0.1:8001/mcp/download/f425223c-df2a-463c-beea-d753bf3ae895"
    
    try:
        response = requests.get(url, stream=True)
        response.raise_for_status()  # Raise an exception for HTTP errors
        
        # Check if the response is a DOCX file
        content_type = response.headers.get('content-type')
        if 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' in content_type:
            # Save the file locally to verify
            with open('downloaded_document.docx', 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            print("✅ Document downloaded successfully as 'downloaded_document.docx'")
            print(f"   - Content-Type: {content_type}")
            print(f"   - File size: {len(response.content)} bytes")
            return True
        else:
            print(f"❌ Unexpected content type: {content_type}")
            print(f"Response content: {response.text[:200]}...")  # Show first 200 chars of response
            return False
            
    except requests.exceptions.RequestException as e:
        print(f"❌ Error downloading document: {e}")
        return False

if __name__ == "__main__":
    test_download()
