import asyncio
import aiohttp
from word_document_server.tools.document_tools import load_example_document

async def test_end_to_end():
    # Step 1: Load the example document
    print("🔍 Loading example document...")
    result = await load_example_document()
    
    if 'error' in result:
        print(f"❌ Error loading document: {result['error']}")
        return
    
    print("✅ Document loaded successfully!")
    print(f"   - File ID: {result['file_id']}")
    print(f"   - Filename: {result['filename']}")
    print(f"   - Download URL: {result['download_url']}")
    
    # Step 2: Try to download the document
    print("\n⬇️  Attempting to download the document...")
    try:
        async with aiohttp.ClientSession() as session:
            async with session.get(result['download_url']) as response:
                if response.status == 200:
                    content = await response.read()
                    with open('downloaded_document.docx', 'wb') as f:
                        f.write(content)
                    print(f"✅ Document downloaded successfully!")
                    print(f"   - Saved as: downloaded_document.docx")
                    print(f"   - Size: {len(content)} bytes")
                else:
                    print(f"❌ Failed to download document. Status: {response.status}")
                    print(f"   - Response: {await response.text()}")
    except Exception as e:
        print(f"❌ Error downloading document: {str(e)}")

if __name__ == "__main__":
    asyncio.run(test_end_to_end())
