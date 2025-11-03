import asyncio
from word_document_server.tools.document_tools import load_example_document

async def test():
    result = await load_example_document()
    print("Test Result:")
    print(result)

if __name__ == "__main__":
    asyncio.run(test())
