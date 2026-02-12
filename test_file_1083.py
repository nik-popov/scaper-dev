import asyncio
from main import generate_download_file

async def test():
    print("Testing FileID 1083 with debug raises enabled...")
    try:
        await generate_download_file("1083", row_offset=0)
        print("✓ Job completed successfully!")
    except Exception as e:
        print(f"✗ Exception caught: {type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    asyncio.run(test())
