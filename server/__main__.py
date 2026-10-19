import os

import uvicorn

if __name__ == "__main__":
    uvicorn.run("server.app:create_app", factory=True, host=os.environ.get("JAKAL_HOST", "127.0.0.1"), port=int(os.environ.get("JAKAL_PORT", "8000")), workers=1)
