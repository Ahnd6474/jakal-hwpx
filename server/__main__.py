import uvicorn

from .runtime import check_runtime, server_options

if __name__ == "__main__":
    options = server_options()
    check_runtime()
    uvicorn.run("server.app:create_app", factory=True, **options)
