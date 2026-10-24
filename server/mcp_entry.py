"""Absolute-path entry point usable from any MCP client's working directory."""
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from server.mcp_bridge import main

if __name__ == '__main__':
    main()
