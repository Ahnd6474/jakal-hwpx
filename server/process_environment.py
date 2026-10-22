"""Do not pass hosting/API credentials to document parsers or converters."""
import os
from pathlib import Path


def parser_environment(repo: Path) -> dict[str, str]:
    allowed = {'PATH', 'SYSTEMROOT', 'WINDIR', 'COMSPEC', 'PATHEXT', 'TEMP', 'TMP', 'TMPDIR',
        'HOME', 'USERPROFILE', 'APPDATA', 'LOCALAPPDATA', 'LANG', 'LANGUAGE', 'LC_ALL',
        'LD_LIBRARY_PATH', 'DYLD_LIBRARY_PATH', 'VIRTUAL_ENV', 'JAKAL_NODE'}
    env = {key: value for key, value in os.environ.items() if key.upper() in allowed}
    env.update(PYTHONPATH=str(repo / 'src') + os.pathsep + str(repo), PYTHONIOENCODING='utf-8',
        PYTHONNOUSERSITE='1')
    return env
