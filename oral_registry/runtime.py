from pathlib import Path
import shutil
from PySide6.QtCore import QStandardPaths

RESOURCE_ROOT=Path(__file__).resolve().parent.parent

def initial_database(directory=None):
    directory=Path(directory) if directory else Path(QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DocumentsLocation))/'ChildRegistry'
    directory.mkdir(parents=True,exist_ok=True)
    target=directory/'studentRegistry.xlsx'
    if not target.exists():
        # Exclusive creation protects existing data, including two simultaneous first launches.
        try:
            with target.open('xb') as output, (RESOURCE_ROOT/'templateRegistry.xlsx').open('rb') as source:
                shutil.copyfileobj(source,output)
        except FileExistsError:
            pass
    return target
