import logging
import traceback
import os
import shutil

def flushFolderContent(folder):
    """
    flush all the content of the folder exepte the .keep file
    folder : str absolute path to the folder
    """

    for filename in os.listdir(folder):
        if filename != '.keep':
            file_path = os.path.join(folder, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                logging.error('Failed to delete %s. Reason: %s' % (file_path, traceback.format_exc()))
    return