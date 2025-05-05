import os
import sys

def resource_path(relative_path):
    """Retorna o caminho absoluto, mesmo se for .exe"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)