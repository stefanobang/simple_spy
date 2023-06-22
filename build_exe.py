from distutils.core import setup
import py2exe

import pyscreenshot as ImageGrab 
from datetime import datetime
from multiprocessing import Process, freeze_support

setup(console=['autoscreenshot.py'])