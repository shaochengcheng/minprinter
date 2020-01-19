#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""This file is used by Pyinstaller to make windows EXE program.
"""

import os
import sys

from minprinter.frontend import add_poppler_to_os_path, gui_main

if getattr(sys, 'frozen', False):
    # we are running in the Pyinstaller bundle
    bundle_dir = sys._MEIPASS
else:
    bundle_dir = os.path.dirname(os.path.abspath(__file__))

poppler_bin = os.path.join(bundle_dir, 'poppler/bin')
add_poppler_to_os_path(poppler_bin)

gui_main()
