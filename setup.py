#!/usr/bin/env python
# -*- coding: utf-8 -*-

from os.path import dirname, join
from setuptools import setup, find_packages

with open(join(dirname(__file__), 'minprinter/VERSION'), 'rb') as f:
    version = f.read().decode('ascii').strip()

setup(
    name='minprinter',
    version=version,
    description='A printer to print signed mobile phone invoice PDFs',
    author='Chengcheng Shao',
    maintainer='Chengcheng Shao',
    maintainer_email='shaoc@indiana.edu',
    license='GPLv3',
    packages=find_packages(exclude=('tests', 'tests.*')),
    include_package_data=True,
    zip_safe=True,
    entry_points={
        'console_scripts': ['minprint = minprinter.frontend:gui_main']
    },
    classifiers=[
        'Development Status :: 5 - Production/Stable',
        'Environment :: GUI',
        'Intended Audience :: Science/Research',
        'License :: GPL :: Version 3',
        'Operating System :: POSIX :: Windows',
        'Programming Language :: Python',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.7',
        'Topic :: Software Development :: Libraries :: Python Modules',
    ],
    install_requires=[
        'appJar>=0.94.0', 'pandas>=0.25.3', 'openpyxl>=3.0.2',
        'pdf2image>=1.11.0'
    ],
)
