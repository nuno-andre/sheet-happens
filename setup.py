from setuptools import setup, find_packages
from os import path

import sheet_happens

rootdir = path.abspath(path.dirname(__file__))

with open(path.join(rootdir, 'README.md'), encoding='utf-8') as f:
    long_description = f.read()

setup(
    name='sheet_happens',
    version=sheet_happens.__version__,
    description=sheet_happens.__description__,
    long_description=long_description,
    long_description_content_type='text/markdown', 
    author='Nuno André',
    author_email='mail@nunoand.re',
    url='https://github.com/nuno-andre/sheet-happens',
    classifiers=[
        'Development Status :: 5 - Production',
        'Intended Audience :: Developers',
        'License :: MIT',
        'Topic :: Text Processing :: Markup',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.4',
        'Programming Language :: Python :: 3.5',
        'Programming Language :: Python :: 3.6',
    ],
    keywords='excel json csv yaml',
    py_modules=['sheet_happens'],
    python_requires='>=3.4, <4',
    entry_points={
        'console_scripts': [
            'sheet-happens=sheet_happens:main',
        ],
    },
    project_urls={
        'Bug Reports': 'https://github.com/nuno-andre/sheet-happens/issues',
        'Source': 'https://github.com/nuno-andre/sheet-happens',
    },
)