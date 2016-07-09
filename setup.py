#!/usr/bin/env python

"""Installer for rst2pptx."""

import ast
import io

import setuptools


def version():
    """Return version string."""
    with io.open('rst2pptx.py', encoding='utf-8') as input_file:
        for line in input_file:
            if line.startswith('__version__'):
                return ast.parse(line).body[0].value.s

with io.open('README.rst', encoding='utf-8') as readme:
    setuptools.setup(
        name='rst2pptx',
        version=version(),
        description='A docutils writer and script for converting '
                    'reStructuredText to the PowerPoint format',
        long_description=readme.read(),
        classifiers=[
            'Development Status :: 2 - Pre-Alpha',
            'License :: OSI Approved :: MIT License',
            'Programming Language :: Python',
            'Programming Language :: Python :: 3',
            'Topic :: Text Processing :: Markup',
            'Topic :: Utilities',
            'Topic :: Multimedia :: Graphics :: Presentation',
        ],
        keywords='presentation,docutils,rst,restructuredtext,powerpoint,pptx',
        url='https://github.com/myint/rst2pptx',
        py_modules=['rst2pptx'],
        zip_safe=False,
        install_requires=[
            'docutils >= 0.11',
            'python-pptx >= 0.5.8',
        ],
        entry_points={
            'console_scripts': [
                'rst2pptx = rst2pptx:main',
            ]
        }
    )
