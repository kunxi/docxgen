#!/usr/bin/env python

from setuptools import setup, find_packages

version = '0.0.1'

setup(
    name='docxgen',
    version=version,
    packages=find_packages(),
    install_requires=['lxml'],

    description='A simple module to read and write Microsoft Office Word 2007 docx documents.',
    author='Kun Xi',
    author_email='kunxi@kunxi.org',
    url='http://github.com/kunxi/docxgen',
    license = 'MIT',
)
