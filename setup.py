#!/usr/bin/env python

from setuptools import setup, find_packages

version = '0.1.3'

setup(
    name='docxgen',
    version=version,
    packages=find_packages(),
    install_requires=['lxml', 'six'],
    include_package_data=True,
    test_suite='nose.collector',
    tests_require=['nose', 'coverage'],

    description='A library to generate Microsoft Office Word 2007 documents.',
    author='Kun Xi',
    author_email='kunxi@kunxi.org',
    url='http://github.com/kunxi/docxgen',
    license='MIT',
)
