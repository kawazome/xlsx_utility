from setuptools import setup, find_packages

setup(
    name='xlsx_utility',
    version="1.0.0",
    description="Utility class to easily manipulate Excel files using openpyxl.",
    long_description="",
    author='kawazome',
    packages=find_packages(),
    install_requires=[
        'openpyxl'
    ],
    dependency_links=[
        'https://github.com/kawazome/debug'
    ]
)