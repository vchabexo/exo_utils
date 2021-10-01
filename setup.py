from setuptools import setup, find_packages

setup(
    name='exo_utils', 
    version='1.0', 
    packages=find_packages(),
    install_requires =[
        'pandas',
        'unidecode',
        'selenium',
        'pyodbc',
        'numpy',
        'xlrd'
    ]
    )