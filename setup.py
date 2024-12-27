from setuptools import setup, find_packages

CURRENT_VERSION = '1.950'

setup(
    name='PyWrike',
    version=CURRENT_VERSION,
    packages=['PyWrike', 'PyWrike.gateways'],
    description='A Python package to make API calls to Wrike',
    author='Ambareen Jawaheer',
    author_email='ambareenjawaheer@gmail.com',
    url='https://github.com/axirestech/PyWrike',
    download_url='https://github.com/axirestech/PyWrike/archive/refs/tags/v%s.tar.gz' % CURRENT_VERSION,
    keywords=['api', 'gateway', 'wrike', 'http', 'REST'],
    install_requires=[
        'requests>=2.0',
        'openpyxl>=3.0',
        'beautifulsoup4>=4.9',
        'pandas>=1.0',
        'Flask>=2.0'
        #'basegateway>=0,<1'
    ],
    #use_2to3=True,
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
)
