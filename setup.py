from distutils.core import setup

CURRENT_VERSION = '0.7'

setup(
  name='PyWrike',
  version=CURRENT_VERSION,
  py_modules=['wrike'],
  description='A class to make API calls to Wrike',
  author='ambrine-jawahir',
  author_email='ambrine@axires.tech',
  url='https://github.com/axirestech/PyWrike',
  download_url='https://github.com/axirestech/PyWrike' % CURRENT_VERSION,
  keywords=['api', 'gateway', 'http', 'REST'],
  install_requires=[
    'basegateway>=0,<1',
    'requests>=2.0',
    'openpyxl>=3.0',
    'beautifulsoup4>=4.9',
    'pandas>=1.0'
  ],
  classifiers=[
    "Topic :: Internet :: WWW/HTTP",
    "Development Status :: 4 - Beta",
    "Intended Audience :: Developers",
    "Programming Language :: Python",
    "Programming Language :: Python :: 3",
    "License :: OSI Approved :: MIT License",
    "Operating System :: OS Independent",
  ],
)
