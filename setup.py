from setuptools import setup, find_packages


setup(name='doccleaner-plugins',
      version='0.2.0',
      description='Plugins to connect the doccleaner module to text processing softwares (MS Word, LibreOffice Writer)',
      url='',
      download_url='',
      author='Jean-Baptiste Bertrand',
      author_email='jean-baptiste.bertrand@openedition.org',
      license='LGPL 3.0',
	  include_package_data=True,
      packages=find_packages(),
      install_requires=['defusedxml', 'doccleaner'],	
      package_data = {
      '': ['*.*'],
      },
      zip_safe=False)
      

