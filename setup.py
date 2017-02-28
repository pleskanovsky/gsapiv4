from setuptools import setup

setup(name='gsapiv4',
      version='0.0.0',
      description='Google Spreadsheets API V4 wrapper',
      url='https://github.com/pleskanovsky/gsapiv4',
      author='pleskanovsky',
      author_email='vnezapno.pochta@gmail.com',
      license='MIT',
      packages=['gsapiv4'],
      install_requires=[
          "google-api-python-client",
      ],
      zip_safe=False)