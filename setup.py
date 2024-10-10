from setuptools import setup
import platform


if platform.system() == 'Windows':
    ext_files = ['*.dll']
elif platform.system() == 'Linux':
    ext_files = ['*.so']
elif platform.system() == 'Darwin':
    ext_files = ['*.dylib']
else:
    ext_files = []

setup(include_package_data=True, package_data={'pyfastexcel': ext_files})
