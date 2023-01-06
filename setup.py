import setuptools

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setuptools.setup(
    name='emails',
    version='0.1',
    author='Mike Smith',
    author_email='mikesmithlab@gmail.com',
    description='Testing installation of Package',
    long_description=long_description,
    long_description_content_type="text/markdown",
    url='https://github.com/mikesmithlab/pydates',
    license='MIT',
    packages=setuptools.find_packages(),
    install_requires=['pywin32',
                      'pydates @ git+https://github.com/mikesmithlab/pydates', 'datetime'],
)
