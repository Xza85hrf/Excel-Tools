from setuptools import setup, find_packages

setup(
    name='excel-comparison-app',
    version='0.1.0',
    author='Your Name',
    author_email='',
    description='Excel Comparison App with GUI and reporting features',
    packages=find_packages(),
    install_requires=[
        'openpyxl>=3.0.0',
    ],
    tests_require=[
        'pytest>=6.0.0',
    ],
    python_requires='>=3.6',
)
