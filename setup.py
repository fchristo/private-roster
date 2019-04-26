from setuptools import setup, find_packages
from os import path
from io import open

here = path.abspath(path.dirname(__file__))

# Get the long description from the README file
with open(path.join(here, 'README.md'), encoding='utf-8') as f:
    long_description = f.read()

# For information on setup.py, see https://github.com/pypa/sampleproject/blob/master/setup.py

setup(
    name='roster',
    version='0.0.1',
    description='A roster viewer for poor Mrs.Jones',
    url='https://github.com/fchristo/private-roster',
    author='Chris Frambers',
    author_email='kjhvgfl@gmail.com',
    classifiers=[
        'Development Status :: 3 - Alpha',
        'Intended Audience :: Private',
        'Topic :: Software Development',
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.4',
        'Programming Language :: Python :: 3.5',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
    ],
    packages=find_packages(exclude=['docs', 'tests']),
    python_requires='>3.5',
    install_requires=[
        'black==19.3b0',
        'docopt==0.6.2',
        'invoke==1.2.0',
        'openpyxl==2.6.2',
        'pandas==0.24.2',
        'pip-tools==3.6.1',
        'Sphinx==2.0.1',
    ],
    package_data={
        'student_data': ['Jones_2019.xlsx'],
    },
    entry_points={
        'console_scripts': [
            'roster=roster:main',
        ],
    },
)
