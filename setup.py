from setuptools import setup, find_packages

setup(
    name='robotframework-pywinautolibrary',
    version='1.0.0',
    description='Robot Framework library for Pywinauto.',
    author='Anoop G R',
    author_email='agrtest001@gmail.com',
    url='https://github.com/AnoopGR/robotframework-pywinautolibrary',
    packages=find_packages(),
    install_requires=[
        'robotframework',
        'pywinauto'
    ],
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: MIT License',
        'Operating System :: Microsoft :: Windows',
    ],
)
