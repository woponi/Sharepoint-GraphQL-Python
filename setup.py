from setuptools import setup, find_packages

setup(
    name='sharepoint_graphql',
    version='0.16',
    description='This Python utility enables users to interact with SharePoint sites via Microsoft Graph API, facilitating tasks such as listing, downloading, uploading, moving, and deleting files.',
    long_description=open('README.md').read(),
    long_description_content_type='text/markdown',
    author='Pong Wong',
    author_email='ninn@opts-db.com',
    url='https://github.com/woponi/Sharepoint-GraphQL-Python',
    packages=find_packages(),
    install_requires=[
        'requests',
        'msal',
        # Add any other dependencies here
    ],
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
    ],
)
