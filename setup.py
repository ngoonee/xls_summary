import setuptools

setuptools.setup(
    name="xls_summary",
    version="0.1.3",
    url="https://github.com/ngoonee/xls_summary",

    author="Ng Oon-Ee",
    author_email="ngoe@utar.edu.my",

    description="Summarizes multiple copies of the same excel spreadsheet",
    long_description=open('README.rst').read(),

    packages=setuptools.find_packages(),

    install_requires=[
        'openpyxl>=2.4.7',
    ],

    classifiers=[
        'Development Status :: 2 - Pre-Alpha',
        'Programming Language :: Python',
        'Programming Language :: Python :: 2',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.4',
        'Programming Language :: Python :: 3.5',
    ],
)
