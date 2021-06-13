import setuptools

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setuptools.setup(
    name="xlsx-extract",
    version="0.1.0",
    author="Martin Aspeli",
    author_email="optilude@gmail.com",
    description="Tools to extract data from (poorly) structured Excel files, building on openpyxl",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/optilude/xlsx-extract",
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    package_dir={"": "src"},
    packages=setuptools.find_packages(where="src"),
    python_requires=">=3.6",
    install_requires=[
        "openpyxl",
    ],
    entry_points = {
        'console_scripts': ['xlsx-extract=xlsx_extract.cli:main'],
    }
)