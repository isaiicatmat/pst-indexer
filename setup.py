from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="pst-indexer",
    version="1.0.0",
    author="Isaí Carrero Martínez",
    author_email="isaiicatmat@gmail.com",
    description="A fast and easy-to-use Outlook PST file email indexer and searcher",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/isaiicatmat/pst-indexer",
    packages=find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "License :: OSI Approved :: MIT License",
        "Operating System :: Microsoft :: Windows",
        "Operating System :: MacOS",
        "Operating System :: POSIX :: Linux",
        "Topic :: Communications :: Email",
        "Topic :: Office/Business",
        "Topic :: Utilities",
    ],
    python_requires=">=3.7",
    install_requires=[
        "libpst-python>=0.1.3",
        "extract-msg>=0.47.1",
        "olefile>=0.47",
    ],
    entry_points={
        "console_scripts": [
            "pst-indexer=pst_indexer.gui:main",
        ],
    },
)
