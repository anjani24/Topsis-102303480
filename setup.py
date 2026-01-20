from setuptools import setup, find_packages

setup(
    name="Topsis-Anjani-102303480",
    version="1.0.1",
    packages=find_packages(),

    description="TOPSIS decision making tool",
    long_description=open("README.md").read(),
    long_description_content_type="text/markdown",

    author="Anjani Agarwal",

    install_requires=[
        "numpy",
        "pandas"
    ],

    entry_points={
        "console_scripts": [
            "topsis = topsis_anjani_102303480.topsis:main"
        ]
    }
)
