# -*- coding:utf-8 -*-
from setuptools import setup, find_packages
from DrissionRecorder import __version__

with open("README.md", "r", encoding='utf-8') as fh:
    long_description = fh.read()

setup(
    name="DrissionRecorder",
    version=__version__,
    author="g1879",
    author_email="g1879@qq.com",
    description="用于记录数据的模块。",
    long_description=long_description,
    long_description_content_type="text/markdown",
    license="MIT",
    keywords="DrissionRecorder",
    url="https://gitee.com/g1879/DrissionRecorder",
    include_package_data=True,
    packages=find_packages(),
    install_requires=[
        "openpyxl",
    ],
    classifiers=[
        "Programming Language :: Python :: 3.6",
        "Development Status :: 4 - Beta",
        "Topic :: Utilities",
        "License :: OSI Approved :: BSD License",
    ],
    python_requires='>=3.6'
)
