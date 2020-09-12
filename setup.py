from setuptools import find_packages, setup

setup(
    name="excel2django",
    version="0.1",
    description="Commandline tool to move excel data into django objects",
    long_description="",
    url="https://basx.dev",
    author="basx Software Development Co., Ltd.",
    author_email="info@basx.dev",
    license="Private",
    install_requires=["django"],
    setup_requires=["setuptools_scm"],
    packages=find_packages(),
    use_scm_version=True,
    zip_safe=False,
)
