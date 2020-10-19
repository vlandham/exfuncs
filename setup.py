import setuptools

version = "0.0.3"

with open("README.md") as readme_file:
    readme = readme_file.read()

install_requirements = [
    "openpyxl==3.0.5",
]

setuptools.setup(
    name="exfuncs",
    description="",
    long_description=readme,
    version=version,
    packages=setuptools.find_packages(exclude=["tests"]),
    install_requires=install_requirements,
    include_package_data=True,
)
