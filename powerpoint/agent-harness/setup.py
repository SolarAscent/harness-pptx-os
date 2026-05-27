from setuptools import find_namespace_packages, setup

setup(
    name="cli-anything-powerpoint",
    version="0.1.0",
    description="CLI-Anything harness for Microsoft PowerPoint.",
    packages=find_namespace_packages(include=["cli_anything.*"]),
    install_requires=["click>=8.0"],
    entry_points={
        "console_scripts": [
            "cli-anything-powerpoint=cli_anything.powerpoint.powerpoint_cli:main",
        ],
    },
)
