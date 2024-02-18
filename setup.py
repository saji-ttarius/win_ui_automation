from setuptools import find_packages, setup

setup(
    name="win_ui_automation",
    version="0.1.0",
    packages=find_packages(),
    install_requires=[
        "comtypes",
    ],
    entry_points={
        "console_scripts": [
            "win_ui_automation_script=win_ui_automation:main",
        ],
    },
    author="s@ji",
    author_email="marthijo@live.jp",
    description="Win UI Automation Utility",
    license="MIT",
    keywords="uiautomation",
    url="https://github.com/akiranger/win_ui_automation",
)
