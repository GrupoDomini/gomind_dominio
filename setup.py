from setuptools import setup

setup(
    name="gomind_dominio",
    python_requires=">=3.6",
    version="0.0.1",
    description="GoMind dominio service",
    url="https://github.com/GrupoDomini/gomind_dominio.git",
    author="JeffersonCarvalhoGM",
    author_email="jefferson.carvalho@grupodomini.com",
    license="unlicense",
    packages=["gomind_dominio"],
    zip_safe=False,
    install_requires=[
        "gomind_web_browser @ git+https://github.com/GrupoDomini/gomind_web_browser.git",
        "gomind_excel @ git+https://github.com/GrupoDomini/gomind_excel.git",
        "gomind_automation @ git+https://github.com/GrupoDomini/gomind_automation.git",
    ],
)
