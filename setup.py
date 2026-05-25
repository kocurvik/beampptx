from setuptools import setup

setup(
    name="beampptx",
    version="0.1.0",
    description="Convert LaTeX Beamer slides to PowerPoint with vector graphics and video support.",
    author="Viktor Kocur",
    author_email="kocurvik@gmail.com",
    py_modules=["generate_pptx"],
    install_requires=[
        "python-pptx",
        "pymupdf",
        "lxml",
    ],
    entry_points={
        "console_scripts": [
            "beampptx = generate_pptx:main",
        ],
    },
    python_requires=">=3.6",
)
